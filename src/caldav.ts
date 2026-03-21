import type { DiscoveredTaskList, LogseqTaskItem, NextcloudSyncSettings, TaskScopeConfig } from "./types";

declare const logseq: any;

export interface TaskSyncResult {
  synced: number;
  deleted: number;
  failed: number;
  verified: number;
  completedRemote: number;
  updatedLocal: number;
  calendarUrl: string;
  tasks: LogseqTaskItem[];
  errors: string[];
}

export interface TaskListDiscoveryResult {
  ok: boolean;
  davRootUrl: string;
  principalUrl: string;
  homeSetUrl: string;
  taskLists: DiscoveredTaskList[];
  message: string;
}

interface ParsedDateTime {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  allDay: boolean;
  sortKey: number;
  date: string;
  time?: string;
}

interface TaskSyncCache {
  calendarUrl: string;
  uids: string[];
  syncedAt: string;
}

interface CalDavTaskItem {
  href: string;
  calendarData: string;
  uid: string;
  summary: string;
  status: string;
  percentComplete: number;
  completed: boolean;
  xLogseqBlockUuid: string;
}

const TASK_SYNC_KEY = "logseq-nextcloud-sync:task-sync";
const TASK_COLLECTION_RE = /(\/calendars\/[^/]+\/[^/]+\/?)$/i;
const TASK_MARKERS = new Set(["TODO", "DOING", "DONE", "LATER", "WAITING", "NOW", "CANCELLED"]);

export async function collectLogseqTasks(): Promise<LogseqTaskItem[]> {
  return collectLogseqTasksWithSettings({} as NextcloudSyncSettings);
}

export async function collectLogseqTasksForScope(scope: TaskScopeConfig, settings: NextcloudSyncSettings): Promise<LogseqTaskItem[]> {
  return collectLogseqTasksWithSettings({
    ...settings,
    taskFilterPageTypes: scope.filterPageTypes,
    taskFilterTags: scope.filterTags
  });
}

export async function collectLogseqTasksWithSettings(settings: NextcloudSyncSettings): Promise<LogseqTaskItem[]> {
  const pages = (await logseq.Editor.getAllPages?.()) ?? [];
  const tasks = new Map<string, LogseqTaskItem>();

  for (const page of pages) {
    const pageName = String(page?.name ?? page?.originalName ?? "").trim();
    if (!pageName) continue;
    const pageType = getPageType(page);
    const pageTags = getPageTags(page);
    if (!pageMatchesTaskScope(pageType, pageTags, settings)) continue;

    const blocks = (await logseq.Editor.getPageBlocksTree?.(pageName)) ?? [];
    for (const block of blocks) {
      collectTaskBlocksRecursive(pageName, block, tasks, pageType, pageTags);
    }
  }

  return Array.from(tasks.values()).sort(compareTasks);
}

export function buildTaskListIcs(tasks: LogseqTaskItem[], timezone: string) {
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Logseq Nextcloud Sync//EN",
    "CALSCALE:GREGORIAN",
    `X-WR-TIMEZONE:${escapeIcsText(timezone || "Europe/Stockholm")}`
  ];

  const timeZoneComponent = buildVTimeZoneComponent(timezone || "Europe/Stockholm");
  if (timeZoneComponent.length) {
    lines.push(...timeZoneComponent);
  }

  for (const task of tasks) {
    lines.push(...buildVTodo(task, timezone));
  }

  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}

export function buildTaskListFilename(prefix = "nextcloud-logseq-tasks") {
  const stamp = new Date().toISOString().slice(0, 10);
  return `${prefix}-${stamp}.ics`;
}

export function downloadTextFile(filename: string, text: string, mimeType = "text/plain;charset=utf-8") {
  const blob = new Blob([text], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  setTimeout(() => {
    URL.revokeObjectURL(url);
    link.remove();
  }, 0);
}

export async function exportTaskListIcs(tasks: LogseqTaskItem[], timezone: string) {
  const ics = buildTaskListIcs(tasks, timezone);
  const filename = buildTaskListFilename();
  downloadTextFile(filename, ics, "text/calendar;charset=utf-8");
  return { filename, ics };
}

export async function syncTasksToCalDav(tasks: LogseqTaskItem[], settings: NextcloudSyncSettings): Promise<TaskSyncResult> {
  const result: TaskSyncResult = {
    synced: 0,
    deleted: 0,
    failed: 0,
    verified: 0,
    completedRemote: 0,
    updatedLocal: 0,
    calendarUrl: "",
    tasks,
    errors: []
  };

  const calendarUrl = normalizeTaskCollectionUrl(settings.caldavTaskListUrl);
  result.calendarUrl = calendarUrl;
  if (!calendarUrl) {
    throw new Error("Set the exact Nextcloud task list collection URL in plugin settings first.");
  }
  if (!settings.caldavUsername || !settings.caldavPassword) {
    throw new Error("Set your Nextcloud username and app password in plugin settings first.");
  }

  const authHeader = `Basic ${btoa(`${settings.caldavUsername}:${settings.caldavPassword}`)}`;
  const previous = readTaskSyncCache(calendarUrl);
  const currentUids = new Set(tasks.map((task) => task.uid));

  const mirrorResult = await mirrorRemoteTaskCompletions(calendarUrl, authHeader, tasks);
  result.completedRemote = mirrorResult.completedRemote;
  result.updatedLocal = mirrorResult.updatedLocal;
  result.failed += mirrorResult.failed;
  result.errors.push(...mirrorResult.errors);

  const refreshedTasks = result.updatedLocal ? await collectLogseqTasksWithSettings(settings) : tasks;

  for (const task of refreshedTasks) {
    const ics = buildTaskListIcs([task], settings.calendarTimezone);
    const targetUrl = buildResourceUrl(calendarUrl, task.uid);

    try {
      await putCalDavText(targetUrl, authHeader, ics);
      result.synced += 1;

      try {
        await getCalDavText(targetUrl, authHeader);
        result.verified += 1;
      } catch (verifyError) {
        result.errors.push(formatRequestError(`GET ${task.title}`, verifyError));
      }
    } catch (error) {
      result.failed += 1;
      result.errors.push(formatRequestError(`PUT ${task.title}`, error));
    }
  }

  for (const uid of previous.uids) {
    if (currentUids.has(uid)) continue;
    const targetUrl = buildResourceUrl(calendarUrl, uid);
    try {
      await deleteCalDavText(targetUrl, authHeader);
      result.deleted += 1;
    } catch (error) {
      result.failed += 1;
      result.errors.push(formatRequestError(`DELETE ${uid}`, error));
    }
  }

  writeTaskSyncCache({
    calendarUrl,
    uids: refreshedTasks.map((task) => task.uid),
    syncedAt: new Date().toISOString()
  });

  result.tasks = refreshedTasks;
  return result;
}

export async function testCalDavConnection(settings: NextcloudSyncSettings) {
  return testCalDavConnectionForUrl(settings.caldavTaskListUrl, settings);
}

export async function testCalDavConnectionForUrl(taskListUrl: string, settings: NextcloudSyncSettings) {
  const calendarUrl = normalizeTaskCollectionUrl(taskListUrl);
  if (!calendarUrl) {
    return {
      ok: false,
      url: "",
      message: "Set the exact Nextcloud task list collection URL in plugin settings first."
    };
  }
  if (!settings.caldavUsername || !settings.caldavPassword) {
    return {
      ok: false,
      url: calendarUrl,
      message: "Set your Nextcloud username and app password in plugin settings first."
    };
  }

  const authHeader = `Basic ${btoa(`${settings.caldavUsername}:${settings.caldavPassword}`)}`;

  try {
    const response = await fetch(calendarUrl, {
      method: "OPTIONS",
      headers: {
        Authorization: authHeader
      },
      credentials: "include"
    });

    const responseText = await response.text();
    if (!response.ok) {
      throw new Error(`HTTP ${response.status} ${response.statusText}${responseText ? `: ${responseText.slice(0, 300)}` : ""}`);
    }

    return {
      ok: true,
      url: calendarUrl,
      message: `Task list collection reachable: ${calendarUrl}`,
      responsePreview: responseText.slice(0, 500)
    };
  } catch (error) {
    return {
      ok: false,
      url: calendarUrl,
      message: formatRequestError(`OPTIONS ${calendarUrl}`, error)
    };
  }
}

export async function discoverCalDavTaskLists(settings: NextcloudSyncSettings): Promise<TaskListDiscoveryResult> {
  if (!settings.caldavUsername || !settings.caldavPassword) {
    return {
      ok: false,
      davRootUrl: "",
      principalUrl: "",
      homeSetUrl: "",
      taskLists: [],
      message: "Set your Nextcloud username and app password in plugin settings first."
    };
  }

  const davRootUrl = resolveDavRootUrl(settings);
  if (!davRootUrl) {
    return {
      ok: false,
      davRootUrl: "",
      principalUrl: "",
      homeSetUrl: "",
      taskLists: [],
      message: "Set Nextcloud DAV root URL first, for example https://host/remote.php/dav"
    };
  }

  const authHeader = `Basic ${btoa(`${settings.caldavUsername}:${settings.caldavPassword}`)}`;
  try {
    const principalUrl = await discoverPrincipalUrl(davRootUrl, authHeader, settings.caldavUsername);
    const homeSetUrl = await discoverCalendarHomeSetUrl(principalUrl, authHeader);
    const taskLists = await listTaskListCollections(homeSetUrl, authHeader);

    return {
      ok: true,
      davRootUrl,
      principalUrl,
      homeSetUrl,
      taskLists,
      message: taskLists.length ? `Found ${taskLists.length} task lists.` : "No task lists found."
    };
  } catch (error) {
    return {
      ok: false,
      davRootUrl,
      principalUrl: "",
      homeSetUrl: "",
      taskLists: [],
      message: formatRequestError(`Discover task lists from ${davRootUrl}`, error)
    };
  }
}

export async function createCalDavTaskList(settings: NextcloudSyncSettings, displayName: string) {
  if (!settings.caldavUsername || !settings.caldavPassword) {
    throw new Error("Set your Nextcloud username and app password in plugin settings first.");
  }

  const davRootUrl = resolveDavRootUrl(settings);
  if (!davRootUrl) {
    throw new Error("Set Nextcloud DAV root URL first, for example https://host/remote.php/dav");
  }

  const authHeader = `Basic ${btoa(`${settings.caldavUsername}:${settings.caldavPassword}`)}`;
  const principalUrl = await discoverPrincipalUrl(davRootUrl, authHeader, settings.caldavUsername);
  const homeSetUrl = await discoverCalendarHomeSetUrl(principalUrl, authHeader);
  const slug = slugify(displayName) || `tasks-${Date.now()}`;
  const collectionUrl = `${stripTrailingSlash(homeSetUrl)}/${encodeURIComponent(slug)}/`;

  const response = await fetch(collectionUrl, {
    method: "MKCALENDAR",
    headers: {
      Authorization: authHeader,
      "Content-Type": "application/xml; charset=utf-8"
    },
    body: `<?xml version="1.0" encoding="UTF-8"?>
<c:mkcalendar xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">
  <d:set>
    <d:prop>
      <d:displayname>${escapeXml(displayName)}</d:displayname>
      <c:supported-calendar-component-set>
        <c:comp name="VTODO" />
      </c:supported-calendar-component-set>
    </d:prop>
  </d:set>
</c:mkcalendar>`,
    credentials: "include"
  });

  const responseText = await response.text();
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}${responseText ? `: ${responseText.slice(0, 400)}` : ""}`);
  }

  return {
    url: collectionUrl,
    displayName
  };
}

function collectTaskBlocksRecursive(
  pageName: string,
  block: any,
  tasks: Map<string, LogseqTaskItem>,
  pageType: string,
  pageTags: string[]
) {
  if (!block || typeof block !== "object") return;

  const content = typeof block.content === "string" ? block.content : "";
  const marker = getTaskState(content);
  if (marker) {
    const deadline = parseDateTime(getContentTimestamp(content, "DEADLINE"));
    const scheduled = parseDateTime(getContentTimestamp(content, "SCHEDULED"));
    const parsed = deadline ?? scheduled;
    const title = getBlockTitle(content);
    const blockUuid = typeof block.uuid === "string" ? block.uuid : "";
    const uid = blockUuid ? `logseq-task-${blockUuid}@logseq.local` : `logseq-task-${slugify(`${pageName}-${title}`)}@logseq.local`;
    const dueLabel = deadline ? "Deadline" : scheduled ? "Scheduled" : "";
    const dueValue = deadline ? formatHumanDate(deadline) : scheduled ? formatHumanDate(scheduled) : "";
    const descriptionParts = [
      `Page: ${pageName}`,
      blockUuid ? `Block UUID: ${blockUuid}` : "",
      dueLabel && dueValue ? `${dueLabel}: ${dueValue}` : "",
      content
    ].filter(Boolean);

    tasks.set(uid, {
      uid,
      pageName,
      title: title || pageName,
      description: descriptionParts.join("\n"),
      date: parsed?.date,
      time: parsed?.time,
      allDay: parsed?.allDay,
      unscheduled: !parsed?.date,
      sourceBlockUuid: blockUuid || undefined,
      sourceBlockContent: content,
      taskState: marker,
      marker,
      pageType,
      pageTags
    });
  }

  if (Array.isArray(block.children)) {
    for (const child of block.children) {
      collectTaskBlocksRecursive(pageName, child, tasks, pageType, pageTags);
    }
  }
}

function pageMatchesTaskScope(pageType: string, pageTags: string[], settings: NextcloudSyncSettings) {
  const allowedPageTypes = parseFilterList(settings.taskFilterPageTypes);
  const allowedTags = parseFilterList(settings.taskFilterTags);
  if (!allowedPageTypes.length && !allowedTags.length) return true;

  const hasPageTypeMatch = allowedPageTypes.length ? allowedPageTypes.includes(normalizeText(pageType)) : false;
  const normalizedTags = pageTags.map((tag) => normalizeText(tag)).filter(Boolean);
  const hasTagMatch = allowedTags.length ? allowedTags.some((tag) => normalizedTags.includes(tag)) : false;

  return hasPageTypeMatch || hasTagMatch;
}

function parseFilterList(input: string) {
  return String(input ?? "")
    .split(",")
    .map((item) => normalizeText(item))
    .filter(Boolean);
}

function getPageType(page: any) {
  const properties = normalizeProperties(page?.properties);
  const values = [properties["page-type"], properties.pagetype, properties.type, properties.kind].filter(Boolean);
  return normalizeText(String(values[0] ?? ""));
}

function getPageTags(page: any) {
  const properties = normalizeProperties(page?.properties);
  const candidates = [
    properties.tags,
    properties.tag,
    properties["page-tags"],
    properties.categories,
    properties.category
  ];
  const tags = candidates.flatMap((value) => normalizeToStringList(value));
  const unique = new Set(tags.map((tag) => normalizeText(tag)).filter(Boolean));
  return Array.from(unique);
}

function normalizeProperties(input: Record<string, any> | undefined | null) {
  const output: Record<string, any> = {};
  if (!input) return output;
  for (const [key, value] of Object.entries(input)) {
    output[normalizeKey(key)] = value;
  }
  return output;
}

function normalizeKey(input: string) {
  return String(input).trim().toLowerCase().replace(/\s+/g, "-");
}

function normalizeToStringList(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value.flatMap((item) => normalizeToStringList(item));
  }
  if (value == null) return [];
  const raw = String(value).trim();
  if (!raw) return [];
  return raw
    .split(/[,\n]/)
    .map((item) => item.replace(/^\[\[|\]\]$/g, "").trim())
    .filter(Boolean);
}

async function mirrorRemoteTaskCompletions(calendarUrl: string, authHeader: string, localTasks: LogseqTaskItem[]) {
  const result = {
    completedRemote: 0,
    updatedLocal: 0,
    failed: 0,
    errors: [] as string[]
  };

  const remoteItems = await fetchRemoteTaskItems(calendarUrl, authHeader);
  const localByUid = new Map(localTasks.map((task) => [task.uid, task]));

  for (const remote of remoteItems) {
    if (!remote.completed) continue;

    const task = localByUid.get(remote.uid);
    if (!task?.sourceBlockUuid) continue;

    try {
      const changed = await markLogseqBlockDone(task.sourceBlockUuid);
      if (changed) {
        result.completedRemote += 1;
        result.updatedLocal += 1;
      }
    } catch (error) {
      result.failed += 1;
      result.errors.push(formatRequestError(`Mirror ${remote.summary || remote.uid}`, error));
    }
  }

  return result;
}

async function discoverPrincipalUrl(davRootUrl: string, authHeader: string, username: string) {
  const responseText = await propfindText(
    davRootUrl,
    authHeader,
    "0",
    `<?xml version="1.0" encoding="UTF-8"?>
<d:propfind xmlns:d="DAV:">
  <d:prop>
    <d:current-user-principal />
  </d:prop>
</d:propfind>`
  );

  const doc = new DOMParser().parseFromString(responseText, "application/xml");
  const principalNode = Array.from(doc.getElementsByTagNameNS("*", "current-user-principal"))[0];
  const href = principalNode ? getNodeText(principalNode, "href") : "";
  if (href) {
    return resolveUrlFromHref(davRootUrl, href);
  }

  return new URL(`/remote.php/dav/principals/users/${encodeURIComponent(username)}/`, davRootUrl).toString();
}

async function discoverCalendarHomeSetUrl(principalUrl: string, authHeader: string) {
  const responseText = await propfindText(
    principalUrl,
    authHeader,
    "0",
    `<?xml version="1.0" encoding="UTF-8"?>
<d:propfind xmlns:d="DAV:" xmlns:cd="urn:ietf:params:xml:ns:caldav">
  <d:prop>
    <cd:calendar-home-set />
  </d:prop>
</d:propfind>`
  );

  const doc = new DOMParser().parseFromString(responseText, "application/xml");
  const parent = Array.from(doc.getElementsByTagNameNS("*", "calendar-home-set"))[0];
  const href = parent ? getNodeText(parent, "href") : "";
  if (!href) {
    throw new Error("Could not find calendar-home-set in DAV discovery response.");
  }
  return resolveUrlFromHref(principalUrl, href);
}

async function listTaskListCollections(homeSetUrl: string, authHeader: string) {
  const responseText = await propfindText(
    homeSetUrl,
    authHeader,
    "1",
    `<?xml version="1.0" encoding="UTF-8"?>
<d:propfind xmlns:d="DAV:" xmlns:cd="urn:ietf:params:xml:ns:caldav">
  <d:prop>
    <d:displayname />
    <d:resourcetype />
    <cd:supported-calendar-component-set />
  </d:prop>
</d:propfind>`
  );

  const doc = new DOMParser().parseFromString(responseText, "application/xml");
  const responses = Array.from(doc.getElementsByTagNameNS("*", "response"));
  const taskLists: DiscoveredTaskList[] = [];

  for (const response of responses) {
    const href = getNodeText(response, "href");
    const absoluteUrl = resolveUrlFromHref(homeSetUrl, href);
    if (!href || stripTrailingSlash(absoluteUrl) === stripTrailingSlash(homeSetUrl)) continue;

    const displayName = getPropValue(response, "displayname") || href.split("/").filter(Boolean).pop() || "Task List";
    const resourceTypeNode = Array.from(response.getElementsByTagNameNS("*", "resourcetype"))[0];
    const resourceTypeNames = resourceTypeNode
      ? Array.from(resourceTypeNode.children).map((node) => node.localName?.toLowerCase() || "")
      : [];
    const componentNodes = Array.from(response.getElementsByTagNameNS("*", "comp"));
    const componentSet = componentNodes.map((node) => node.getAttribute("name") || "").filter(Boolean);
    const isTaskListCollection = resourceTypeNames.includes("calendar");

    if (!isTaskListCollection) continue;
    if (componentSet.length && !componentSet.includes("VTODO")) continue;

    taskLists.push({
      url: absoluteUrl,
      href,
      displayName,
      componentSet,
      isTaskListCollection
    });
  }

  return taskLists.sort((a, b) => a.displayName.localeCompare(b.displayName) || a.url.localeCompare(b.url));
}

async function propfindText(url: string, authHeader: string, depth: "0" | "1", body: string) {
  const response = await fetch(url, {
    method: "PROPFIND",
    headers: {
      Authorization: authHeader,
      Depth: depth,
      "Content-Type": "application/xml; charset=utf-8"
    },
    body,
    credentials: "include"
  });
  const responseText = await response.text();
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}${responseText ? `: ${responseText.slice(0, 500)}` : ""}`);
  }
  return responseText;
}

async function fetchRemoteTaskItems(calendarUrl: string, authHeader: string): Promise<CalDavTaskItem[]> {
  const response = await fetch(calendarUrl, {
    method: "REPORT",
    headers: {
      Authorization: authHeader,
      Depth: "1",
      "Content-Type": "application/xml; charset=utf-8"
    },
    body: `<?xml version="1.0" encoding="UTF-8"?>
<c:calendar-query xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">
  <d:prop>
    <d:getetag />
    <d:displayname />
    <d:getcontenttype />
    <c:calendar-data />
  </d:prop>
  <c:filter>
    <c:comp-filter name="VCALENDAR">
      <c:comp-filter name="VTODO" />
    </c:comp-filter>
  </c:filter>
</c:calendar-query>`,
    credentials: "include"
  });

  const responseText = await response.text();
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}${responseText ? `: ${responseText.slice(0, 500)}` : ""}`);
  }

  const items: CalDavTaskItem[] = [];
  const doc = new DOMParser().parseFromString(responseText, "application/xml");
  const responses = Array.from(doc.getElementsByTagNameNS("*", "response"));

  for (const responseNode of responses) {
    const href = getNodeText(responseNode, "href");
    const calendarData = getNodeText(responseNode, "calendar-data");
    const todos = parseVTodoData(calendarData);
    for (const todo of todos) {
      items.push({
        href,
        calendarData,
        uid: todo["UID"] || "",
        summary: todo["SUMMARY"] || "",
        status: String(todo["STATUS"] || ""),
        percentComplete: Number(todo["PERCENT-COMPLETE"] || 0),
        completed: isCompletedRemoteTodo(todo),
        xLogseqBlockUuid: todo["X-LOGSEQ-BLOCK-UUID"] || ""
      });
    }
  }

  return items.filter((item) => Boolean(item.uid));
}

async function markLogseqBlockDone(blockUuid: string) {
  const editor = logseq.Editor as any;
  const block = await editor.getBlock?.(blockUuid, { includeChildren: false });
  if (!block || typeof block.content !== "string") return false;
  if (/^\s*DONE\b/i.test(block.content)) return false;

  const updated = String(block.content).replace(/^\s*(TODO|DOING|LATER|WAITING|NOW|CANCELLED)\b/i, "DONE");
  await editor.updateBlock?.(blockUuid, updated);
  return updated !== block.content;
}

async function putCalDavText(url: string, authHeader: string, text: string) {
  const response = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: authHeader,
      "Content-Type": "text/calendar; charset=utf-8"
    },
    body: text,
    credentials: "include"
  });
  const responseText = await response.text();
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}${responseText ? `: ${responseText.slice(0, 500)}` : ""}`);
  }
}

async function getCalDavText(url: string, authHeader: string) {
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: authHeader
    },
    credentials: "include"
  });
  const responseText = await response.text();
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}${responseText ? `: ${responseText.slice(0, 500)}` : ""}`);
  }
  return responseText;
}

async function deleteCalDavText(url: string, authHeader: string) {
  const response = await fetch(url, {
    method: "DELETE",
    headers: {
      Authorization: authHeader
    },
    credentials: "include"
  });
  const responseText = await response.text();
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}${responseText ? `: ${responseText.slice(0, 500)}` : ""}`);
  }
}

function buildVTodo(task: LogseqTaskItem, timezone: string) {
  const uid = escapeIcsText(task.uid);
  const summary = escapeIcsText(task.title);
  const description = escapeIcsText(task.description || task.pageName);
  const dtstamp = formatUtcDateTime(new Date());
  const isCompleted = isCompletedState(task.taskState);
  const lines = [
    "BEGIN:VTODO",
    `UID:${uid}`,
    `DTSTAMP:${dtstamp}`,
    `SUMMARY:${summary}`,
    `DESCRIPTION:${description}`,
    isCompleted ? "STATUS:COMPLETED" : "STATUS:NEEDS-ACTION",
    "PRIORITY:5",
    isCompleted ? "PERCENT-COMPLETE:100" : "PERCENT-COMPLETE:0",
    isCompleted ? `COMPLETED:${dtstamp}` : "",
    `X-LOGSEQ-PAGE:${escapeIcsText(task.pageName)}`,
    task.sourceBlockUuid ? `X-LOGSEQ-BLOCK-UUID:${escapeIcsText(task.sourceBlockUuid)}` : "",
    "END:VTODO"
  ];

  if (task.date) {
    const dateTimeValue = task.allDay
      ? `;VALUE=DATE:${task.date}`
      : `;TZID=${escapeIcsText(timezone || "Europe/Stockholm")}:${task.date}T${task.time ?? "090000"}`;
    lines.splice(5, 0, `DTSTART${dateTimeValue}`, `DUE${dateTimeValue}`);
  }

  return lines.filter(Boolean);
}

function buildVTimeZoneComponent(timezone: string) {
  const tz = String(timezone || "Europe/Stockholm").trim();
  if (tz !== "Europe/Stockholm") return [];

  return [
    "BEGIN:VTIMEZONE",
    "TZID:Europe/Stockholm",
    "X-LIC-LOCATION:Europe/Stockholm",
    "BEGIN:DAYLIGHT",
    "TZOFFSETFROM:+0100",
    "TZOFFSETTO:+0200",
    "TZNAME:CEST",
    "DTSTART:19700329T020000",
    "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU",
    "END:DAYLIGHT",
    "BEGIN:STANDARD",
    "TZOFFSETFROM:+0200",
    "TZOFFSETTO:+0100",
    "TZNAME:CET",
    "DTSTART:19701025T030000",
    "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU",
    "END:STANDARD",
    "END:VTIMEZONE"
  ];
}

function buildResourceUrl(calendarUrl: string, uid: string) {
  const base = calendarUrl.replace(/\/+$/, "");
  return `${base}/${encodeURIComponent(uid)}.ics`;
}

function normalizeTaskCollectionUrl(input: string) {
  const raw = String(input ?? "").trim();
  if (!raw) return "";

  const url = new URL(raw);
  url.hash = "";
  url.search = "";
  const pathname = url.pathname.replace(/\/+$/, "");
  if (!TASK_COLLECTION_RE.test(pathname)) return "";
  return `${url.origin}${pathname}`;
}

function resolveDavRootUrl(settings: NextcloudSyncSettings) {
  const raw = String(settings.nextcloudDavUrl ?? "").trim();
  if (raw) {
    const url = new URL(raw);
    url.hash = "";
    url.search = "";
    return stripTrailingSlash(url.toString());
  }

  const source = settings.caldavTaskListUrl || settings.caldavCalendarUrl;
  if (!source) return "";
  const url = new URL(source);
  return `${url.origin}/remote.php/dav`;
}

function resolveUrlFromHref(baseUrl: string, href: string) {
  return new URL(String(href || "").trim(), baseUrl).toString();
}

function stripTrailingSlash(input: string) {
  return String(input ?? "").replace(/\/+$/, "");
}

function parseDateTime(value: unknown): ParsedDateTime | null {
  if (value == null) return null;

  const raw = Array.isArray(value) ? value.map((item) => String(item)).join(" ") : String(value);
  const normalized = raw.trim().replace(/[<>]/g, "");
  if (!normalized) return null;

  const candidate = normalized.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[^\d]+(\d{1,2})(?::(\d{2}))?\s*(am|pm)?)?/i);
  if (!candidate) return null;

  const year = Number(candidate[1]);
  const month = Number(candidate[2]);
  const day = Number(candidate[3]);
  let hour = candidate[4] ? Number(candidate[4]) : 0;
  const minute = candidate[5] ? Number(candidate[5]) : 0;
  const meridiem = candidate[6]?.toLowerCase();
  const hasTime = Boolean(candidate[4]);

  if (meridiem === "pm" && hour < 12) hour += 12;
  if (meridiem === "am" && hour === 12) hour = 0;

  return {
    year,
    month,
    day,
    hour,
    minute,
    allDay: !hasTime,
    sortKey: new Date(year, month - 1, day, hour, minute, 0, 0).getTime(),
    date: `${pad(year, 4)}${pad(month)}${pad(day)}`,
    time: hasTime ? `${pad(hour)}${pad(minute)}00` : undefined
  };
}

function parseVTodoData(calendarData: string) {
  const text = unfoldIcsText(String(calendarData ?? ""));
  const blocks = extractIcsComponentBlocks(text, "VTODO");
  return blocks.map(parseIcsProperties).filter((item) => Object.keys(item).length > 0);
}

function extractIcsComponentBlocks(text: string, component: string) {
  const lines = text.split(/\r?\n/);
  const blocks: string[] = [];
  const startMarker = `BEGIN:${component}`;
  const endMarker = `END:${component}`;
  let depth = 0;
  let buffer: string[] = [];

  for (const line of lines) {
    if (line === startMarker) {
      depth += 1;
      if (depth === 1) buffer = [line];
      else buffer.push(line);
      continue;
    }

    if (!depth) continue;

    buffer.push(line);
    if (line === endMarker) {
      depth -= 1;
      if (depth === 0) {
        blocks.push(buffer.join("\n"));
        buffer = [];
      }
    }
  }

  return blocks;
}

function parseIcsProperties(componentText: string) {
  const lines = componentText.split(/\r?\n/);
  const props: Record<string, string> = {};

  for (const rawLine of lines.slice(1, -1)) {
    if (!rawLine || rawLine.startsWith(" ") || rawLine.startsWith("\t")) continue;
    const idx = rawLine.indexOf(":");
    if (idx === -1) continue;
    const keyPart = rawLine.slice(0, idx);
    const value = unescapeIcsText(rawLine.slice(idx + 1));
    const key = keyPart.split(";")[0].toUpperCase();
    props[key] = value;
  }

  return props;
}

function unfoldIcsText(text: string) {
  return text.replace(/\r?\n[ \t]/g, "");
}

function unescapeIcsText(text: string) {
  return String(text ?? "")
    .replace(/\\n/g, "\n")
    .replace(/\\,/g, ",")
    .replace(/\\;/g, ";")
    .replace(/\\\\/g, "\\");
}

function isCompletedRemoteTodo(todo: Record<string, string>) {
  const status = normalizeText(todo["STATUS"] || "");
  const percent = Number(todo["PERCENT-COMPLETE"] || 0);
  return status === "completed" || status === "cancelled" || percent >= 100;
}

function compareTasks(a: LogseqTaskItem, b: LogseqTaskItem) {
  const aKey = taskSortKey(a);
  const bKey = taskSortKey(b);
  return aKey - bKey || a.title.localeCompare(b.title) || a.uid.localeCompare(b.uid);
}

function taskSortKey(task: LogseqTaskItem) {
  if (!task.date) return Number.MAX_SAFE_INTEGER;
  const year = Number(task.date.slice(0, 4));
  const month = Number(task.date.slice(4, 6));
  const day = Number(task.date.slice(6, 8));
  const hour = task.allDay || !task.time ? 0 : Number(task.time.slice(0, 2));
  const minute = task.allDay || !task.time ? 0 : Number(task.time.slice(2, 4));
  return new Date(year, month - 1, day, hour, minute, 0, 0).getTime();
}

function getContentTimestamp(content: string, key: string) {
  const marker = String(key).trim().toUpperCase();
  const pattern = new RegExp(`${marker}\\s*:\\s*<([^>]+)>`, "i");
  return content.match(pattern)?.[1] ?? "";
}

function getBlockTitle(content: string) {
  return String(content)
    .replace(/^\s*(TODO|DOING|DONE|LATER|WAITING|NOW|CANCELLED)\s+/i, "")
    .replace(/SCHEDULED:\s*<[^>]+>/gi, "")
    .replace(/DEADLINE:\s*<[^>]+>/gi, "")
    .trim()
    .split("\n")[0]
    ?.trim() || "";
}

function getTaskState(content: string) {
  const match = String(content ?? "").match(/^\s*(TODO|DOING|DONE|LATER|WAITING|NOW|CANCELLED)\b/i);
  const marker = match?.[1]?.toUpperCase() || "";
  return TASK_MARKERS.has(marker) ? marker : "";
}

function isCompletedState(taskState?: string) {
  const normalized = normalizeText(taskState || "");
  return normalized === "done" || normalized === "completed" || normalized === "cancelled";
}

function normalizeText(input: string) {
  return String(input ?? "").trim().toLowerCase();
}

function slugify(input: string) {
  return String(input)
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function pad(value: number, length = 2) {
  return String(value).padStart(length, "0");
}

function formatUtcDateTime(date: Date) {
  return [
    pad(date.getUTCFullYear(), 4),
    pad(date.getUTCMonth() + 1),
    pad(date.getUTCDate())
  ].join("") + `T${pad(date.getUTCHours())}${pad(date.getUTCMinutes())}${pad(date.getUTCSeconds())}Z`;
}

function escapeXml(input: string) {
  return String(input ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function formatHumanDate(parsed: ParsedDateTime) {
  const date = `${pad(parsed.year, 4)}-${pad(parsed.month)}-${pad(parsed.day)}`;
  return parsed.allDay ? date : `${date} ${pad(parsed.hour)}:${pad(parsed.minute)}`;
}

function escapeIcsText(input: string) {
  return String(input ?? "")
    .replace(/\\/g, "\\\\")
    .replace(/\r?\n/g, "\\n")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,");
}

function readTaskSyncCache(calendarUrl: string): TaskSyncCache {
  try {
    const raw = localStorage.getItem(TASK_SYNC_KEY);
    if (!raw) return { calendarUrl, uids: [], syncedAt: "" };
    const parsed = JSON.parse(raw);
    if (parsed?.calendarUrl !== calendarUrl || !Array.isArray(parsed?.uids)) {
      return { calendarUrl, uids: [], syncedAt: "" };
    }
    return {
      calendarUrl,
      uids: parsed.uids.filter((uid: unknown) => typeof uid === "string"),
      syncedAt: typeof parsed.syncedAt === "string" ? parsed.syncedAt : ""
    };
  } catch {
    return { calendarUrl, uids: [], syncedAt: "" };
  }
}

function writeTaskSyncCache(cache: TaskSyncCache) {
  localStorage.setItem(TASK_SYNC_KEY, JSON.stringify(cache));
}

function formatRequestError(action: string, error: unknown) {
  if (error instanceof Error) return `${action}: ${error.message}`;
  return `${action}: ${String(error)}`;
}

function getNodeText(root: Element, localName: string) {
  const node = Array.from(root.getElementsByTagNameNS("*", localName))[0];
  return node?.textContent?.trim() ?? "";
}

function getPropValue(response: Element, localName: string) {
  const prop = Array.from(response.getElementsByTagNameNS("*", "prop"))[0];
  if (!prop) return "";
  return getNodeText(prop, localName);
}
