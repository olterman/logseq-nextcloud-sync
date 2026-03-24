import type { CalendarEventKind, CalendarScopeConfig, DiscoveredCalendar, LogseqCalendarEvent, NextcloudSyncSettings } from "./types";

declare const logseq: any;

export interface CalendarSyncResult {
  synced: number;
  deleted: number;
  failed: number;
  verified: number;
  calendarUrl: string;
  events: LogseqCalendarEvent[];
  errors: string[];
}

export interface CalendarDiscoveryResult {
  ok: boolean;
  davRootUrl: string;
  principalUrl: string;
  homeSetUrl: string;
  calendars: DiscoveredCalendar[];
  message: string;
}

export interface RemoteCalendarImportItem {
  uid: string;
  title: string;
  description: string;
  date: string;
  time?: string;
  endDate?: string;
  endTime?: string;
  allDay?: boolean;
  sourcePage?: string;
  sourceBlockUuid?: string;
  remoteResourceUrl?: string;
}

interface ParsedDateTime {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  allDay: boolean;
  date: string;
  time?: string;
  sortKey: number;
}

interface CalendarSyncCache {
  calendarUrl: string;
  uids: string[];
  syncedAt: string;
}

const CALENDAR_SYNC_KEY = "logseq-nextcloud-sync:calendar-sync";
const CALENDAR_COLLECTION_RE = /(\/calendars\/[^/]+\/[^/]+\/?)$/i;
const DEFAULT_EVENT_DURATION_MINUTES = 60;
const PROPERTY_KIND_MAP: Record<CalendarEventKind, string[]> = {
  scheduled: ["scheduled"],
  deadline: ["deadline", "due", "due-date"],
  date: ["date", "event-date", "meeting-date"],
  start: ["start", "start-date", "start-time"]
};

export async function collectLogseqCalendarEvents(): Promise<LogseqCalendarEvent[]> {
  return collectLogseqCalendarEventsWithScope();
}

export async function collectLogseqCalendarEventsForScope(scope: CalendarScopeConfig): Promise<LogseqCalendarEvent[]> {
  return collectLogseqCalendarEventsWithScope(scope);
}

async function collectLogseqCalendarEventsWithScope(
  scope?: Pick<CalendarScopeConfig, "id" | "propertyKey" | "propertyValue" | "filterPageTypes" | "filterTags" | "ignoredTags" | "prefilterPagesOnly"> | null
): Promise<LogseqCalendarEvent[]> {
  const pages = (await logseq.Editor.getAllPages?.()) ?? [];
  const events = new Map<string, LogseqCalendarEvent>();

  for (const page of pages) {
    const pageName = String(page?.name ?? page?.originalName ?? "").trim();
    if (!pageName) continue;
    if (isSyncConflictPage(pageName)) continue;

    const pageProperties = normalizeProperties(page?.properties);
    const pageType = getPageType(page);
    const pageTags = getPageTags(page);
    if (hasIgnoredScopeTags(pageTags, scope)) continue;
    const pageMatches = pageMatchesCalendarScope(pageType, pageTags, scope);
    const blocks = (await logseq.Editor.getPageBlocksTree?.(pageName)) ?? [];
    if (scope?.prefilterPagesOnly && !pageMatches && !pageContainsExplicitProfileBlocks(blocks, scope?.id || "")) continue;
    const pageEvents = isJournalLikePage(page, pageType, pageTags)
      ? []
      : extractPropertyEvents(pageName, pageProperties, pageType, pageTags, pageTags, undefined, undefined, undefined, scope);
    for (const event of pageEvents) {
      if (events.has(event.uid)) continue;
      events.set(event.uid, event);
    }
    for (const block of blocks) {
      collectBlockEventsRecursive(pageName, pageProperties, pageType, pageTags, block, events, scope);
    }
  }

  for (const event of await collectImportedMarkdownEvents(scope)) {
    if (events.has(event.uid)) continue;
    events.set(event.uid, event);
  }

  return Array.from(events.values()).sort(compareEvents);
}

export function buildCalendarIcs(events: LogseqCalendarEvent[], timezone: string) {
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

  for (const event of events) {
    lines.push(...buildVEvent(event, timezone));
  }

  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}

export function buildCalendarFilename(prefix = "nextcloud-logseq-calendar") {
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

export async function exportCalendarIcs(events: LogseqCalendarEvent[], timezone: string) {
  const ics = buildCalendarIcs(events, timezone);
  const filename = buildCalendarFilename();
  downloadTextFile(filename, ics, "text/calendar;charset=utf-8");
  return { filename, ics };
}

export async function syncCalendarToCalDav(
  events: LogseqCalendarEvent[],
  settings: NextcloudSyncSettings,
  protectedUids: ReadonlySet<string> = new Set()
): Promise<CalendarSyncResult> {
  const result: CalendarSyncResult = {
    synced: 0,
    deleted: 0,
    failed: 0,
    verified: 0,
    calendarUrl: "",
    events,
    errors: []
  };

  const calendarUrl = normalizeCalendarCollectionUrl(settings.caldavCalendarUrl);
  result.calendarUrl = calendarUrl;
  if (!calendarUrl) {
    throw new Error("Set the exact Nextcloud calendar collection URL in plugin settings first.");
  }
  if (!settings.caldavUsername || !settings.caldavPassword) {
    throw new Error("Set your Nextcloud username and app password in plugin settings first.");
  }

  const authHeader = `Basic ${btoa(`${settings.caldavUsername}:${settings.caldavPassword}`)}`;
  const previous = readCalendarSyncCache(calendarUrl);
  const currentUids = new Set(events.map((event) => event.uid));
  let remoteResourceUrls = new Map<string, string>();

  try {
    remoteResourceUrls = await fetchRemoteCalendarResourceUrls(calendarUrl, authHeader);
  } catch (error) {
    console.warn("[nextcloud-sync] could not preload remote calendar resource URLs", error);
  }

  for (const event of events) {
    const ics = buildCalendarIcs([event], settings.calendarTimezone);
    const targetUrl = event.remoteResourceUrl || remoteResourceUrls.get(event.uid) || buildResourceUrl(calendarUrl, event.uid);

    try {
      console.warn(`[nextcloud-sync] calendar PUT ${JSON.stringify({
        uid: event.uid,
        title: event.title,
        date: event.date,
        time: event.time,
        targetUrl
      })}`);
      await putCalDavText(targetUrl, authHeader, ics);
      result.synced += 1;

      try {
        await getCalDavText(targetUrl, authHeader);
        result.verified += 1;
      } catch (verifyError) {
        result.errors.push(formatRequestError(`GET ${event.title}`, verifyError));
      }
    } catch (error) {
      result.failed += 1;
      result.errors.push(formatRequestError(`PUT ${event.title}`, error));
    }
  }

  for (const uid of previous.uids) {
    if (currentUids.has(uid)) continue;
    if (protectedUids.has(uid)) continue;
    const targetUrl = remoteResourceUrls.get(uid) || buildResourceUrl(calendarUrl, uid);
    try {
      await deleteCalDavText(targetUrl, authHeader);
      result.deleted += 1;
    } catch (error) {
      result.failed += 1;
      result.errors.push(formatRequestError(`DELETE ${uid}`, error));
    }
  }

  writeCalendarSyncCache({
    calendarUrl,
    uids: events.map((event) => event.uid),
    syncedAt: new Date().toISOString()
  });

  return result;
}

export async function testCalendarConnection(settings: NextcloudSyncSettings) {
  return testCalendarConnectionForUrl(settings.caldavCalendarUrl, settings);
}

export async function testCalendarConnectionForUrl(calendarCollectionUrl: string, settings: NextcloudSyncSettings) {
  const calendarUrl = normalizeCalendarCollectionUrl(calendarCollectionUrl);
  if (!calendarUrl) {
    return {
      ok: false,
      url: "",
      message: "Set the exact Nextcloud calendar collection URL in plugin settings first."
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
      message: `Calendar collection reachable: ${calendarUrl}`,
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

export async function discoverCalDavCalendars(settings: NextcloudSyncSettings): Promise<CalendarDiscoveryResult> {
  if (!settings.caldavUsername || !settings.caldavPassword) {
    return {
      ok: false,
      davRootUrl: "",
      principalUrl: "",
      homeSetUrl: "",
      calendars: [],
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
      calendars: [],
      message: "Set Nextcloud DAV root URL first, for example https://host/remote.php/dav"
    };
  }

  const authHeader = `Basic ${btoa(`${settings.caldavUsername}:${settings.caldavPassword}`)}`;

  try {
    const principalUrl = await discoverPrincipalUrl(davRootUrl, authHeader, settings.caldavUsername);
    const homeSetUrl = await discoverCalendarHomeSetUrl(principalUrl, authHeader);
    const calendars = await listCalendarCollections(homeSetUrl, authHeader);

    return {
      ok: true,
      davRootUrl,
      principalUrl,
      homeSetUrl,
      calendars,
      message: calendars.length ? `Found ${calendars.length} calendars.` : "No calendars found."
    };
  } catch (error) {
    return {
      ok: false,
      davRootUrl,
      principalUrl: "",
      homeSetUrl: "",
      calendars: [],
      message: formatRequestError(`Discover calendars from ${davRootUrl}`, error)
    };
  }
}

function collectBlockEventsRecursive(
  pageName: string,
  pageProperties: Record<string, any>,
  pageType: string,
  pageTags: string[],
  block: any,
  events: Map<string, LogseqCalendarEvent>,
  scope?: Pick<CalendarScopeConfig, "id" | "propertyKey" | "propertyValue" | "filterPageTypes" | "filterTags" | "ignoredTags" | "prefilterPagesOnly"> | null,
  ancestry: string[] = []
) {
  if (!block || typeof block !== "object") return;

  const blockProperties = normalizeProperties({
    ...(block.properties ?? {}),
    ...(block.meta?.properties ?? {}),
    ...readPropertyChildren(block)
  });
  const effectiveProperties = {
    ...pageProperties,
    ...blockProperties
  };
  const content = typeof block.content === "string" ? block.content : "";
  const nextAncestry = content ? [...ancestry, content] : ancestry;
  if (isPropertyOnlyBlock(content)) {
    return;
  }
  if (isTaskBlock(content)) {
    if (Array.isArray(block.children)) {
      for (const child of block.children) {
        collectBlockEventsRecursive(pageName, pageProperties, pageType, pageTags, child, events, scope, nextAncestry);
      }
    }
    return;
  }
  const blockType = getPageTypeFromProperties(blockProperties);
  const blockTags = mergeNormalizedTags(getTagsFromProperties(blockProperties), getInlineTags(content));
  const explicitProfileId = getExplicitProfileId(effectiveProperties);
  const profileMismatch = explicitProfileId && explicitProfileId !== normalizeText(scope?.id || "");
  const hasBlockScopeOverride = Boolean(blockType) || blockTags.length > 0;
  const effectiveType = hasBlockScopeOverride ? blockType : pageType;
  const effectiveTags = hasBlockScopeOverride ? blockTags : pageTags;
  const allScopeTags = mergeNormalizedTags(pageTags, blockTags);
  const title = getBlockTitle(content) || pageName;
  const blockUuid = typeof block.uuid === "string" ? block.uuid : undefined;
  const remoteUid = getRemoteSyncUid(effectiveProperties);
  const remoteResourceUrl = getRemoteResourceUrl(effectiveProperties);
  const insideImportedSection = isInsideImportedNextcloudSection(nextAncestry);

  if (profileMismatch) {
    if (Array.isArray(block.children)) {
      for (const child of block.children) {
        collectBlockEventsRecursive(pageName, pageProperties, pageType, pageTags, child, events, scope, nextAncestry);
      }
    }
    return;
  }

  if (insideImportedSection && !remoteUid) {
    if (Array.isArray(block.children)) {
      for (const child of block.children) {
        collectBlockEventsRecursive(pageName, pageProperties, pageType, pageTags, child, events, scope, nextAncestry);
      }
    }
    return;
  }

  for (const event of extractPropertyEvents(
    pageName,
    effectiveProperties,
    effectiveType,
    effectiveTags,
    allScopeTags,
    blockUuid,
    title,
    content,
    scope
  )) {
    if (events.has(event.uid)) continue;
    events.set(event.uid, event);
  }

  const inlineTimestampKinds: Array<[CalendarEventKind, RegExp]> = [
    ["scheduled", /SCHEDULED:\s*<([^>]+)>/i],
    ["deadline", /DEADLINE:\s*<([^>]+)>/i]
  ];

  for (const [kind, pattern] of inlineTimestampKinds) {
    const match = content.match(pattern);
    const parsed = match?.[1] ? parseDateTime(match[1]) : null;
    if (!parsed) continue;
    const scopeMatch = matchCalendarScope(effectiveType, effectiveTags, allScopeTags, effectiveProperties, scope);
    if (!scopeMatch.matches) continue;
    const event = makeEvent({
      kind,
      pageName,
      parsed,
      scopePropertyKey: scopeMatch.propertyKey,
      scopePropertyValue: scopeMatch.propertyValue,
      sourceBlockUuid: blockUuid,
      sourceTitle: title,
      sourceText: content,
      explicitUid: remoteUid
    });
    if (events.has(event.uid)) continue;
    events.set(event.uid, event);
  }

  if (Array.isArray(block.children)) {
    for (const child of block.children) {
      collectBlockEventsRecursive(pageName, pageProperties, pageType, pageTags, child, events, scope, nextAncestry);
    }
  }
}

function isInsideImportedNextcloudSection(ancestry: string[]) {
  const normalized = ancestry.map((entry) => normalizeText(entry)).filter(Boolean);
  if (!normalized.includes("nextcloud")) return false;
  return normalized.includes("events") || normalized.includes("tasks");
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
  const principalHref = getFirstTagText(doc, "href", "current-user-principal");
  if (principalHref) {
    return resolveUrlFromHref(davRootUrl, principalHref);
  }

  const fallback = new URL(`/remote.php/dav/principals/users/${encodeURIComponent(username)}/`, davRootUrl).toString();
  return fallback;
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
  const href = getFirstTagText(doc, "href", "calendar-home-set");
  if (!href) {
    throw new Error("Could not find calendar-home-set in DAV discovery response.");
  }

  return resolveUrlFromHref(principalUrl, href);
}

async function listCalendarCollections(homeSetUrl: string, authHeader: string) {
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
  const collections: DiscoveredCalendar[] = [];

  for (const response of responses) {
    const href = getNodeText(response, "href");
    const absoluteUrl = resolveUrlFromHref(homeSetUrl, href);
    if (!href || stripTrailingSlash(absoluteUrl) === stripTrailingSlash(homeSetUrl)) continue;

    const displayName = getPropValue(response, "displayname") || href.split("/").filter(Boolean).pop() || "Calendar";
    const resourceTypeNode = Array.from(response.getElementsByTagNameNS("*", "resourcetype"))[0];
    const resourceTypeNames = resourceTypeNode
      ? Array.from(resourceTypeNode.children).map((node) => node.localName?.toLowerCase() || "")
      : [];
    const componentNodes = Array.from(response.getElementsByTagNameNS("*", "comp"));
    const componentSet = componentNodes.map((node) => node.getAttribute("name") || "").filter(Boolean);
    const isCalendarCollection = resourceTypeNames.includes("calendar");

    if (!isCalendarCollection) continue;
    if (componentSet.length && !componentSet.includes("VEVENT")) continue;

    collections.push({
      url: absoluteUrl,
      href,
      displayName,
      componentSet,
      isCalendarCollection
    });
  }

  return collections.sort((a, b) => a.displayName.localeCompare(b.displayName) || a.url.localeCompare(b.url));
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

function extractPropertyEvents(
  pageName: string,
  properties: Record<string, any>,
  pageType: string,
  pageTags: string[],
  allScopeTags: string[],
  sourceBlockUuid?: string,
  sourceTitle?: string,
  sourceText?: string,
  scope?: Pick<CalendarScopeConfig, "id" | "propertyKey" | "propertyValue" | "filterPageTypes" | "filterTags" | "ignoredTags" | "prefilterPagesOnly"> | null
) {
  const events: LogseqCalendarEvent[] = [];
  const scopeMatch = matchCalendarScope(pageType, pageTags, allScopeTags, properties, scope);
  if (!scopeMatch.matches) return events;

  for (const [kind, keys] of Object.entries(PROPERTY_KIND_MAP) as Array<[CalendarEventKind, string[]]>) {
    for (const key of keys) {
      const value = getPropertyValue(properties, key);
      const parsed = parseDateTime(value);
      if (!parsed) continue;
      events.push(
        makeEvent({
          kind,
          pageName,
          parsed,
          scopePropertyKey: scopeMatch.propertyKey,
          scopePropertyValue: scopeMatch.propertyValue,
          sourceBlockUuid,
          sourceTitle,
          sourceText,
          propertyKey: key,
          explicitUid: getRemoteSyncUid(properties),
          remoteResourceUrl: getRemoteResourceUrl(properties)
        })
      );
    }
  }

  return events;
}

function makeEvent({
  kind,
  pageName,
  parsed,
  scopePropertyKey,
  scopePropertyValue,
  sourceBlockUuid,
  sourceTitle,
  sourceText,
  propertyKey,
  explicitUid,
  remoteResourceUrl
}: {
  kind: CalendarEventKind;
  pageName: string;
  parsed: ParsedDateTime;
  scopePropertyKey?: string;
  scopePropertyValue?: string;
  sourceBlockUuid?: string;
  sourceTitle?: string;
  sourceText?: string;
  propertyKey?: string;
  explicitUid?: string;
  remoteResourceUrl?: string;
}): LogseqCalendarEvent {
  const title = sourceTitle || buildEventTitle(kind, pageName);
  const description = [
    `Page: ${pageName}`,
    propertyKey ? `Property: ${propertyKey}` : "",
    sourceBlockUuid ? `Block UUID: ${sourceBlockUuid}` : "",
    sourceText || ""
  ]
    .filter(Boolean)
    .join("\n");
  const uidSource = [kind, pageName, parsed.date, parsed.time || "all-day", sourceBlockUuid || sourceTitle || ""]
    .map((value) => slugify(String(value)))
    .filter(Boolean)
    .join("-");

  return {
    uid: explicitUid || `logseq-event-${uidSource || "event"}@logseq.local`,
    kind,
    pageName,
    title,
    description,
    date: parsed.date,
    time: parsed.time,
    allDay: parsed.allDay,
    scopePropertyKey,
    scopePropertyValue,
    sourceBlockUuid,
    sourceBlockContent: sourceText,
    remoteResourceUrl
  };
}

function buildEventTitle(kind: CalendarEventKind, pageName: string) {
  switch (kind) {
    case "deadline":
      return `${pageName} deadline`;
    case "start":
      return `${pageName} starts`;
    case "scheduled":
    case "date":
    default:
      return pageName;
  }
}

function buildVEvent(event: LogseqCalendarEvent, timezone: string) {
  const uid = escapeIcsText(event.uid);
  const summary = escapeIcsText(event.title);
  const description = escapeIcsText(event.description || event.pageName);
  const dtstamp = formatUtcDateTime(new Date());
  const lines = [
    "BEGIN:VEVENT",
    `UID:${uid}`,
    `DTSTAMP:${dtstamp}`,
    `SUMMARY:${summary}`,
    `DESCRIPTION:${description}`,
    `CATEGORIES:${escapeIcsText(event.kind)}`,
    `X-LOGSEQ-PAGE:${escapeIcsText(event.pageName)}`,
    event.sourceBlockUuid ? `X-LOGSEQ-BLOCK-UUID:${escapeIcsText(event.sourceBlockUuid)}` : "",
    "END:VEVENT"
  ];

  if (event.allDay) {
    const endDate = shiftDateKey(event.date, 1);
    lines.splice(5, 0, `DTSTART;VALUE=DATE:${event.date}`, `DTEND;VALUE=DATE:${endDate}`);
  } else {
    const start = `${event.date}T${event.time ?? "090000"}`;
    const end = shiftDateTimeKey(event.date, event.time ?? "090000", DEFAULT_EVENT_DURATION_MINUTES);
    lines.splice(
      5,
      0,
      `DTSTART;TZID=${escapeIcsText(timezone || "Europe/Stockholm")}:${start}`,
      `DTEND;TZID=${escapeIcsText(timezone || "Europe/Stockholm")}:${end}`
    );
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

async function fetchRemoteCalendarResourceUrls(calendarUrl: string, authHeader: string) {
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
    <c:calendar-data />
  </d:prop>
  <c:filter>
    <c:comp-filter name="VCALENDAR">
      <c:comp-filter name="VEVENT" />
    </c:comp-filter>
  </c:filter>
</c:calendar-query>`,
    credentials: "include"
  });
  const responseText = await response.text();
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}${responseText ? `: ${responseText.slice(0, 500)}` : ""}`);
  }

  const byUid = new Map<string, string>();
  const doc = new DOMParser().parseFromString(responseText, "application/xml");
  const responses = Array.from(doc.getElementsByTagNameNS("*", "response"));

  for (const responseNode of responses) {
    const href = getNodeText(responseNode, "href");
    const resourceUrl = href ? resolveUrlFromHref(calendarUrl, href) : "";
    if (!resourceUrl) continue;
    const calendarData = getNodeText(responseNode, "calendar-data");
    const blocks = extractIcsComponentBlocks(unfoldIcsText(String(calendarData ?? "")), "VEVENT");
    for (const blockText of blocks) {
      const props = parseIcsProperties(blockText);
      const uid = String(props["UID"] || "").trim();
      if (uid && !byUid.has(uid)) {
        byUid.set(uid, resourceUrl);
      }
    }
  }

  return byUid;
}

function buildResourceUrl(calendarUrl: string, uid: string) {
  const base = calendarUrl.replace(/\/+$/, "");
  return `${base}/${encodeURIComponent(uid)}.ics`;
}

function resolveDavRootUrl(settings: NextcloudSyncSettings) {
  const explicit = normalizeDavRootUrl(settings.nextcloudDavUrl);
  if (explicit) return explicit;

  const source = settings.caldavCalendarUrl || settings.caldavTaskListUrl;
  if (!source) return "";

  try {
    const url = new URL(source);
    return `${url.origin}/remote.php/dav`;
  } catch {
    return "";
  }
}

function normalizeDavRootUrl(input: string) {
  const raw = String(input ?? "").trim();
  if (!raw) return "";
  const url = new URL(raw);
  url.hash = "";
  url.search = "";
  return stripTrailingSlash(url.toString());
}

function resolveUrlFromHref(baseUrl: string, href: string) {
  return new URL(String(href || "").trim(), baseUrl).toString();
}

function stripTrailingSlash(input: string) {
  return String(input ?? "").replace(/\/+$/, "");
}

function normalizeCalendarCollectionUrl(input: string) {
  const raw = String(input ?? "").trim();
  if (!raw) return "";

  const url = new URL(raw);
  url.hash = "";
  url.search = "";
  const pathname = url.pathname.replace(/\/+$/, "");
  if (!CALENDAR_COLLECTION_RE.test(pathname)) return "";
  return `${url.origin}${pathname}`;
}

export async function fetchRemoteCalendarEventsForImport(
  settings: NextcloudSyncSettings,
  calendarCollectionUrl?: string
): Promise<RemoteCalendarImportItem[]> {
  const calendarUrl = normalizeCalendarCollectionUrl(calendarCollectionUrl || settings.caldavCalendarUrl);
  if (!calendarUrl) {
    throw new Error("Set the exact Nextcloud remote collection URL first.");
  }
  if (!settings.caldavUsername || !settings.caldavPassword) {
    throw new Error("Set your Nextcloud username and app password in plugin settings first.");
  }

  const authHeader = `Basic ${btoa(`${settings.caldavUsername}:${settings.caldavPassword}`)}`;
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
    <c:calendar-data />
  </d:prop>
  <c:filter>
    <c:comp-filter name="VCALENDAR">
      <c:comp-filter name="VEVENT" />
    </c:comp-filter>
  </c:filter>
</c:calendar-query>`,
    credentials: "include"
  });

  const responseText = await response.text();
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}${responseText ? `: ${responseText.slice(0, 500)}` : ""}`);
  }

  const doc = new DOMParser().parseFromString(responseText, "application/xml");
  const responses = Array.from(doc.getElementsByTagNameNS("*", "response"));
  const items: RemoteCalendarImportItem[] = [];

  for (const responseNode of responses) {
    const href = getNodeText(responseNode, "href");
    const remoteResourceUrl = href ? resolveUrlFromHref(calendarUrl, href) : "";
    const calendarData = getNodeText(responseNode, "calendar-data");
    const blocks = extractIcsComponentBlocks(unfoldIcsText(String(calendarData ?? "")), "VEVENT");
    for (const blockText of blocks) {
      const props = parseIcsProperties(blockText);
      const start = parseIcsDateValue(readIcsPropertyValue(blockText, "DTSTART"));
      if (!props["UID"] || !start) continue;
      const end = parseIcsDateValue(readIcsPropertyValue(blockText, "DTEND"));
      items.push({
        uid: props["UID"],
        title: props["SUMMARY"] || "Untitled event",
        description: props["DESCRIPTION"] || "",
        date: start.date,
        time: start.time,
        endDate: end?.date,
        endTime: end?.time,
        allDay: start.allDay,
        sourcePage: props["X-LOGSEQ-PAGE"] || "",
        sourceBlockUuid: props["X-LOGSEQ-BLOCK-UUID"] || "",
        remoteResourceUrl
      });
    }
  }

  return items;
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
    date: `${pad(year, 4)}${pad(month)}${pad(day)}`,
    time: hasTime ? `${pad(hour)}${pad(minute)}00` : undefined,
    sortKey: new Date(year, month - 1, day, hour, minute, 0, 0).getTime()
  };
}

function getPropertyValue(properties: Record<string, any>, key: string) {
  const normalizedKey = normalizeKey(key);
  for (const [propKey, propValue] of Object.entries(properties)) {
    if (normalizeKey(propKey) !== normalizedKey) continue;
    if (Array.isArray(propValue)) {
      const first = propValue.find((value) => value != null && String(value).trim());
      return first ?? "";
    }
    return propValue;
  }
  return "";
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

function readIcsPropertyValue(componentText: string, key: string) {
  const pattern = new RegExp(`^${key}(?:;[^:]*)?:(.+)$`, "im");
  const match = unfoldIcsText(componentText).match(pattern);
  return match?.[1]?.trim() || "";
}

function parseIcsDateValue(value: string): ParsedDateTime | null {
  const raw = String(value || "").trim();
  if (!raw) return null;

  const dateOnly = raw.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (dateOnly) {
    const year = Number(dateOnly[1]);
    const month = Number(dateOnly[2]);
    const day = Number(dateOnly[3]);
    return {
      year,
      month,
      day,
      hour: 0,
      minute: 0,
      allDay: true,
      date: `${pad(year, 4)}${pad(month)}${pad(day)}`,
      sortKey: new Date(year, month - 1, day, 0, 0, 0, 0).getTime()
    };
  }

  const dateTime = raw.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})\d{0,2}Z?$/);
  if (!dateTime) return null;
  const year = Number(dateTime[1]);
  const month = Number(dateTime[2]);
  const day = Number(dateTime[3]);
  const hour = Number(dateTime[4]);
  const minute = Number(dateTime[5]);
  return {
    year,
    month,
    day,
    hour,
    minute,
    allDay: false,
    date: `${pad(year, 4)}${pad(month)}${pad(day)}`,
    time: `${pad(hour)}${pad(minute)}00`,
    sortKey: new Date(year, month - 1, day, hour, minute, 0, 0).getTime()
  };
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

function getFirstTagText(doc: Document, childLocalName: string, parentLocalName: string) {
  const parent = Array.from(doc.getElementsByTagNameNS("*", parentLocalName))[0];
  if (!parent) return "";
  const child = Array.from(parent.getElementsByTagNameNS("*", childLocalName))[0];
  return child?.textContent?.trim() ?? "";
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

function normalizeProperties(input: Record<string, any> | undefined | null) {
  const output: Record<string, any> = {};
  if (!input) return output;
  for (const [key, value] of Object.entries(input)) {
    output[normalizeKey(key)] = value;
  }
  return output;
}

function normalizeKey(input: string) {
  return String(input).trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function getExplicitProfileId(properties: Record<string, any>) {
  const raw =
    properties["nextcloud-profile"] ??
    properties["nextcloud-profile-id"] ??
    properties["nextcloudprofile"] ??
    properties["nextcloudprofileid"];
  return normalizeText(String(Array.isArray(raw) ? raw[0] ?? "" : raw ?? ""));
}

function getRemoteSyncUid(properties: Record<string, any>) {
  const raw =
    properties["nextcloud-remote-uid"] ??
    properties["nextcloud_remote_uid"] ??
    properties["nextcloudremoteuid"];
  return String(Array.isArray(raw) ? raw[0] ?? "" : raw ?? "").trim();
}

function getRemoteResourceUrl(properties: Record<string, any>) {
  const raw =
    properties["nextcloud-resource-url"] ??
    properties["nextcloud_resource_url"] ??
    properties["nextcloudresourceurl"];
  return String(Array.isArray(raw) ? raw[0] ?? "" : raw ?? "").trim();
}

function getPageType(page: any) {
  const properties = normalizeProperties(page?.properties);
  const explicitType = getPageTypeFromProperties(properties);
  if (explicitType) return explicitType;
  return page?.journal === true ? "journal" : "";
}

function getPageTags(page: any) {
  const properties = normalizeProperties(page?.properties);
  const tags = getTagsFromProperties(properties);
  if (page?.journal === true) {
    return mergeNormalizedTags(tags, ["daily", "journal"]);
  }
  return tags;
}

function getPageTypeFromProperties(properties: Record<string, any>) {
  const values = [properties["page-type"], properties.pagetype, properties.type, properties.kind].filter(Boolean);
  return normalizeText(String(values[0] ?? ""));
}

function getTagsFromProperties(properties: Record<string, any>) {
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

function readPropertyChildren(block: any) {
  const properties: Record<string, any> = {};
  const children = Array.isArray(block?.children) ? block.children : [];

  for (const child of children) {
    const content = String(child?.content || "").trim();
    const match = content.match(/^([A-Za-z0-9_-]+)::\s*(.*)$/);
    if (!match) continue;
    properties[match[1]] = match[2];
  }

  return properties;
}

function isPropertyOnlyBlock(content: string) {
  return /^[A-Za-z0-9_-]+::\s*.*$/.test(String(content ?? "").trim());
}

function isJournalLikePage(page: any, pageType: string, pageTags: string[]) {
  if (page?.journal === true) return true;
  const normalizedType = normalizeText(pageType);
  const normalizedTags = pageTags.map((tag) => normalizeText(tag)).filter(Boolean);
  return normalizedType === "daily" || normalizedType === "journal" || normalizedTags.includes("daily") || normalizedTags.includes("journal");
}

function normalizeText(input: string) {
  return String(input ?? "").trim().toLowerCase();
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

function getInlineTags(content: string) {
  const matches = Array.from(String(content ?? "").matchAll(/(^|\s)#([A-Za-z0-9/_-]+)/g));
  const unique = new Set(matches.map((match) => normalizeText(match[2] || "")).filter(Boolean));
  return Array.from(unique);
}

function mergeNormalizedTags(...tagGroups: string[][]) {
  const unique = new Set(
    tagGroups.flatMap((tags) => tags.map((tag) => normalizeText(tag)).filter(Boolean))
  );
  return Array.from(unique);
}

function matchCalendarScope(
  pageType: string,
  pageTags: string[],
  allScopeTags: string[],
  properties: Record<string, any>,
  scope?: Pick<CalendarScopeConfig, "id" | "propertyKey" | "propertyValue" | "filterPageTypes" | "filterTags" | "ignoredTags" | "prefilterPagesOnly"> | null
) {
  const propertyKey = normalizeKey(String(scope?.propertyKey || "calendar"));
  if (hasIgnoredScopeTags(allScopeTags, scope)) {
    return {
      matches: false,
      propertyKey,
      propertyValue: ""
    };
  }

  const explicitProfileId = getExplicitProfileId(properties);
  const explicitProfileMatch = explicitProfileId && explicitProfileId === normalizeText(scope?.id || "");
  if (explicitProfileMatch) {
    return {
      matches: true,
      propertyKey,
      propertyValue: ""
    };
  }

  const filterMatch = pageMatchesCalendarScope(pageType, pageTags, scope);
  const propertyValue = normalizeText(String(scope?.propertyValue || ""));
  if (!propertyValue) {
    return {
      matches: filterMatch,
      propertyKey,
      propertyValue: ""
    };
  }

  const rawValue = getPropertyValue(properties, propertyKey);
  const values = normalizeToStringList(rawValue).map((value) => normalizeText(value)).filter(Boolean);
  const matchedValue = values.find((value) => value === propertyValue) || "";
  return {
    matches: filterMatch || Boolean(matchedValue),
    propertyKey,
    propertyValue: matchedValue
  };
}

function pageMatchesCalendarScope(
  pageType: string,
  pageTags: string[],
  scope?: Pick<CalendarScopeConfig, "filterPageTypes" | "filterTags"> | null
) {
  const allowedPageTypes = parseFilterList(scope?.filterPageTypes || "");
  const allowedTags = parseFilterList(scope?.filterTags || "");
  if (!allowedPageTypes.length && !allowedTags.length) return false;

  const hasPageTypeMatch = allowedPageTypes.length ? allowedPageTypes.includes(normalizeText(pageType)) : false;
  const normalizedTags = pageTags.map((tag) => normalizeText(tag)).filter(Boolean);
  const hasTagMatch = allowedTags.length ? allowedTags.some((tag) => normalizedTags.includes(tag)) : false;
  return hasPageTypeMatch || hasTagMatch;
}

function hasIgnoredScopeTags(
  pageTags: string[],
  scope?: Pick<CalendarScopeConfig, "ignoredTags"> | null
) {
  const ignoredTags = parseFilterList(scope?.ignoredTags || "");
  if (!ignoredTags.length) return false;
  const normalizedTags = pageTags.map((tag) => normalizeText(tag)).filter(Boolean);
  return ignoredTags.some((tag) => normalizedTags.includes(tag));
}

function pageContainsExplicitProfileBlocks(blocks: any[], profileId: string) {
  const target = normalizeText(profileId || "");
  if (!target) return false;

  const visit = (block: any): boolean => {
    const properties = normalizeProperties({
      ...(block?.properties ?? {}),
      ...(block?.meta?.properties ?? {}),
      ...readPropertyChildren(block)
    });
    if (getExplicitProfileId(properties) === target) return true;
    if (Array.isArray(block?.children)) {
      return block.children.some((child: any) => visit(child));
    }
    return false;
  };

  return blocks.some((block) => visit(block));
}

function parseFilterList(input: string) {
  return String(input ?? "")
    .split(",")
    .map((item) => normalizeText(item))
    .filter(Boolean);
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

function isTaskBlock(content: string) {
  return /^\s*(TODO|DOING|DONE|LATER|WAITING|NOW|CANCELLED)\b/i.test(String(content ?? ""));
}

function isSyncConflictPage(pageName: string) {
  return /\.sync-conflict-/i.test(String(pageName ?? ""));
}

function compareEvents(a: LogseqCalendarEvent, b: LogseqCalendarEvent) {
  const aKey = eventSortKey(a);
  const bKey = eventSortKey(b);
  return aKey - bKey || a.title.localeCompare(b.title) || a.uid.localeCompare(b.uid);
}

async function collectImportedMarkdownEvents(
  scope?: Pick<CalendarScopeConfig, "id"> | null
): Promise<LogseqCalendarEvent[]> {
  const profileId = normalizeText(scope?.id || "");
  if (!profileId) return [];

  try {
    const graphPath = await resolveGraphRootPath();
    const req = (globalThis as any).require || (globalThis as any).window?.require || (globalThis as any).top?.require;
    if (!graphPath || typeof req !== "function") return [];
    const fs = req("fs") as any;
    const path = req("path") as any;

    const directories = [path.join(graphPath, "journals"), path.join(graphPath, "pages")];
    const events: LogseqCalendarEvent[] = [];

    for (const directory of directories) {
      const exists = fs.existsSync(directory);
      const entries = exists ? fs.readdirSync(directory) : [];
      if (!exists) continue;
      for (const entry of entries) {
        if (!String(entry).endsWith(".md")) continue;
        const filePath = path.join(directory, entry);
        const text = String(fs.readFileSync(filePath, "utf8"));
        events.push(...parseImportedMarkdownEvents(text, entry, profileId));
      }
    }

    return events;
  } catch (error) {
    return [];
  }
}

function parseImportedMarkdownEvents(text: string, fileName: string, profileId: string) {
  const lines = String(text || "").split(/\r?\n/);
  const events: LogseqCalendarEvent[] = [];
  let insideNextcloud = false;
  let insideProfile = false;
  let insideEvents = false;
  let current: { title: string; props: Record<string, string> } | null = null;

  const flush = () => {
    if (!current) return;
    const currentProfileId = normalizeText(current.props["nextcloud-profile"] || "");
    const uid = String(current.props["nextcloud-remote-uid"] || "").trim();
    const startValue = current.props.start || current.props.date || "";
    const parsed = parseDateTime(startValue);
    if (currentProfileId === profileId && uid && parsed) {
      events.push({
        uid,
        kind: current.props.start ? "start" : "date",
        pageName: fileName.replace(/\.md$/i, ""),
        title: current.title,
        description: `Imported markdown event from ${fileName}`,
        date: parsed.date,
        time: parsed.time,
        allDay: parsed.allDay,
        scopePropertyKey: "nextcloud-profile",
        scopePropertyValue: profileId,
        sourceBlockContent: current.title
      });
    }
    current = null;
  };

  for (const line of lines) {
    const bulletMatch = line.match(/^(\t*)-\s+(.*)$/);
    const propertyMatch = line.match(/^\t+\s*([A-Za-z0-9_-]+)::\s*(.*)$/);

    if (bulletMatch) {
      const depth = bulletMatch[1].length;
      const content = String(bulletMatch[2] || "").trim();

      if (depth <= 0) {
        flush();
        insideNextcloud = content === "Nextcloud";
        insideProfile = false;
        insideEvents = false;
        continue;
      }

      if (depth === 1) {
        flush();
        insideProfile = insideNextcloud;
        insideEvents = false;
        continue;
      }

      if (depth === 2) {
        flush();
        insideEvents = insideProfile && content === "Events";
        continue;
      }

      if (depth === 3) {
        flush();
        if (insideEvents) {
          current = { title: content, props: {} };
        }
        continue;
      }
    }

    if (propertyMatch && current) {
      current.props[normalizeKey(propertyMatch[1])] = String(propertyMatch[2] || "").trim();
    }
  }

  flush();
  return events;
}

async function resolveGraphRootPath() {
  const explicit = String(logseq?.settings?.graphRootPath || "").trim();
  if (explicit) return explicit;
  try {
    const graph = await logseq.App.getCurrentGraph?.();
    return String(graph?.path || "").trim();
  } catch {
    return "";
  }
}

function eventSortKey(event: LogseqCalendarEvent) {
  const year = Number(event.date.slice(0, 4));
  const month = Number(event.date.slice(4, 6));
  const day = Number(event.date.slice(6, 8));
  const hour = event.allDay || !event.time ? 0 : Number(event.time.slice(0, 2));
  const minute = event.allDay || !event.time ? 0 : Number(event.time.slice(2, 4));
  return new Date(year, month - 1, day, hour, minute, 0, 0).getTime();
}

function shiftDateKey(date: string, days: number) {
  const year = Number(date.slice(0, 4));
  const month = Number(date.slice(4, 6));
  const day = Number(date.slice(6, 8));
  const shifted = new Date(year, month - 1, day);
  shifted.setDate(shifted.getDate() + days);
  return `${pad(shifted.getFullYear(), 4)}${pad(shifted.getMonth() + 1)}${pad(shifted.getDate())}`;
}

function shiftDateTimeKey(date: string, time: string, minutes: number) {
  const year = Number(date.slice(0, 4));
  const month = Number(date.slice(4, 6));
  const day = Number(date.slice(6, 8));
  const hour = Number(time.slice(0, 2));
  const minute = Number(time.slice(2, 4));
  const shifted = new Date(year, month - 1, day, hour, minute, 0, 0);
  shifted.setMinutes(shifted.getMinutes() + minutes);
  return `${pad(shifted.getFullYear(), 4)}${pad(shifted.getMonth() + 1)}${pad(shifted.getDate())}T${pad(shifted.getHours())}${pad(shifted.getMinutes())}00`;
}

function formatUtcDateTime(date: Date) {
  return [
    pad(date.getUTCFullYear(), 4),
    pad(date.getUTCMonth() + 1),
    pad(date.getUTCDate())
  ].join("") + `T${pad(date.getUTCHours())}${pad(date.getUTCMinutes())}${pad(date.getUTCSeconds())}Z`;
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

function escapeIcsText(input: string) {
  return String(input ?? "")
    .replace(/\\/g, "\\\\")
    .replace(/\r?\n/g, "\\n")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,");
}

function readCalendarSyncCache(calendarUrl: string): CalendarSyncCache {
  try {
    const raw = localStorage.getItem(CALENDAR_SYNC_KEY);
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

function writeCalendarSyncCache(cache: CalendarSyncCache) {
  localStorage.setItem(CALENDAR_SYNC_KEY, JSON.stringify(cache));
}

function formatRequestError(action: string, error: unknown) {
  if (error instanceof Error) return `${action}: ${error.message}`;
  return `${action}: ${String(error)}`;
}
