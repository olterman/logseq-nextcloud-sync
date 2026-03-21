import type { CalendarEventKind, DiscoveredCalendar, LogseqCalendarEvent, NextcloudSyncSettings } from "./types";

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
  start: ["start", "start-date", "start-time"],
  end: ["end", "end-date", "end-time"]
};

export async function collectLogseqCalendarEvents(): Promise<LogseqCalendarEvent[]> {
  const pages = (await logseq.Editor.getAllPages?.()) ?? [];
  const events = new Map<string, LogseqCalendarEvent>();

  for (const page of pages) {
    const pageName = String(page?.name ?? page?.originalName ?? "").trim();
    if (!pageName) continue;

    const pageProperties = normalizeProperties(page?.properties);
    const pageEvents = extractPropertyEvents(pageName, pageProperties);
    for (const event of pageEvents) {
      events.set(event.uid, event);
    }

    const blocks = (await logseq.Editor.getPageBlocksTree?.(pageName)) ?? [];
    for (const block of blocks) {
      collectBlockEventsRecursive(pageName, block, events);
    }
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

export async function syncCalendarToCalDav(events: LogseqCalendarEvent[], settings: NextcloudSyncSettings): Promise<CalendarSyncResult> {
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

  for (const event of events) {
    const ics = buildCalendarIcs([event], settings.calendarTimezone);
    const targetUrl = buildResourceUrl(calendarUrl, event.uid);

    try {
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
    const targetUrl = buildResourceUrl(calendarUrl, uid);
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
  const calendarUrl = normalizeCalendarCollectionUrl(settings.caldavCalendarUrl);
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

function collectBlockEventsRecursive(pageName: string, block: any, events: Map<string, LogseqCalendarEvent>) {
  if (!block || typeof block !== "object") return;

  const blockProperties = normalizeProperties({
    ...(block.properties ?? {}),
    ...(block.meta?.properties ?? {})
  });
  const content = typeof block.content === "string" ? block.content : "";
  const title = getBlockTitle(content) || pageName;
  const blockUuid = typeof block.uuid === "string" ? block.uuid : undefined;

  for (const event of extractPropertyEvents(pageName, blockProperties, blockUuid, title, content)) {
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
    const event = makeEvent({
      kind,
      pageName,
      parsed,
      sourceBlockUuid: blockUuid,
      sourceTitle: title,
      sourceText: content
    });
    events.set(event.uid, event);
  }

  if (Array.isArray(block.children)) {
    for (const child of block.children) {
      collectBlockEventsRecursive(pageName, child, events);
    }
  }
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
  sourceBlockUuid?: string,
  sourceTitle?: string,
  sourceText?: string
) {
  const events: LogseqCalendarEvent[] = [];

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
          sourceBlockUuid,
          sourceTitle,
          sourceText,
          propertyKey: key
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
  sourceBlockUuid,
  sourceTitle,
  sourceText,
  propertyKey
}: {
  kind: CalendarEventKind;
  pageName: string;
  parsed: ParsedDateTime;
  sourceBlockUuid?: string;
  sourceTitle?: string;
  sourceText?: string;
  propertyKey?: string;
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
    uid: `logseq-event-${uidSource || "event"}@logseq.local`,
    kind,
    pageName,
    title,
    description,
    date: parsed.date,
    time: parsed.time,
    allDay: parsed.allDay,
    sourceBlockUuid,
    sourceBlockContent: sourceText
  };
}

function buildEventTitle(kind: CalendarEventKind, pageName: string) {
  switch (kind) {
    case "deadline":
      return `${pageName} deadline`;
    case "start":
      return `${pageName} starts`;
    case "end":
      return `${pageName} ends`;
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
  return String(input).trim().toLowerCase().replace(/\s+/g, "-");
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

function compareEvents(a: LogseqCalendarEvent, b: LogseqCalendarEvent) {
  const aKey = eventSortKey(a);
  const bKey = eventSortKey(b);
  return aKey - bKey || a.title.localeCompare(b.title) || a.uid.localeCompare(b.uid);
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
