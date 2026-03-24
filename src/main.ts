import "@logseq/libs";
import {
  collectLogseqTasksForScope,
  collectLogseqTasksWithSettings,
  createCalDavTaskList,
  discoverCalDavTaskLists,
  exportTaskListIcs,
  fetchRemoteTasksForImport,
  syncTasksToCalDav,
  testCalDavConnection,
  testCalDavConnectionForUrl
} from "./caldav";
import type { RemoteTaskImportItem } from "./caldav";
import {
  collectLogseqCalendarEvents,
  collectLogseqCalendarEventsForScope,
  discoverCalDavCalendars,
  exportCalendarIcs,
  fetchRemoteCalendarEventsForImport,
  syncCalendarToCalDav,
  testCalendarConnectionForUrl
} from "./calendar";
import type { CalendarDiscoveryResult, RemoteCalendarImportItem } from "./calendar";
import type {
  CalendarScopeConfig,
  DiscoveredTaskList,
  LogseqCalendarEvent,
  LogseqTaskItem,
  NextcloudSyncSettings,
  SyncProfileConfig,
  TaskScopeConfig
} from "./types";

declare const logseq: any;

type PluginRuntimeState = {
  initialized: boolean;
  commandLabels: Set<string>;
  slashLabels: Set<string>;
  profileSyncTimers: Map<string, number>;
  syncingProfileIds: Set<string>;
};

const runtimeState: PluginRuntimeState = ((globalThis as any).__logseqNextcloudSyncRuntime ??= {
  initialized: false,
  commandLabels: new Set<string>(),
  slashLabels: new Set<string>(),
  profileSyncTimers: new Map<string, number>(),
  syncingProfileIds: new Set<string>()
}) as PluginRuntimeState;

const defaultSettings: NextcloudSyncSettings = {
  syncOnStartup: true,
  calendarTimezone: "Europe/Stockholm",
  nextcloudDavUrl: "",
  graphRootPath: "",
  syncProfilesJson: "",
  importedCalendarUidCacheJson: "",
  calendarSyncStateJson: "",
  activeSyncProfileId: "",
  caldavTaskListUrl: "",
  caldavCalendarUrl: "",
  calendarScopesJson: "",
  activeCalendarScopeId: "",
  caldavUsername: "",
  caldavPassword: "",
  taskPageName: "Nextcloud Tasks",
  calendarPageName: "Nextcloud Calendar",
  taskFilterPageTypes: "",
  taskFilterTags: "",
  taskIgnoreTags: "",
  prefilterPagesOnly: false,
  taskScopesJson: "",
  activeTaskScopeId: ""
};

const state = {
  tasks: [] as LogseqTaskItem[],
  events: [] as LogseqCalendarEvent[],
  activeTaskProfileId: "",
  activeCalendarProfileId: "",
  settingsOverride: {} as Partial<NextcloudSyncSettings>,
  calendarDiscovery: null as CalendarDiscoveryResult | null,
  taskListDiscovery: null as TaskListDiscoveryResult | null
};

const uiState = {
  calendarSelectorSlot: "",
  taskScopeManagerSlot: "",
  syncHubSlot: "",
  profileEditorSlot: ""
};

const uiKeys = {
  calendarSelector: "nextcloud-calendar-selector",
  taskScopeManager: "nextcloud-task-scope-manager",
  syncHub: "nextcloud-sync-hub",
  profileEditor: "nextcloud-profile-editor"
} as const;

function slotUiKey(baseKey: string, slot: string): string {
  return slot ? `${baseKey}-${slot}` : baseKey;
}

const syncHubPage = "Nextcloud Sync";
const calendarSelectorPage = "Nextcloud Calendar Picker";
const taskScopeManagerPage = "Nextcloud Sync Profiles";
const profileEditorPage = "Nextcloud Profile Editor";

type TaskListDiscoveryResult = Awaited<ReturnType<typeof discoverCalDavTaskLists>>;

type CalendarSyncSnapshot = {
  title: string;
  date: string;
  time: string;
  endDate: string;
  endTime: string;
  allDay: boolean;
  remoteResourceUrl: string;
};

function settings(): NextcloudSyncSettings {
  return { ...defaultSettings, ...(logseq?.settings ?? {}), ...state.settingsOverride };
}

function registerCommand(label: string, handler: () => Promise<void> | void) {
  if (runtimeState.commandLabels.has(label)) return;
  if (typeof logseq.App?.registerCommandPalette === "function") {
    logseq.App.registerCommandPalette({ key: label, label }, handler);
    runtimeState.commandLabels.add(label);
  }
}

function registerSlashCommand(label: string, handler: () => Promise<void> | void) {
  if (runtimeState.slashLabels.has(label)) return;
  const editor = logseq.Editor as any;
  if (typeof editor.registerSlashCommand === "function") {
    editor.registerSlashCommand(label, handler);
    runtimeState.slashLabels.add(label);
  }
}

function configureSettings() {
  if (typeof logseq.useSettingsSchema !== "function") return;

  logseq.useSettingsSchema([
    {
      key: "syncOnStartup",
      title: "Sync tasks to Nextcloud on startup",
      type: "boolean",
      default: true
    },
    {
      key: "calendarTimezone",
      title: "Task timezone",
      type: "string",
      default: "Europe/Stockholm"
    },
    {
      key: "nextcloudDavUrl",
      title: "Nextcloud DAV root URL",
      description: "Usually https://host/remote.php/dav",
      type: "string",
      default: ""
    },
    {
      key: "graphRootPath",
      title: "Logseq graph root path",
      description: "Optional fallback absolute path to your graph root, for example /mnt/Storage/Nextcloud/Notes/logseq",
      type: "string",
      default: ""
    },
    {
      key: "syncProfilesJson",
      title: "Sync profiles JSON",
      description: "Managed by the sync profiles UI.",
      type: "string",
      default: ""
    },
    {
      key: "calendarSyncStateJson",
      title: "Calendar sync state JSON",
      description: "Managed by the plugin.",
      type: "string",
      default: ""
    },
    {
      key: "activeSyncProfileId",
      title: "Active sync profile id",
      description: "Managed by the sync profiles UI.",
      type: "string",
      default: ""
    },
    {
      key: "caldavTaskListUrl",
      title: "Nextcloud task list URL",
      description: "Paste the exact task list collection URL from Nextcloud.",
      type: "string",
      default: ""
    },
    {
      key: "caldavCalendarUrl",
      title: "Nextcloud calendar URL",
      description: "Paste the exact Nextcloud calendar collection URL.",
      type: "string",
      default: ""
    },
    {
      key: "calendarScopesJson",
      title: "Calendar scopes JSON",
      description: "Managed by the calendar scope UI.",
      type: "string",
      default: ""
    },
    {
      key: "activeCalendarScopeId",
      title: "Active calendar scope id",
      description: "Managed by the calendar scope UI.",
      type: "string",
      default: ""
    },
    {
      key: "caldavUsername",
      title: "Nextcloud username",
      type: "string",
      default: ""
    },
    {
      key: "caldavPassword",
      title: "Nextcloud app password",
      type: "string",
      inputAs: "password",
      default: ""
    },
    {
      key: "taskPageName",
      title: "Task overview page",
      type: "string",
      default: "Nextcloud Tasks"
    },
    {
      key: "calendarPageName",
      title: "Calendar overview page",
      type: "string",
      default: "Nextcloud Calendar"
    },
    {
      key: "taskFilterPageTypes",
      title: "Task filter page types",
      description: "Comma-separated page-type values to include in task sync. Empty means no page-type filter.",
      type: "string",
      default: ""
    },
    {
      key: "taskFilterTags",
      title: "Task filter tags",
      description: "Comma-separated page tags to include in task sync. A page is included if it matches a listed page type or tag.",
      type: "string",
      default: ""
    },
    {
      key: "taskIgnoreTags",
      title: "Task ignore tags",
      description: "Comma-separated tags that always exclude a page or block from sync.",
      type: "string",
      default: ""
    },
    {
      key: "taskScopesJson",
      title: "Task scopes JSON",
      description: "Managed by the task scope UI.",
      type: "string",
      default: ""
    },
    {
      key: "activeTaskScopeId",
      title: "Active task scope id",
      description: "Managed by the task scope UI.",
      type: "string",
      default: ""
    }
  ]);
}

function promptValue(message: string, defaultValue = "") {
  if (typeof window?.prompt !== "function") return defaultValue;
  return String(window.prompt(message, defaultValue) ?? "").trim();
}

function openPage(pageName: string) {
  logseq.App.pushState?.("page", { name: pageName });
}

function escapeHtml(input: string) {
  return String(input ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function slugify(input: string) {
  return String(input ?? "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function createScopeId(name: string) {
  return `scope-${slugify(name) || Date.now()}`;
}

function buildLegacyTaskScope(current: NextcloudSyncSettings): TaskScopeConfig {
  return {
    id: "legacy-default",
    name: "Default task scope",
    taskListUrl: current.caldavTaskListUrl,
    filterPageTypes: current.taskFilterPageTypes,
    filterTags: current.taskFilterTags,
    ignoredTags: current.taskIgnoreTags,
    prefilterPagesOnly: false,
    enabled: true
  };
}

function buildLegacyCalendarScope(current: NextcloudSyncSettings): CalendarScopeConfig {
  return {
    id: "legacy-calendar-default",
    name: "Default calendar scope",
    calendarUrl: current.caldavCalendarUrl,
    propertyKey: "calendar",
    propertyValue: "",
    enabled: true
  };
}

function buildLegacySyncProfile(current: NextcloudSyncSettings): SyncProfileConfig {
  const remoteUrl = current.caldavTaskListUrl || current.caldavCalendarUrl;
  return {
    id: "legacy-default",
    name: "Default sync profile",
    remoteUrl,
    taskListUrl: remoteUrl,
    filterPageTypes: current.taskFilterPageTypes,
    filterTags: current.taskFilterTags,
    ignoredTags: current.taskIgnoreTags,
    prefilterPagesOnly: false,
    calendarUrl: remoteUrl,
    propertyKey: "calendar",
    propertyValue: "",
    writeToJournal: false,
    defaultImportPage: "",
    simpleChecklistMode: false,
    updatePeriodically: false,
    updateIntervalMinutes: 15,
    enabled: true
  };
}

function normalizeProfileName(input: string) {
  return String(input || "").trim().toLowerCase();
}

function normalizeProfileDestinations<T extends SyncProfileConfig>(profile: T): T {
  const remoteUrl = String(profile.remoteUrl || profile.taskListUrl || profile.calendarUrl || "").trim();
  return {
    ...profile,
    remoteUrl,
    taskListUrl: remoteUrl,
    calendarUrl: remoteUrl,
    updatePeriodically: profile.updatePeriodically === true,
    updateIntervalMinutes: Math.max(1, Number(profile.updateIntervalMinutes || 15) || 15)
  };
}

function readLegacyTaskScopes(current: NextcloudSyncSettings) {
  try {
    const parsed = JSON.parse(String(current.taskScopesJson || "[]"));
    if (!Array.isArray(parsed)) return [];
    return parsed
      .map((scope) => ({
        id: String(scope?.id || "").trim(),
        name: String(scope?.name || "").trim(),
        taskListUrl: String(scope?.taskListUrl || "").trim(),
        filterPageTypes: String(scope?.filterPageTypes || "").trim(),
        filterTags: String(scope?.filterTags || "").trim(),
        ignoredTags: String(scope?.ignoredTags || "").trim(),
        prefilterPagesOnly: scope?.prefilterPagesOnly === true,
        enabled: scope?.enabled !== false
      }))
      .filter((scope) => scope.id && scope.name);
  } catch {
    return [];
  }
}

function readLegacyCalendarScopes(current: NextcloudSyncSettings) {
  try {
    const parsed = JSON.parse(String(current.calendarScopesJson || "[]"));
    if (!Array.isArray(parsed)) return [];
    return parsed
      .map((scope) => ({
        id: String(scope?.id || "").trim(),
        name: String(scope?.name || "").trim(),
        calendarUrl: String(scope?.calendarUrl || "").trim(),
        propertyKey: String(scope?.propertyKey || "calendar").trim() || "calendar",
        propertyValue: String(scope?.propertyValue || "").trim(),
        enabled: scope?.enabled !== false
      }))
      .filter((scope) => scope.id && scope.name);
  } catch {
    return [];
  }
}

function mergeLegacyScopesIntoProfiles(current: NextcloudSyncSettings): SyncProfileConfig[] {
  const taskScopes = readLegacyTaskScopes(current);
  const calendarScopes = readLegacyCalendarScopes(current);
  const byName = new Map<string, SyncProfileConfig>();

  for (const scope of taskScopes) {
    byName.set(normalizeProfileName(scope.name) || scope.id, normalizeProfileDestinations({
      id: scope.id,
      name: scope.name,
      remoteUrl: scope.taskListUrl,
      taskListUrl: scope.taskListUrl,
      filterPageTypes: scope.filterPageTypes,
      filterTags: scope.filterTags,
      ignoredTags: scope.ignoredTags,
      prefilterPagesOnly: scope.prefilterPagesOnly === true,
      calendarUrl: "",
      propertyKey: "calendar",
      propertyValue: "",
      enabled: scope.enabled
    }));
  }

  for (const scope of calendarScopes) {
    const key = normalizeProfileName(scope.name) || scope.id;
    const existing = byName.get(key);
    if (existing) {
      existing.calendarUrl = scope.calendarUrl;
      existing.propertyKey = scope.propertyKey;
      existing.propertyValue = scope.propertyValue;
      existing.writeToJournal = existing.writeToJournal === true;
      existing.defaultImportPage = existing.defaultImportPage || "";
      existing.simpleChecklistMode = existing.simpleChecklistMode === true;
      existing.updatePeriodically = existing.updatePeriodically === true;
      existing.updateIntervalMinutes = Math.max(1, Number(existing.updateIntervalMinutes || 15) || 15);
      existing.enabled = existing.enabled || scope.enabled;
      continue;
    }

    byName.set(key, normalizeProfileDestinations({
      id: scope.id,
      name: scope.name,
      remoteUrl: scope.calendarUrl,
      taskListUrl: "",
      filterPageTypes: "",
      filterTags: "",
      ignoredTags: "",
      prefilterPagesOnly: false,
      calendarUrl: scope.calendarUrl,
      propertyKey: scope.propertyKey,
      propertyValue: scope.propertyValue,
      writeToJournal: false,
      defaultImportPage: "",
      simpleChecklistMode: false,
      updatePeriodically: false,
      updateIntervalMinutes: 15,
      enabled: scope.enabled
    }));
  }

  return Array.from(byName.values()).filter((profile) => profile.id && profile.name);
}

function readStoredSyncProfiles(current: NextcloudSyncSettings) {
  try {
    const parsed = JSON.parse(String(current.syncProfilesJson || "[]"));
    if (Array.isArray(parsed) && parsed.length) {
      return parsed
        .map((profile) => normalizeProfileDestinations({
          id: String(profile?.id || "").trim(),
          name: String(profile?.name || "").trim(),
          remoteUrl: String(profile?.remoteUrl || profile?.taskListUrl || profile?.calendarUrl || "").trim(),
          taskListUrl: String(profile?.taskListUrl || profile?.remoteUrl || profile?.calendarUrl || "").trim(),
          filterPageTypes: String(profile?.filterPageTypes || "").trim(),
          filterTags: String(profile?.filterTags || "").trim(),
          ignoredTags: String(profile?.ignoredTags || "").trim(),
          prefilterPagesOnly: profile?.prefilterPagesOnly === true,
          calendarUrl: String(profile?.calendarUrl || profile?.remoteUrl || profile?.taskListUrl || "").trim(),
          propertyKey: String(profile?.propertyKey || "calendar").trim() || "calendar",
          propertyValue: String(profile?.propertyValue || "").trim(),
          writeToJournal: profile?.writeToJournal === true,
          defaultImportPage: String(profile?.defaultImportPage || "").trim(),
          simpleChecklistMode: profile?.simpleChecklistMode === true,
          updatePeriodically: profile?.updatePeriodically === true,
          updateIntervalMinutes: Math.max(1, Number(profile?.updateIntervalMinutes || 15) || 15),
          enabled: profile?.enabled !== false
        }))
        .filter((profile) => profile.id && profile.name);
    }
  } catch {
    // fall through to compatibility fallbacks
  }

  const merged = mergeLegacyScopesIntoProfiles(current);
  return merged.length ? merged : [];
}

function getSyncProfiles(current = settings()) {
  const stored = readStoredSyncProfiles(current);
  return stored.length ? stored : [buildLegacySyncProfile(current)];
}

function getPersistedSyncProfiles(current = settings()) {
  return readStoredSyncProfiles(current);
}

function getActiveSyncProfile(current = settings()) {
  const profiles = getSyncProfiles(current);
  if (!profiles.length) return null;
  const activeId = String(current.activeSyncProfileId || "").trim();
  if (activeId) {
    return profiles.find((profile) => profile.id === activeId) ?? profiles[0];
  }

  const legacyTaskId = String(current.activeTaskScopeId || "").trim();
  const legacyCalendarId = String(current.activeCalendarScopeId || "").trim();
  return profiles.find((profile) => profile.id === legacyTaskId || profile.id === legacyCalendarId) ?? profiles[0];
}

function getTaskScopes(current = settings()) {
  return getSyncProfiles(current).map((profile) => ({
    id: profile.id,
    name: profile.name,
    taskListUrl: profile.remoteUrl || profile.taskListUrl,
    filterPageTypes: profile.filterPageTypes,
    filterTags: profile.filterTags,
    ignoredTags: profile.ignoredTags,
    prefilterPagesOnly: profile.prefilterPagesOnly === true,
    enabled: profile.enabled
  }));
}

function getPersistedTaskScopes(current = settings()) {
  return getPersistedSyncProfiles(current).map((profile) => ({
    id: profile.id,
    name: profile.name,
    taskListUrl: profile.remoteUrl || profile.taskListUrl,
    filterPageTypes: profile.filterPageTypes,
    filterTags: profile.filterTags,
    ignoredTags: profile.ignoredTags,
    prefilterPagesOnly: profile.prefilterPagesOnly === true,
    enabled: profile.enabled
  }));
}

function getCalendarScopes(current = settings()) {
  return getSyncProfiles(current).map((profile) => ({
    id: profile.id,
    name: profile.name,
    calendarUrl: profile.remoteUrl || profile.calendarUrl,
    filterPageTypes: profile.filterPageTypes,
    filterTags: profile.filterTags,
    ignoredTags: profile.ignoredTags,
    prefilterPagesOnly: profile.prefilterPagesOnly === true,
    propertyKey: profile.propertyKey,
    propertyValue: profile.propertyValue,
    enabled: profile.enabled
  }));
}

function getPersistedCalendarScopes(current = settings()) {
  return getPersistedSyncProfiles(current).map((profile) => ({
    id: profile.id,
    name: profile.name,
    calendarUrl: profile.remoteUrl || profile.calendarUrl,
    filterPageTypes: profile.filterPageTypes,
    filterTags: profile.filterTags,
    ignoredTags: profile.ignoredTags,
    prefilterPagesOnly: profile.prefilterPagesOnly === true,
    propertyKey: profile.propertyKey,
    propertyValue: profile.propertyValue,
    enabled: profile.enabled
  }));
}

function getActiveTaskScope(current = settings()) {
  const profile = getActiveSyncProfile(current);
  return profile
    ? {
        id: profile.id,
        name: profile.name,
        taskListUrl: profile.remoteUrl || profile.taskListUrl,
        filterPageTypes: profile.filterPageTypes,
        filterTags: profile.filterTags,
        ignoredTags: profile.ignoredTags,
        prefilterPagesOnly: profile.prefilterPagesOnly === true,
        enabled: profile.enabled
      }
    : null;
}

function getActiveCalendarScope(current = settings()) {
  const profile = getActiveSyncProfile(current);
  return profile
    ? {
        id: profile.id,
        name: profile.name,
        calendarUrl: profile.remoteUrl || profile.calendarUrl,
        filterPageTypes: profile.filterPageTypes,
        filterTags: profile.filterTags,
        ignoredTags: profile.ignoredTags,
        prefilterPagesOnly: profile.prefilterPagesOnly === true,
        propertyKey: profile.propertyKey,
        propertyValue: profile.propertyValue,
        enabled: profile.enabled
      }
    : null;
}

function withScopeSettings(current: NextcloudSyncSettings, scope: TaskScopeConfig): NextcloudSyncSettings {
  return {
    ...current,
    caldavTaskListUrl: scope.taskListUrl,
    taskFilterPageTypes: scope.filterPageTypes,
    taskFilterTags: scope.filterTags,
    taskIgnoreTags: scope.ignoredTags,
    prefilterPagesOnly: scope.prefilterPagesOnly === true,
    simpleChecklistMode: false
  };
}

function withProfileSettings(current: NextcloudSyncSettings, profile: SyncProfileConfig): NextcloudSyncSettings {
  return {
    ...current,
    caldavTaskListUrl: profile.remoteUrl || profile.taskListUrl,
    taskFilterPageTypes: profile.filterPageTypes,
    taskFilterTags: profile.filterTags,
    taskIgnoreTags: profile.ignoredTags,
    prefilterPagesOnly: profile.prefilterPagesOnly === true,
    caldavCalendarUrl: profile.remoteUrl || profile.calendarUrl || profile.taskListUrl,
    simpleChecklistMode: profile.simpleChecklistMode === true
  };
}

async function saveSyncProfiles(profiles: SyncProfileConfig[], activeProfileId?: string) {
  const normalizedProfiles = profiles.map((profile) => normalizeProfileDestinations(profile));
  const nextActiveId = activeProfileId ?? profiles[0]?.id ?? "";
  state.settingsOverride = {
    ...state.settingsOverride,
    syncProfilesJson: JSON.stringify(normalizedProfiles),
    activeSyncProfileId: nextActiveId
  };
  await logseq.updateSettings?.({
    syncProfilesJson: JSON.stringify(normalizedProfiles),
    activeSyncProfileId: nextActiveId
  });
  refreshProfileAutoSyncSchedules();
}

function withCalendarScopeSettings(current: NextcloudSyncSettings, scope: CalendarScopeConfig): NextcloudSyncSettings {
  return {
    ...current,
    caldavCalendarUrl: scope.calendarUrl
  };
}

async function saveTaskScopes(scopes: TaskScopeConfig[], activeScopeId?: string) {
  const current = settings();
  const profiles = getPersistedSyncProfiles(current);
  const profileById = new Map(profiles.map((profile) => [profile.id, profile]));
  const nextProfiles = scopes.map((scope) => {
    const existing = profileById.get(scope.id);
    return {
      id: scope.id,
      name: scope.name,
      taskListUrl: scope.taskListUrl,
      filterPageTypes: scope.filterPageTypes,
      filterTags: scope.filterTags,
      ignoredTags: scope.ignoredTags,
      prefilterPagesOnly: scope.prefilterPagesOnly === true,
      remoteUrl: scope.taskListUrl,
      calendarUrl: existing?.calendarUrl || scope.taskListUrl || "",
      propertyKey: existing?.propertyKey || "calendar",
      propertyValue: existing?.propertyValue || "",
      writeToJournal: existing?.writeToJournal === true,
      defaultImportPage: existing?.defaultImportPage || "",
      simpleChecklistMode: existing?.simpleChecklistMode === true,
      updatePeriodically: existing?.updatePeriodically === true,
      updateIntervalMinutes: Math.max(1, Number(existing?.updateIntervalMinutes || 15) || 15),
      enabled: scope.enabled
    };
  });
  await saveSyncProfiles(nextProfiles, activeScopeId);
}

async function saveCalendarScopes(scopes: CalendarScopeConfig[], activeScopeId?: string) {
  const current = settings();
  const profiles = getPersistedSyncProfiles(current);
  const profileById = new Map(profiles.map((profile) => [profile.id, profile]));
  const nextProfiles = scopes.map((scope) => {
    const existing = profileById.get(scope.id);
    return {
      id: scope.id,
      name: scope.name,
      remoteUrl: scope.calendarUrl,
      taskListUrl: existing?.taskListUrl || scope.calendarUrl || "",
      filterPageTypes: existing?.filterPageTypes || "",
      filterTags: existing?.filterTags || "",
      ignoredTags: existing?.ignoredTags || "",
      prefilterPagesOnly: existing?.prefilterPagesOnly === true,
      calendarUrl: scope.calendarUrl,
      propertyKey: scope.propertyKey,
      propertyValue: scope.propertyValue,
      writeToJournal: existing?.writeToJournal === true,
      defaultImportPage: existing?.defaultImportPage || "",
      simpleChecklistMode: existing?.simpleChecklistMode === true,
      updatePeriodically: existing?.updatePeriodically === true,
      updateIntervalMinutes: Math.max(1, Number(existing?.updateIntervalMinutes || 15) || 15),
      enabled: scope.enabled
    };
  });
  await saveSyncProfiles(nextProfiles, activeScopeId);
}

function formatTaskDate(task: LogseqTaskItem) {
  if (!task.date) return "Unscheduled";
  const year = task.date.slice(0, 4);
  const month = task.date.slice(4, 6);
  const day = task.date.slice(6, 8);
  if (!task.time) return `${year}-${month}-${day}`;
  return `${year}-${month}-${day} ${task.time.slice(0, 2)}:${task.time.slice(2, 4)}`;
}

function formatEventDate(event: LogseqCalendarEvent) {
  const year = event.date.slice(0, 4);
  const month = event.date.slice(4, 6);
  const day = event.date.slice(6, 8);
  if (!event.time) return `${year}-${month}-${day}`;
  return `${year}-${month}-${day} ${event.time.slice(0, 2)}:${event.time.slice(2, 4)}`;
}

function stripLogseqRefs(text: string) {
  return String(text ?? "")
    .replace(/\[\[([^\]]+)\]\]/g, "$1")
    .replace(/\(\(([^\)]+)\)\)/g, "$1")
    .replace(/\s+/g, " ")
    .trim();
}

function renderTaskLine(task: LogseqTaskItem) {
  const pageRef = `[[${task.pageName}]]`;
  const blockRef = task.sourceBlockUuid ? `((${task.sourceBlockUuid}))` : "";
  const state = task.taskState ? `[${task.taskState}]` : "";
  const due = formatTaskDate(task);
  const title = stripLogseqRefs(task.title);
  const scopeParts = [
    task.pageType ? `page-type:${task.pageType}` : "",
    task.pageTags.length ? `tags:${task.pageTags.join(",")}` : ""
  ].filter(Boolean);
  const scope = scopeParts.length ? ` {${scopeParts.join(" | ")}}` : "";
  return `${pageRef} :: ${title} ${state} {${due}}${scope} ${blockRef}`.replace(/\s+/g, " ").trim();
}

function renderEventLine(event: LogseqCalendarEvent) {
  const pageRef = `[[${event.pageName}]]`;
  const blockRef = event.sourceBlockUuid ? `((${event.sourceBlockUuid}))` : "";
  const title = stripLogseqRefs(event.title);
  const scope = event.scopePropertyValue ? ` {${event.scopePropertyKey || "calendar"}:${event.scopePropertyValue}}` : "";
  return `${pageRef} :: ${title} {${event.kind} · ${formatEventDate(event)}}${scope} ${blockRef}`.replace(/\s+/g, " ").trim();
}

async function writeTaskOverviewPage(tasks: LogseqTaskItem[], scope = getActiveTaskScope(settings())) {
  const pageName = settings().taskPageName || defaultSettings.taskPageName;
  const timestamp = new Date().toISOString();
  const scopeSummary = describeTaskScope(scope);
  const blocks = [
    "Nextcloud task sync snapshot",
    `Updated: ${timestamp}`,
    `Active scope: ${scope?.name ?? "None"}`,
    `Tasks found: ${tasks.length}`,
    `Scope: ${scopeSummary}`,
    `Task list: ${scope?.taskListUrl || "Not selected"}`,
    ...tasks.map((task) => renderTaskLine(task))
  ];
  await replacePageBlocks(pageName, blocks, "[nextcloud-sync] could not update task overview page");
}

function describeTaskScope(scope?: Pick<TaskScopeConfig, "filterPageTypes" | "filterTags"> | null) {
  const pageTypes = String(scope?.filterPageTypes ?? "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
  const tags = String(scope?.filterTags ?? "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);

  if (!pageTypes.length && !tags.length) return "No automatic match";
  const parts = [
    pageTypes.length ? `page types: ${pageTypes.join(", ")}` : "",
    tags.length ? `tags: ${tags.join(", ")}` : ""
  ].filter(Boolean);
  return parts.join(" OR ");
}

function describeIgnoredTags(scope?: Pick<TaskScopeConfig, "ignoredTags"> | null) {
  const tags = String(scope?.ignoredTags ?? "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
  return tags.length ? tags.join(", ") : "None";
}

function describePagePrefilter(scope?: Pick<TaskScopeConfig, "prefilterPagesOnly"> | null) {
  return scope?.prefilterPagesOnly ? "Matched pages only" : "Whole graph";
}

function describeJournalImport(scope?: Pick<SyncProfileConfig, "writeToJournal"> | null) {
  return scope?.writeToJournal === true ? "yes" : "no";
}

function describeDefaultImportPage(scope?: Pick<SyncProfileConfig, "defaultImportPage" | "name"> | null) {
  const page = String(scope?.defaultImportPage || "").trim();
  return page || nextcloudInboxPageName({ name: scope?.name || "Profile" });
}

function describeSimpleChecklistMode(scope?: Pick<SyncProfileConfig, "simpleChecklistMode"> | null) {
  return scope?.simpleChecklistMode === true ? "yes" : "no";
}

function describePeriodicSync(scope?: Pick<SyncProfileConfig, "updatePeriodically" | "updateIntervalMinutes"> | null) {
  if (scope?.updatePeriodically !== true) return "no";
  const minutes = Math.max(1, Number(scope?.updateIntervalMinutes || 15) || 15);
  return `yes, every ${minutes} min`;
}

function describeCalendarScope(scope?: Pick<CalendarScopeConfig, "propertyKey" | "propertyValue"> | null) {
  const propertyKey = String(scope?.propertyKey || "calendar").trim() || "calendar";
  const propertyValue = String(scope?.propertyValue || "").trim();
  return propertyValue ? `task filters plus ${propertyKey} = ${propertyValue} override` : "same as task filters only";
}

async function writeCalendarOverviewPage(events: LogseqCalendarEvent[], scope = getActiveCalendarScope(settings())) {
  const pageName = settings().calendarPageName || defaultSettings.calendarPageName;
  const timestamp = new Date().toISOString();
  const blocks = [
    "Nextcloud calendar sync snapshot",
    `Updated: ${timestamp}`,
    `Active scope: ${scope?.name ?? "None"}`,
    `Scope: ${describeCalendarScope(scope)}`,
    `Calendar: ${scope?.calendarUrl || "Not selected"}`,
    `Events found: ${events.length}`,
    ...events.map((event) => renderEventLine(event))
  ];
  await replacePageBlocks(pageName, blocks, "[nextcloud-sync] could not update calendar overview page");
}

async function replacePageBlocks(pageName: string, blocks: string[], warningMessage: string) {
  const editor = logseq.Editor as any;

  try {
    await editor.createPage?.(pageName, {}, { redirect: false, createFirstBlock: false });
  } catch {
    // Page may already exist.
  }

  try {
    const existingBlocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
    for (const block of existingBlocks) {
      if (block?.uuid) {
        await editor.removeBlock?.(block.uuid);
      }
    }
    for (const block of blocks) {
      if (!String(block ?? "").trim()) continue;
      await editor.appendBlockInPage?.(pageName, block);
    }
  } catch (error) {
    console.warn(warningMessage, error);
  }
}

function formatDateForJournalPage(dateKey: string, format = "MMM do, yyyy") {
  const year = Number(dateKey.slice(0, 4));
  const month = Number(dateKey.slice(4, 6));
  const day = Number(dateKey.slice(6, 8));
  const monthNamesShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const monthNamesLong = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
  ];
  const ordinal = (() => {
    const remainder10 = day % 10;
    const remainder100 = day % 100;
    if (remainder10 === 1 && remainder100 !== 11) return `${day}st`;
    if (remainder10 === 2 && remainder100 !== 12) return `${day}nd`;
    if (remainder10 === 3 && remainder100 !== 13) return `${day}rd`;
    return `${day}th`;
  })();

  return String(format || "MMM do, yyyy")
    .replace(/yyyy/g, String(year))
    .replace(/MMMM/g, monthNamesLong[month - 1] || String(month))
    .replace(/MMM/g, monthNamesShort[month - 1] || String(month))
    .replace(/do/g, ordinal)
    .replace(/\bdd\b/g, String(day).padStart(2, "0"))
    .replace(/\bd\b/g, String(day));
}

function readGraphConfigTemplateName(config: unknown) {
  if (!config || typeof config !== "object") return "";
  const source = config as Record<string, any>;
  const direct = source["default-templates"] ?? source[":default-templates"] ?? source.defaultTemplates;
  const directJournal =
    direct?.journals ??
    direct?.journal ??
    source["default-templates/journals"] ??
    source[":default-templates/journals"];
  return String(directJournal || "").trim();
}

function readGraphConfigJournalTitleFormat(config: unknown) {
  if (!config || typeof config !== "object") return "";
  const source = config as Record<string, any>;
  const direct =
    source["journal/page-title-format"] ??
    source[":journal/page-title-format"] ??
    source.journal?.["page-title-format"] ??
    source.journalPageTitleFormat;
  return String(direct || "").trim();
}

async function cloneTemplateBlockTree(targetPageName: string, templateBlock: any) {
  const editor = logseq.Editor as any;
  const root = await editor.appendBlockInPage?.(targetPageName, String(templateBlock?.content || ""));
  if (!root?.uuid || !Array.isArray(templateBlock?.children)) return;

  const appendChildren = async (parentUuid: string, children: any[]) => {
    for (const child of children) {
      const inserted = await editor.insertBlock?.(parentUuid, String(child?.content || ""), { sibling: false, isPageBlock: false });
      if (inserted?.uuid && Array.isArray(child?.children) && child.children.length) {
        await appendChildren(inserted.uuid, child.children);
      }
    }
  };

  await appendChildren(root.uuid, templateBlock.children);
}

async function ensureJournalPageForImport(pageName: string) {
  const editor = logseq.Editor as any;
  const existingPage = await editor.getPage?.(pageName);
  await editor.createPage?.(pageName, {}, { redirect: false, createFirstBlock: false, journal: true });
  if (existingPage) return;

  const existingBlocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
  if (existingBlocks.length) return;

  const graphConfig = await logseq.App.getCurrentGraphConfigs?.();
  const templateName = readGraphConfigTemplateName(graphConfig);
  if (!templateName) return;
  const templateBlock = await logseq.App.getTemplate?.(templateName);
  if (!templateBlock) return;
  await cloneTemplateBlockTree(pageName, templateBlock);
}

function nextcloudInboxPageName(profile: Pick<SyncProfileConfig, "name">) {
  return `Nextcloud Inbox - ${profile.name}`;
}

async function openInboxPage(scope = getActiveSyncProfile(settings())) {
  if (!scope) {
    logseq.UI.showMsg?.("Create or select a sync profile first.", "warning", { timeout: 5000 });
    return;
  }
  openPage(nextcloudInboxPageName(scope));
}

function formatPropertyDate(date?: string, time?: string) {
  if (!date) return "";
  const year = date.slice(0, 4);
  const month = date.slice(4, 6);
  const day = date.slice(6, 8);
  if (!time) return `${year}-${month}-${day}`;
  return `${year}-${month}-${day} ${time.slice(0, 2)}:${time.slice(2, 4)}`;
}

function formatLogseqTimestamp(date?: string, time?: string, allDay?: boolean) {
  if (!date) return "";
  const year = date.slice(0, 4);
  const month = date.slice(4, 6);
  const day = date.slice(6, 8);
  if (allDay || !time) return `<${year}-${month}-${day}>`;
  return `<${year}-${month}-${day} ${time.slice(0, 2)}:${time.slice(2, 4)}>`;
}

function importTargetPageName(profile: SyncProfileConfig, date?: string, journalTitleFormat = "MMM do, yyyy") {
  if (date && profile.writeToJournal !== false) {
    return formatDateForJournalPage(date, journalTitleFormat);
  }
  return String(profile.defaultImportPage || "").trim() || nextcloudInboxPageName(profile);
}

function importFallbackPageName(profile: SyncProfileConfig) {
  return String(profile.defaultImportPage || "").trim() || nextcloudInboxPageName(profile);
}

async function resolveGraphRootPath() {
  const explicit = String(settings().graphRootPath || "").trim();
  if (explicit) return explicit;
  try {
    const graph = await logseq.App.getCurrentGraph?.();
    return String(graph?.path || "").trim();
  } catch {
    return "";
  }
}

function importedCalendarUidCacheKey(profileId: string) {
  return String(profileId || "").trim();
}

function readCalendarSyncState(current = settings()) {
  try {
    const parsed = JSON.parse(String(current.calendarSyncStateJson || "{}"));
    return typeof parsed === "object" && parsed ? parsed as Record<string, Record<string, CalendarSyncSnapshot>> : {};
  } catch {
    return {};
  }
}

async function writeCalendarSyncState(stateMap: Record<string, Record<string, CalendarSyncSnapshot>>) {
  const json = JSON.stringify(stateMap);
  state.settingsOverride = {
    ...state.settingsOverride,
    calendarSyncStateJson: json
  };
  await logseq.updateSettings?.({ calendarSyncStateJson: json });
}

function calendarSnapshotFromLocal(event: LogseqCalendarEvent): CalendarSyncSnapshot {
  return {
    title: String(event.title || "").trim(),
    date: String(event.date || "").trim(),
    time: String(event.time || "").trim(),
    endDate: "",
    endTime: "",
    allDay: event.allDay === true,
    remoteResourceUrl: String(event.remoteResourceUrl || "").trim()
  };
}

function calendarSnapshotFromRemote(event: RemoteCalendarImportItem): CalendarSyncSnapshot {
  return {
    title: String(event.title || "").trim(),
    date: String(event.date || "").trim(),
    time: String(event.time || "").trim(),
    endDate: String(event.endDate || "").trim(),
    endTime: String(event.endTime || "").trim(),
    allDay: event.allDay === true,
    remoteResourceUrl: String(event.remoteResourceUrl || "").trim()
  };
}

function sameCalendarSnapshot(a?: CalendarSyncSnapshot | null, b?: CalendarSyncSnapshot | null) {
  if (!a || !b) return false;
  return (
    a.title === b.title &&
    a.date === b.date &&
    a.time === b.time &&
    a.allDay === b.allDay
  );
}

function readImportedCalendarUidCache(profileId: string, current = settings()) {
  const key = importedCalendarUidCacheKey(profileId);
  if (!key) return new Set<string>();
  try {
    const parsed = JSON.parse(String(current.importedCalendarUidCacheJson || "{}"));
    const values = Array.isArray(parsed?.[key]) ? parsed[key] : [];
    return new Set(Array.isArray(values) ? values.map((value) => String(value || "").trim()).filter(Boolean) : []);
  } catch {
    return new Set<string>();
  }
}

async function writeImportedCalendarUidCache(profileId: string, values: Set<string>, current = settings()) {
  const key = importedCalendarUidCacheKey(profileId);
  if (!key) return;
  try {
    const parsed = JSON.parse(String(current.importedCalendarUidCacheJson || "{}"));
    const next = typeof parsed === "object" && parsed ? parsed : {};
    next[key] = Array.from(values).sort();
    await logseq.updateSettings?.({ importedCalendarUidCacheJson: JSON.stringify(next) });
  } catch {
    // Ignore cache persistence failures; sync can still continue.
  }
}

async function journalMarkdownContainsRemoteUid(dateKey: string | undefined, remoteUid: string) {
  const normalizedDate = String(dateKey || "").trim().replace(/-/g, "");
  const normalizedUid = String(remoteUid || "").trim();
  if (!/^\d{8}$/.test(normalizedDate) || !normalizedUid) return false;

  try {
    const graphPath = await resolveGraphRootPath();
    const req = (globalThis as any).require || (globalThis as any).window?.require || (globalThis as any).top?.require;
    if (!graphPath || typeof req !== "function") return false;

    const fs = req("fs") as any;
    const path = req("path") as any;
    const year = normalizedDate.slice(0, 4);
    const month = normalizedDate.slice(4, 6);
    const day = normalizedDate.slice(6, 8);
    const filePath = path.join(graphPath, "journals", `${year}_${month}_${day}.md`);
    if (!fs.existsSync(filePath)) return false;

    const text = fs.readFileSync(filePath, "utf8");
    return text.includes(`nextcloud-remote-uid:: ${normalizedUid}`);
  } catch (error) {
    console.warn("[nextcloud-sync] could not inspect journal markdown for remote UID", error);
    return false;
  }
}

async function ensureChildBlock(parentUuid: string, content: string) {
  const editor = logseq.Editor as any;
  const children = ((await editor.getBlock?.(parentUuid, { includeChildren: true }))?.children ?? []) as any[];
  const existing = children.find((child) => String(child?.content || "").trim() === content);
  if (existing?.uuid) return existing;
  return editor.insertBlock?.(parentUuid, content, { sibling: false, isPageBlock: false });
}

function findBlockByContentRecursive(blocks: any[], content: string): any | null {
  const target = String(content || "").trim();
  for (const block of blocks) {
    if (String(block?.content || "").trim() === target) {
      return block;
    }
    if (Array.isArray(block?.children)) {
      const nested = findBlockByContentRecursive(block.children, target);
      if (nested) return nested;
    }
  }
  return null;
}

async function ensureImportSection(pageName: string, profile: SyncProfileConfig, section: "Tasks" | "Events") {
  const editor = logseq.Editor as any;
  let page = await editor.getPage?.(pageName);
  if (!page) {
    await editor.createPage?.(pageName, {}, { redirect: false, createFirstBlock: false });
    page = await editor.getPage?.(pageName);
  }
  const pageBlocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
  let root = findBlockByContentRecursive(pageBlocks, "Nextcloud");
  if (!root) {
    root = await editor.appendBlockInPage?.(pageName, "Nextcloud");
  }
  if (!root?.uuid) throw new Error(`Could not create Nextcloud section on ${pageName}.`);
  const profileBlock = await ensureChildBlock(root.uuid, profile.name);
  if (!profileBlock?.uuid) throw new Error(`Could not create ${profile.name} section on ${pageName}.`);
  const sectionBlock = await ensureChildBlock(profileBlock.uuid, section);
  if (!sectionBlock?.uuid) throw new Error(`Could not create ${section} section on ${pageName}.`);
  return sectionBlock.uuid;
}

async function findImportSectionBlock(pageName: string, profile: SyncProfileConfig, section: "Tasks" | "Events") {
  const editor = logseq.Editor as any;
  const pageBlocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
  const root = findBlockByContentRecursive(pageBlocks, "Nextcloud");
  if (!root?.uuid || !Array.isArray(root.children)) return null;
  const profileBlock = root.children.find((child: any) => String(child?.content || "").trim() === profile.name);
  if (!profileBlock?.uuid || !Array.isArray(profileBlock.children)) return null;
  return profileBlock.children.find((child: any) => String(child?.content || "").trim() === section) ?? null;
}

async function buildImportedUidIndex() {
  const editor = logseq.Editor as any;
  const pages = (await editor.getAllPages?.()) ?? [];
  const byUid = new Map<string, any>();

  const visit = async (block: any) => {
    const properties = await readImportedBlockPropertiesAsync(block);
    const uid = getImportedRemoteUid(properties);
    if (uid && block?.uuid && !byUid.has(uid)) {
      byUid.set(uid, block);
    }
    if (Array.isArray(block?.children)) {
      for (const child of block.children) await visit(child);
    }
  };

  for (const page of pages) {
    const pageName = String(page?.name ?? page?.originalName ?? "").trim();
    if (!pageName) continue;
    const blocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
    for (const block of blocks) await visit(block);
  }

  return byUid;
}

async function findImportedBlockByUidOnPage(pageName: string, uid: string, profileId?: string) {
  const editor = logseq.Editor as any;
  const targetUid = String(uid || "").trim();
  const targetProfileId = String(profileId || "").trim();
  if (!pageName || !targetUid) return null;

  const visit = async (block: any): Promise<any | null> => {
    const properties = await readImportedBlockPropertiesAsync(block);
    const blockUid = getImportedRemoteUid(properties);
    const blockProfileId = getImportedProfileId(properties);
    if (blockUid === targetUid && (!targetProfileId || blockProfileId === targetProfileId)) {
      return block;
    }
    if (Array.isArray(block?.children)) {
      for (const child of block.children) {
        const nested = await visit(child);
        if (nested) return nested;
      }
    }
    return null;
  };

  const blocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
  for (const block of blocks) {
    const found = await visit(block);
    if (found) return found;
  }

  return null;
}

async function findImportedEventByIdentityOnPage(
  pageName: string,
  profileId: string,
  title: string,
  date?: string,
  time?: string
) {
  const editor = logseq.Editor as any;
  if (!pageName || !profileId || !title) return null;
  const blocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
  const normalizedTitle = normalizeText(title);
  const normalizedDate = String(date || "").trim().replace(/-/g, "");
  const normalizedTime = String(time || "").trim();

  const visit = async (block: any): Promise<any | null> => {
    const properties = await readImportedBlockPropertiesAsync(block);
    const blockProfileId = getImportedProfileId(properties);
    if (blockProfileId === profileId) {
      const blockTitle = normalizeText(String(block?.content || ""));
      const blockDate = String(properties?.start || properties?.date || "").trim();
      const parsed = parseCompactDateTime(blockDate);
      if (
        blockTitle === normalizedTitle &&
        (!normalizedDate || parsed?.date === normalizedDate) &&
        (!normalizedTime || parsed?.time === normalizedTime)
      ) {
        return block;
      }
    }
    if (Array.isArray(block?.children)) {
      for (const child of block.children) {
        const nested = await visit(child);
        if (nested) return nested;
      }
    }
    return null;
  };

  for (const block of blocks) {
    const found = await visit(block);
    if (found) return found;
  }

  return null;
}

async function findImportedEventInSection(
  pageName: string,
  profile: SyncProfileConfig,
  item: Pick<RemoteCalendarImportItem, "title" | "date" | "time" | "endDate" | "endTime">
) {
  const sectionBlock = await findImportSectionBlock(pageName, profile, "Events");
  const children = Array.isArray(sectionBlock?.children) ? sectionBlock.children : [];
  const normalizedTitle = normalizeText(item.title);
  const normalizedDate = String(item.date || "").trim().replace(/-/g, "");
  const normalizedTime = String(item.time || "").trim();
  const normalizedEndDate = String(item.endDate || "").trim().replace(/-/g, "");
  const normalizedEndTime = String(item.endTime || "").trim();

  for (const child of children) {
    if (normalizeText(String(child?.content || "")) !== normalizedTitle) continue;
    const properties = await readImportedBlockPropertiesAsync(child);
    const startValue = String(properties?.start || properties?.date || "").trim();
    const endValue = String(properties?.end || "").trim();
    const parsedStart = parseCompactDateTime(startValue);
    const parsedEnd = parseCompactDateTime(endValue);
    const sameStart =
      (!normalizedDate || parsedStart?.date === normalizedDate) &&
      (!normalizedTime || parsedStart?.time === normalizedTime);
    const sameEnd =
      (!normalizedEndDate || parsedEnd?.date === normalizedEndDate) &&
      (!normalizedEndTime || parsedEnd?.time === normalizedEndTime);
    if (sameStart && sameEnd) {
      return child;
    }
  }

  return null;
}

async function findSimpleChecklistTaskByTitleOnPage(pageName: string, title: string) {
  const editor = logseq.Editor as any;
  if (!pageName || !title) return null;
  const blocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
  const normalizedTitle = normalizeChecklistTaskTitle(title);

  const visit = async (block: any, ancestry: string[] = []): Promise<any | null> => {
    const content = String(block?.content || "").trim();
    const nextAncestry = content ? [...ancestry, content] : ancestry;
    const markerMatch = content.match(/^\s*(TODO|DOING|DONE|LATER|WAITING|NOW|CANCELLED)\b/i);
    if (markerMatch) {
      const properties = await readImportedBlockPropertiesAsync(block);
      const blockRemoteUid = getImportedRemoteUid(properties);
      const inImportedSection =
        nextAncestry.some((item) => String(item).trim() === "Nextcloud") &&
        nextAncestry.some((item) => String(item).trim() === "Tasks");
      if (!inImportedSection && !blockRemoteUid && normalizeChecklistTaskTitle(content) === normalizedTitle) {
        return block;
      }
    }
    if (Array.isArray(block?.children)) {
      for (const child of block.children) {
        const nested = await visit(child, nextAncestry);
        if (nested) return nested;
      }
    }
    return null;
  };

  for (const block of blocks) {
    const found = await visit(block);
    if (found) return found;
  }

  return null;
}

async function findSimpleChecklistInsertParent(pageName: string) {
  const editor = logseq.Editor as any;
  const pageBlocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
  const headings = pageBlocks.filter((block: any) => /^#{1,6}\s+/.test(String(block?.content || "").trim()));
  return headings.length ? headings[headings.length - 1] : null;
}

async function collectImportedBlocks() {
  const editor = logseq.Editor as any;
  const pages = (await editor.getAllPages?.()) ?? [];
  const imported: Array<{ pageName: string; block: any; profileId: string; remoteUid: string }> = [];

  const visit = async (pageName: string, block: any) => {
    const properties = await readImportedBlockPropertiesAsync(block);
    const remoteUid = getImportedRemoteUid(properties);
    const profileId = getImportedProfileId(properties);
    if (remoteUid && profileId && block?.uuid) {
      imported.push({ pageName, block, profileId, remoteUid });
    }
    if (Array.isArray(block?.children)) {
      for (const child of block.children) await visit(pageName, child);
    }
  };

  for (const page of pages) {
    const pageName = String(page?.name ?? page?.originalName ?? "").trim();
    if (!pageName) continue;
    const blocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
    for (const block of blocks) await visit(pageName, block);
  }

  return imported;
}

function readImportedBlockProperties(block: any) {
  return {
    ...(block?.properties ?? {}),
    ...(block?.meta?.properties ?? {}),
    ...readPropertyChildren(block)
  } as Record<string, any>;
}

function normalizeText(input: string) {
  return String(input ?? "").trim().toLowerCase();
}

function normalizeChecklistTaskTitle(input: string) {
  return String(input ?? "")
    .replace(/^\s*(TODO|DOING|DONE|LATER|WAITING|NOW|CANCELLED)\s+/i, "")
    .replace(/\[\#[A-Z]\]/g, "")
    .replace(/SCHEDULED:\s*<[^>]+>/gi, "")
    .replace(/DEADLINE:\s*<[^>]+>/gi, "")
    .replace(/(^|\s)#[A-Za-z0-9/_-]+/g, "$1")
    .trim()
    .split("\n")[0]
    ?.trim()
    .toLowerCase() || "";
}

async function readImportedBlockPropertiesAsync(block: any) {
  const editor = logseq.Editor as any;
  let liveProperties: Record<string, any> = {};
  if (block?.uuid && typeof editor.getBlockProperties === "function") {
    try {
      liveProperties = (await editor.getBlockProperties(block.uuid)) ?? {};
    } catch {
      liveProperties = {};
    }
  }

  return {
    ...liveProperties,
    ...readImportedBlockProperties(block)
  } as Record<string, any>;
}

function normalizeImportedPropertyKey(input: string) {
  return String(input ?? "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function getImportedProperty(properties: Record<string, any>, ...keys: string[]) {
  const normalizedKeys = keys.map((key) => normalizeImportedPropertyKey(key));
  for (const [rawKey, rawValue] of Object.entries(properties || {})) {
    if (!normalizedKeys.includes(normalizeImportedPropertyKey(rawKey))) continue;
    return Array.isArray(rawValue) ? rawValue[0] ?? "" : rawValue;
  }
  return "";
}

function getImportedRemoteUid(properties: Record<string, any>) {
  return String(
    getImportedProperty(properties, "nextcloud-remote-uid", "nextcloud_remote_uid", "nextcloudremoteuid")
  ).trim();
}

function getImportedProfileId(properties: Record<string, any>) {
  return String(
    getImportedProperty(properties, "nextcloud-profile", "nextcloud_profile", "nextcloudprofile")
  ).trim();
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

async function normalizeImportedBlockProperties() {
  const editor = logseq.Editor as any;
  const pages = (await editor.getAllPages?.()) ?? [];

  const visit = async (block: any, ancestry: string[]) => {
    const content = String(block?.content || "").trim();
    const nextAncestry = content ? [...ancestry, content] : ancestry;
    const inImportedSection =
      nextAncestry.some((item) => item === "Nextcloud") &&
      nextAncestry.some((item) => item === "Events" || item === "Tasks");

    if (inImportedSection && block?.uuid) {
      const properties = readPropertyChildren(block);
      for (const [key, value] of Object.entries(properties)) {
        const normalizedKey = key === "nextcloud-resource-url" ? "nextcloud_resource_url" : key;
        try {
          await editor.upsertBlockProperty?.(block.uuid, normalizedKey, value);
          if (normalizedKey !== key) {
            await editor.removeBlockProperty?.(block.uuid, key);
          }
        } catch (error) {
          console.warn(`[nextcloud-sync] could not normalize imported property ${key}`, error);
        }
      }
    }

    if (Array.isArray(block?.children)) {
      for (const child of block.children) {
        await visit(child, nextAncestry);
      }
    }
  };

  for (const page of pages) {
    const pageName = String(page?.name ?? page?.originalName ?? "").trim();
    if (!pageName) continue;
    const blocks = (await editor.getPageBlocksTree?.(pageName)) ?? [];
    for (const block of blocks) {
      await visit(block, []);
    }
  }
}

function parseCompactDateTime(value: string) {
  const raw = String(value || "").trim().replace(/[<>]/g, "");
  const dateTime = raw.match(/^(\d{4})[-_]?(\d{2})[-_]?(\d{2})(?:\s+(\d{2}):(\d{2}))?$/);
  if (!dateTime) return null;
  const date = `${dateTime[1]}${dateTime[2]}${dateTime[3]}`;
  const time = dateTime[4] && dateTime[5] ? `${dateTime[4]}${dateTime[5]}00` : undefined;
  return { date, time, allDay: !time };
}

function readImportedTaskDate(block: any) {
  const content = String(block?.content || "");
  const match = content.match(/(?:DEADLINE|SCHEDULED):\s*<([^>]+)>/i);
  return match?.[1] ? parseCompactDateTime(match[1]) : null;
}

function readImportedEventDate(block: any) {
  const properties = readImportedBlockProperties(block);
  const start = properties.start ?? properties.date ?? properties.deadline ?? properties.scheduled;
  return parseCompactDateTime(String(Array.isArray(start) ? start[0] ?? "" : start ?? ""));
}

function classifyImportedBlock(block: any) {
  const content = String(block?.content || "");
  return /^\s*(TODO|DOING|DONE|LATER|WAITING|NOW|CANCELLED)\b/i.test(content) ? "Tasks" : "Events";
}

async function dedupeImportedItems(scope = getActiveSyncProfile(settings())) {
  if (!scope) {
    logseq.UI.showMsg?.("Create or select a sync profile first.", "warning", { timeout: 5000 });
    return;
  }

  try {
    logseq.UI.showMsg?.(`Running dedupe for ${scope.name}...`, "success", { timeout: 2500 });
    const editor = logseq.Editor as any;
    const imported = await collectImportedBlocks();
    const seen = new Set<string>();
    let removed = 0;

    for (const item of imported) {
      if (item.profileId !== scope.id) continue;
      const kind = classifyImportedBlock(item.block);
      const key = `${scope.id}::${kind}::${item.remoteUid}`;
      if (!seen.has(key)) {
        seen.add(key);
        continue;
      }
      if (item.block?.uuid) {
        await editor.removeBlock?.(item.block.uuid);
        removed += 1;
      }
    }

    const remoteRemoved = await cleanupRemoteCalendarDuplicates(scope);
    logseq.UI.showMsg?.(
      `${scope.name}: removed ${removed} duplicate Logseq blocks and ${remoteRemoved} stale remote calendar events.`,
      removed || remoteRemoved ? "success" : "warning",
      { timeout: 8000 }
    );
  } catch (error) {
    console.error("[nextcloud-sync] dedupe failed", error);
    logseq.UI.showMsg?.(
      error instanceof Error ? `Dedupe failed: ${error.message}` : "Dedupe failed.",
      "error",
      { timeout: 8000 }
    );
  }
}

async function dedupeImportedItemsForAllProfiles() {
  const profiles = getSyncProfiles(settings()).filter((profile) => profile.enabled);
  if (!profiles.length) {
    logseq.UI.showMsg?.("No enabled sync profiles found.", "warning", { timeout: 5000 });
    return;
  }

  for (const profile of profiles) {
    await dedupeImportedItems(profile);
  }
}

type RemoteCleanupEvent = {
  uid: string;
  summary: string;
  dtstart: string;
  dtend: string;
};

async function cleanupRemoteCalendarDuplicates(profile: SyncProfileConfig) {
  const calendarUrl = String(profile.remoteUrl || profile.calendarUrl || "").trim();
  const current = settings();
  if (!calendarUrl || !current.caldavUsername || !current.caldavPassword) return 0;

  const authHeader = `Basic ${btoa(`${current.caldavUsername}:${current.caldavPassword}`)}`;
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
    throw new Error(`Remote calendar cleanup failed: HTTP ${response.status} ${response.statusText}`);
  }

  const events = Array.from(responseText.matchAll(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g))
    .map((match) => parseRemoteCleanupEvent(match[0]))
    .filter((item): item is RemoteCleanupEvent => Boolean(item));

  const canonicalKeys = new Set(
    events
      .filter((event) => !event.uid.endsWith("@logseq.local"))
      .map((event) => remoteCleanupKey(event))
  );

  const stale = events.filter((event) => event.uid.endsWith("@logseq.local") && canonicalKeys.has(remoteCleanupKey(event)));
  let removed = 0;

  for (const event of stale) {
    const targetUrl = `${calendarUrl.replace(/\/+$/, "")}/${encodeURIComponent(event.uid)}.ics`;
    const deleteResponse = await fetch(targetUrl, {
      method: "DELETE",
      headers: {
        Authorization: authHeader
      },
      credentials: "include"
    });
    if (deleteResponse.ok) {
      removed += 1;
    }
  }

  return removed;
}

function parseRemoteCleanupEvent(block: string): RemoteCleanupEvent | null {
  const read = (key: string) => ((block.match(new RegExp(`^${key}(?:;[^:]*)?:(.+)$`, "im")) || [])[1] || "").trim();
  const uid = read("UID");
  const summary = read("SUMMARY");
  const dtstart = read("DTSTART");
  const dtend = read("DTEND");
  if (!uid || !dtstart) return null;
  return { uid, summary, dtstart, dtend };
}

function remoteCleanupKey(event: RemoteCleanupEvent) {
  return [event.summary.trim(), event.dtstart.trim(), event.dtend.trim()].join("::");
}

async function cleanupImportedItems(scope = getActiveSyncProfile(settings())) {
  if (!scope) {
    logseq.UI.showMsg?.("Create or select a sync profile first.", "warning", { timeout: 5000 });
    return;
  }

  const editor = logseq.Editor as any;
  const graphConfig = await logseq.App.getCurrentGraphConfigs?.();
  const journalTitleFormat = readGraphConfigJournalTitleFormat(graphConfig) || "MMM do, yyyy";
  const profiles = new Map(getSyncProfiles(settings()).map((profile) => [profile.id, profile]));
  const imported = await collectImportedBlocks();
  let moved = 0;

  for (const item of imported) {
    if (item.profileId !== scope.id) continue;
    const profile = profiles.get(item.profileId);
    if (!profile) continue;

    const section = classifyImportedBlock(item.block);
    const dateInfo = section === "Tasks" ? readImportedTaskDate(item.block) : readImportedEventDate(item.block);
    const targetPageName = importTargetPageName(profile, dateInfo?.date, journalTitleFormat);

    if (dateInfo?.date && profile.writeToJournal !== false) {
      await ensureJournalPageForImport(targetPageName);
    }
    const targetParentUuid = await ensureImportSection(targetPageName, profile, section);
    if (!targetParentUuid || !item.block?.uuid) continue;
    if (item.pageName === targetPageName && String(item.block?.parent?.uuid || item.block?.parent?.id || "") === String(targetParentUuid)) {
      continue;
    }

    await editor.moveBlock?.(item.block.uuid, targetParentUuid, { children: true });
    moved += 1;
  }

  logseq.UI.showMsg?.(
    `${scope.name}: cleaned up ${moved} imported Nextcloud blocks.`,
    moved ? "success" : "warning",
    { timeout: 7000 }
  );
}

async function cleanupImportedItemsForAllProfiles() {
  const profiles = getSyncProfiles(settings()).filter((profile) => profile.enabled);
  if (!profiles.length) {
    logseq.UI.showMsg?.("No enabled sync profiles found.", "warning", { timeout: 5000 });
    return;
  }

  for (const profile of profiles) {
    await cleanupImportedItems(profile);
  }
}

function calendarSelectorTemplate() {
  const current = settings();
  const discovery = state.calendarDiscovery;
  const scopes = getCalendarScopes(current);
  const activeScope = getActiveCalendarScope(current);
  const items = discovery?.calendars ?? [];
  const status = discovery
    ? `<div class="nextcloud-selector__status">${escapeHtml(discovery.message)}</div>`
    : `<div class="nextcloud-selector__status">Create scopes, then run discovery to load calendars from your Nextcloud account.</div>`;

  const scopeCards = scopes
    .map(
      (scope) => `
        <div class="nextcloud-selector__card">
          <div class="nextcloud-selector__title">${escapeHtml(scope.name)}${scope.id === activeScope?.id ? " (active)" : ""}</div>
          <div class="nextcloud-selector__meta">Rule: ${escapeHtml(describeCalendarScope(scope))}</div>
          <div class="nextcloud-selector__meta">Calendar: ${escapeHtml(scope.calendarUrl || "Not selected")}</div>
          <div class="nextcloud-selector__actions">
            <button data-on-click="activateCalendarScope" data-scope-id="${escapeHtml(scope.id)}">Activate</button>
            <button data-on-click="refreshCalendarScopePreview" data-scope-id="${escapeHtml(scope.id)}">Preview</button>
            <button data-on-click="syncCalendarScope" data-scope-id="${escapeHtml(scope.id)}">Sync</button>
            <button data-on-click="editCalendarScope" data-scope-id="${escapeHtml(scope.id)}">Edit</button>
          </div>
        </div>
      `
    )
    .join("");

  const cards = items.length
    ? items
        .map(
          (calendar, index) => `
            <div class="nextcloud-selector__card">
              <div class="nextcloud-selector__title">${escapeHtml(calendar.displayName || `Calendar ${index + 1}`)}</div>
              <div class="nextcloud-selector__meta">${escapeHtml(calendar.url)}</div>
              <div class="nextcloud-selector__meta">${escapeHtml(calendar.componentSet.join(", ") || "VEVENT")}</div>
              <div class="nextcloud-selector__actions">
                <button data-on-click="selectCalendarOption" data-url="${escapeHtml(calendar.url)}">Assign To Active Scope</button>
              </div>
            </div>
          `
        )
        .join("")
    : `<div class="nextcloud-selector__empty">No calendars discovered yet.</div>`;

  return `
    <div class="nextcloud-selector">
      <div class="nextcloud-selector__header">
        <div>
          <div class="nextcloud-selector__headline">Nextcloud Calendar Picker</div>
          <div class="nextcloud-selector__hint">Calendars follow the profile's task filters by default. The calendar property fields are optional extra include rules.</div>
        </div>
        <div class="nextcloud-selector__actions">
          <button data-on-click="createCalendarScope">Create Scope</button>
          <button data-on-click="refreshCalendarDiscovery">Discover Calendars</button>
          <button data-on-click="syncAllCalendarScopes">Sync All Enabled Scopes</button>
        </div>
      </div>
      ${status}
      <div class="nextcloud-selector__section-title">Calendar Scopes</div>
      <div class="nextcloud-selector__list">${scopeCards}</div>
      <div class="nextcloud-selector__section-title">Discovered Calendars</div>
      <div class="nextcloud-selector__list">${cards}</div>
    </div>
  `;
}

function syncCalendarSelectorUI() {
  if (!uiState.calendarSelectorSlot || typeof logseq.provideUI !== "function") return;
  let template = "";
  try {
    template = calendarSelectorTemplate();
  } catch (error) {
    console.error("[nextcloud-sync] could not render calendar selector", error);
    const message = error instanceof Error ? error.message : String(error);
    template = `<div class="nextcloud-selector"><div class="nextcloud-selector__status">Calendar selector failed to render: ${escapeHtml(message)}</div></div>`;
  }
  try {
    logseq.provideUI({
      key: slotUiKey(uiKeys.calendarSelector, uiState.calendarSelectorSlot),
      slot: uiState.calendarSelectorSlot,
      reset: true,
      template
    });
  } catch (error) {
    console.warn("[nextcloud-sync] dropped stale calendar selector slot", error);
    uiState.calendarSelectorSlot = "";
  }
}

function taskScopeManagerTemplate() {
  const current = settings();
  const profiles = getSyncProfiles(current);
  const activeProfile = getActiveSyncProfile(current);
  const discoveredTaskLists = state.taskListDiscovery?.taskLists ?? [];

  const profileCards = profiles
    .map(
      (profile) => `
        <div class="nextcloud-selector__card nextcloud-profile-card" data-scope-id="${escapeHtml(profile.id)}">
          <div class="nextcloud-selector__title">
            ${escapeHtml(profile.name)}${profile.id === activeProfile?.id ? " (active)" : ""}
          </div>
          <div class="nextcloud-selector__meta">Tasks: ${escapeHtml(describeTaskScope(profile))}</div>
          <div class="nextcloud-selector__meta">Ignore tags: ${escapeHtml(describeIgnoredTags(profile))}</div>
          <div class="nextcloud-selector__meta">Page scan: ${escapeHtml(describePagePrefilter(profile))}</div>
          <div class="nextcloud-selector__meta">Remote URL: ${escapeHtml(profile.remoteUrl || profile.taskListUrl || profile.calendarUrl || "Not selected")}</div>
          <div class="nextcloud-selector__meta">Calendar: ${escapeHtml(describeCalendarScope(profile))}</div>
          <div class="nextcloud-selector__meta">Write to journal: ${escapeHtml(describeJournalImport(profile))}</div>
          <div class="nextcloud-selector__meta">Default import page: ${escapeHtml(describeDefaultImportPage(profile))}</div>
          <div class="nextcloud-selector__meta">Simple checklist mode: ${escapeHtml(describeSimpleChecklistMode(profile))}</div>
          <div class="nextcloud-selector__meta">Periodic sync: ${escapeHtml(describePeriodicSync(profile))}</div>
          <div class="nextcloud-selector__meta">Enabled: ${profile.enabled ? "yes" : "no"}</div>
          <div class="nextcloud-selector__actions">
            <button data-on-click="activateTaskScope" data-scope-id="${escapeHtml(profile.id)}">Activate</button>
            <button data-on-click="refreshTaskScopePreview" data-scope-id="${escapeHtml(profile.id)}">Preview Tasks</button>
            <button data-on-click="refreshCalendarScopePreview" data-scope-id="${escapeHtml(profile.id)}">Preview Calendar</button>
            <button data-on-click="syncTaskScope" data-scope-id="${escapeHtml(profile.id)}">Sync Both</button>
            <button data-on-click="editTaskScope" data-scope-id="${escapeHtml(profile.id)}">Edit</button>
          </div>
        </div>
      `
    )
    .join("");

  const discoveredCards = discoveredTaskLists.length
    ? discoveredTaskLists
        .map(
          (item) => `
            <div class="nextcloud-selector__card">
              <div class="nextcloud-selector__title">${escapeHtml(item.displayName)}</div>
              <div class="nextcloud-selector__meta">${escapeHtml(item.url)}</div>
              <div class="nextcloud-selector__actions">
                <button data-on-click="assignTaskListToActiveScope" data-url="${escapeHtml(item.url)}">Assign To Active Scope</button>
              </div>
            </div>
          `
        )
        .join("")
    : `<div class="nextcloud-selector__empty">No remote collections discovered yet.</div>`;

  const status = state.taskListDiscovery
    ? `<div class="nextcloud-selector__status">${escapeHtml(state.taskListDiscovery.message)}</div>`
    : `<div class="nextcloud-selector__status">Create a profile, then discover or create a remote collection for it.</div>`;

  return `
    <div class="nextcloud-selector">
      <div class="nextcloud-selector__header">
        <div>
          <div class="nextcloud-selector__headline">Nextcloud Sync Profiles</div>
          <div class="nextcloud-selector__hint">Each profile owns both task and calendar sync. Empty filters match nothing automatically.</div>
        </div>
        <div class="nextcloud-selector__actions">
          <button data-on-click="createTaskScope">Create Profile</button>
          <button data-on-click="discoverTaskLists">Discover Collections</button>
          <button data-on-click="refreshCalendarDiscovery">Discover Calendars</button>
          <button data-on-click="createRemoteTaskListForActiveScope">Create Remote Collection</button>
          <button data-on-click="syncAllTaskScopes">Sync All Enabled Profiles</button>
        </div>
      </div>
      ${status}
      <div class="nextcloud-selector__section-title">Profiles</div>
      <div class="nextcloud-selector__list">${profileCards}</div>
      <div class="nextcloud-selector__section-title">Discovered Collections</div>
      <div class="nextcloud-selector__list">${discoveredCards}</div>
    </div>
  `;
}

function syncHubTemplate() {
  const current = settings();
  const profiles = getSyncProfiles(current);
  const activeProfile = getActiveSyncProfile(current);
  const taskPageName = current.taskPageName || defaultSettings.taskPageName;
  const calendarPageName = current.calendarPageName || defaultSettings.calendarPageName;
  const navCards = [
    { label: "Profiles", hint: "Edit filters, remote collection, and active profile.", action: "openSyncProfilesPage" },
    { label: "Tasks", hint: "Open the latest task preview page.", action: "openTaskOverviewPage" },
    { label: "Inbox", hint: "Open the active profile's imported remote items inbox.", action: "openInboxOverviewPage" },
    { label: "Calendar", hint: "Open the latest calendar preview page.", action: "openCalendarOverviewPage" },
    { label: "Collections", hint: "Discover and assign remote collections.", action: "openCalendarPickerPage" }
  ];

  const profileCards = profiles.length
    ? profiles
        .map(
          (profile) => `
            <div class="nextcloud-selector__card nextcloud-selector__profile-card">
              <div class="nextcloud-selector__title">${escapeHtml(profile.name)}${profile.id === activeProfile?.id ? " (active)" : ""}</div>
              <div class="nextcloud-selector__meta">Tasks: ${escapeHtml(describeTaskScope(profile))}</div>
              <div class="nextcloud-selector__meta">Ignore tags: ${escapeHtml(describeIgnoredTags(profile))}</div>
              <div class="nextcloud-selector__meta">Page scan: ${escapeHtml(describePagePrefilter(profile))}</div>
              <div class="nextcloud-selector__meta">Remote URL: ${escapeHtml(profile.remoteUrl || "Not selected")}</div>
              <div class="nextcloud-selector__meta">Write to journal: ${escapeHtml(describeJournalImport(profile))}</div>
              <div class="nextcloud-selector__meta">Inbox: ${escapeHtml(nextcloudInboxPageName(profile))}</div>
              <div class="nextcloud-selector__meta">Default import page: ${escapeHtml(describeDefaultImportPage(profile))}</div>
              <div class="nextcloud-selector__meta">Simple checklist mode: ${escapeHtml(describeSimpleChecklistMode(profile))}</div>
              <div class="nextcloud-selector__meta">Periodic sync: ${escapeHtml(describePeriodicSync(profile))}</div>
              <div class="nextcloud-selector__meta">Enabled: ${profile.enabled ? "yes" : "no"}</div>
              <div class="nextcloud-selector__actions nextcloud-selector__actions--profile">
                <button data-on-click="activateTaskScope" data-scope-id="${escapeHtml(profile.id)}">Activate</button>
                <button data-on-click="syncTaskScope" data-scope-id="${escapeHtml(profile.id)}">Sync ${escapeHtml(profile.name)}</button>
                <button data-on-click="importRemoteProfileItems" data-scope-id="${escapeHtml(profile.id)}">Import</button>
                <button data-on-click="openProfileInboxPage" data-scope-id="${escapeHtml(profile.id)}">Open Inbox</button>
                <button data-on-click="cleanupImportedProfileItems" data-scope-id="${escapeHtml(profile.id)}">Cleanup</button>
                <button data-on-click="refreshTaskScopePreview" data-scope-id="${escapeHtml(profile.id)}">Preview Tasks</button>
                <button data-on-click="refreshCalendarScopePreview" data-scope-id="${escapeHtml(profile.id)}">Preview Calendar</button>
                <button data-on-click="editTaskScope" data-scope-id="${escapeHtml(profile.id)}">Edit</button>
              </div>
            </div>
          `
        )
        .join("")
    : `<div class="nextcloud-selector__empty">No sync profiles yet. Create one from the profiles page.</div>`;

  const nav = navCards
    .map(
      (card) => `
        <button class="nextcloud-selector__nav-card" data-on-click="${card.action}">
          <span class="nextcloud-selector__nav-title">${escapeHtml(card.label)}</span>
          <span class="nextcloud-selector__nav-hint">${escapeHtml(card.hint)}</span>
        </button>
      `
    )
    .join("");

  return `
    <div class="nextcloud-selector">
      <div class="nextcloud-selector__header">
        <div>
          <div class="nextcloud-selector__headline">Nextcloud Sync</div>
          <div class="nextcloud-selector__hint">Use this page as the daily control center for profile sync. Preview reads Logseq only. Import pulls remote-only Nextcloud items into the right journal page or profile inbox.</div>
        </div>
        <div class="nextcloud-selector__actions">
          <button class="nextcloud-selector__primary-action" data-on-click="syncAllTaskScopes">Sync All Profiles</button>
          <button data-on-click="importAllRemoteItems">Import All Remote Items</button>
          <button data-on-click="cleanupAllImportedItems">Cleanup Imported Items</button>
        </div>
      </div>
      <div class="nextcloud-selector__status">
        Active profile: ${escapeHtml(activeProfile?.name || "None")} | Tasks page: ${escapeHtml(taskPageName)} | Calendar page: ${escapeHtml(calendarPageName)}
      </div>
      <div class="nextcloud-selector__section-title">Open Pages</div>
      <div class="nextcloud-selector__nav-grid">${nav}</div>
      <div class="nextcloud-selector__section-title">Profiles</div>
      <div class="nextcloud-selector__list">${profileCards}</div>
    </div>
  `;
}

function syncHubUI() {
  if (!uiState.syncHubSlot || typeof logseq.provideUI !== "function") return;
  let template = "";
  try {
    template = syncHubTemplate();
  } catch (error) {
    console.error("[nextcloud-sync] could not render sync hub", error);
    const message = error instanceof Error ? error.message : String(error);
    template = `<div class="nextcloud-selector"><div class="nextcloud-selector__status">Nextcloud Sync failed to render: ${escapeHtml(message)}</div></div>`;
  }
  try {
    logseq.provideUI({
      key: slotUiKey(uiKeys.syncHub, uiState.syncHubSlot),
      slot: uiState.syncHubSlot,
      reset: true,
      template
    });
  } catch (error) {
    console.warn("[nextcloud-sync] dropped stale sync hub slot", error);
    uiState.syncHubSlot = "";
  }
}

function profileEditorTemplate() {
  return `
    <div class="nextcloud-selector">
      <div class="nextcloud-selector__header">
        <div>
          <div class="nextcloud-selector__headline">Profile Editor</div>
          <div class="nextcloud-selector__hint">Edit the blocks on this page, then use Save to persist them.</div>
        </div>
        <div class="nextcloud-selector__actions">
          <button class="nextcloud-selector__primary-action" data-on-click="saveProfileEditorFromUi">Save Profile</button>
          <button data-on-click="openSyncProfilesPage">Back To Profiles</button>
        </div>
      </div>
    </div>
  `;
}

function syncProfileEditorUI() {
  if (!uiState.profileEditorSlot || typeof logseq.provideUI !== "function") return;
  try {
    logseq.provideUI({
      key: slotUiKey(uiKeys.profileEditor, uiState.profileEditorSlot),
      slot: uiState.profileEditorSlot,
      reset: true,
      template: profileEditorTemplate()
    });
  } catch (error) {
    console.warn("[nextcloud-sync] dropped stale profile editor slot", error);
    uiState.profileEditorSlot = "";
  }
}

function syncTaskScopeManagerUI() {
  if (!uiState.taskScopeManagerSlot || typeof logseq.provideUI !== "function") return;
  let template = "";
  try {
    template = taskScopeManagerTemplate();
  } catch (error) {
    console.error("[nextcloud-sync] could not render sync profiles", error);
    const message = error instanceof Error ? error.message : String(error);
    template = `<div class="nextcloud-selector"><div class="nextcloud-selector__status">Sync profiles failed to render: ${escapeHtml(message)}</div></div>`;
  }
  try {
    logseq.provideUI({
      key: slotUiKey(uiKeys.taskScopeManager, uiState.taskScopeManagerSlot),
      slot: uiState.taskScopeManagerSlot,
      reset: true,
      template
    });
  } catch (error) {
    console.warn("[nextcloud-sync] dropped stale sync profiles slot", error);
    uiState.taskScopeManagerSlot = "";
  }
}

function providePluginStyle() {
  if (typeof logseq.provideStyle !== "function") return;
  logseq.provideStyle(`
    .nextcloud-selector {
      border: 1px solid var(--ls-border-color);
      border-radius: 12px;
      padding: 16px;
      background: var(--ls-secondary-background-color);
      margin: 8px 0;
    }
    .nextcloud-selector__header,
    .nextcloud-selector__actions {
      display: flex;
      gap: 8px;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
    }
    .nextcloud-selector__header {
      gap: 16px;
    }
    .nextcloud-selector__headline {
      font-weight: 700;
      margin-bottom: 4px;
    }
    .nextcloud-selector__hint,
    .nextcloud-selector__meta,
    .nextcloud-selector__status,
    .nextcloud-selector__empty {
      opacity: 0.8;
      font-size: 0.9em;
    }
    .nextcloud-selector__list {
      display: grid;
      gap: 12px;
      margin-top: 12px;
    }
    .nextcloud-selector__nav-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 12px;
      margin-top: 12px;
    }
    .nextcloud-selector__card {
      border: 1px solid var(--ls-border-color);
      border-radius: 10px;
      padding: 12px;
      background: var(--ls-primary-background-color);
    }
    .nextcloud-selector__title {
      font-weight: 600;
      margin-bottom: 6px;
    }
    .nextcloud-selector__section-title {
      margin-top: 16px;
      margin-bottom: 8px;
      font-weight: 700;
    }
    .nextcloud-selector button {
      appearance: none;
      border: 1px solid var(--ls-border-color);
      border-radius: 10px;
      background: var(--ls-primary-background-color);
      color: var(--ls-primary-text-color);
      padding: 8px 12px;
      font-weight: 600;
      cursor: pointer;
      transition: background 120ms ease, border-color 120ms ease, transform 120ms ease;
    }
    .nextcloud-selector button:hover {
      background: var(--ls-tertiary-background-color);
      border-color: var(--ls-link-text-color, var(--ls-border-color));
      transform: translateY(-1px);
    }
    .nextcloud-selector__primary-action {
      background: var(--ls-link-text-color, var(--ls-primary-background-color));
      color: var(--ls-primary-background-color);
      border-color: var(--ls-link-text-color, var(--ls-border-color));
      padding: 10px 16px;
    }
    .nextcloud-selector__nav-card {
      text-align: left;
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      gap: 4px;
      min-height: 88px;
      justify-content: center;
    }
    .nextcloud-selector__nav-title {
      font-weight: 700;
    }
    .nextcloud-selector__nav-hint {
      font-size: 0.9em;
      opacity: 0.8;
      line-height: 1.3;
      white-space: normal;
    }
    .nextcloud-selector__actions--profile {
      justify-content: flex-start;
      margin-top: 10px;
    }
  `);
}

function mountInlineUi() {
  if (typeof logseq.App?.onMacroRendererSlotted !== "function" || typeof logseq.provideUI !== "function") return;
  logseq.App.onMacroRendererSlotted(({ slot, payload }: any) => {
    const args = Array.isArray(payload?.arguments) ? payload.arguments : [];
    const macroName = args[0];
    if (macroName === ":nextcloud-sync-hub") {
      uiState.syncHubSlot = slot;
      scheduleSlottedRender(() => {
        if (uiState.syncHubSlot === slot) syncHubUI();
      });
      return;
    }
    if (macroName === ":nextcloud-profile-editor") {
      uiState.profileEditorSlot = slot;
      scheduleSlottedRender(() => {
        if (uiState.profileEditorSlot === slot) syncProfileEditorUI();
      });
      return;
    }
    if (macroName === ":nextcloud-calendar-selector") {
      uiState.calendarSelectorSlot = slot;
      scheduleSlottedRender(() => {
        if (uiState.calendarSelectorSlot === slot) syncCalendarSelectorUI();
      });
      return;
    }
    if (macroName === ":nextcloud-task-scope-manager") {
      uiState.taskScopeManagerSlot = slot;
      scheduleSlottedRender(() => {
        if (uiState.taskScopeManagerSlot === slot) syncTaskScopeManagerUI();
      });
    }
  });
}

function scheduleSlottedRender(render: () => void) {
  window.setTimeout(render, 60);
}

async function loadTasksSilently() {
  try {
    const activeProfile = getActiveSyncProfile(settings());
    const activeTaskScope = getActiveTaskScope(settings());
    const activeCalendarScope = getActiveCalendarScope(settings());
    state.tasks = activeTaskScope ? await collectLogseqTasksForScope(activeTaskScope, settings()) : await collectLogseqTasksWithSettings(settings());
    state.events = activeCalendarScope ? await collectLogseqCalendarEventsForScope(activeCalendarScope) : await collectLogseqCalendarEvents();
    state.activeTaskProfileId = activeProfile?.id || "";
    state.activeCalendarProfileId = activeProfile?.id || "";
  } catch (error) {
    console.warn("[nextcloud-sync] silent refresh failed", error);
  }
}

function scheduleStartupSync() {
  const current = settings();
  if (!current.syncOnStartup) return;

  // Let Logseq finish its own startup work before we do any graph-wide scans.
  window.setTimeout(() => {
    void syncOnStartup().catch((error) => {
      console.warn("[nextcloud-sync] startup sync failed", error);
    });
  }, 3000);
}

function clearProfileAutoSyncSchedules() {
  for (const timerId of runtimeState.profileSyncTimers.values()) {
    window.clearInterval(timerId);
  }
  runtimeState.profileSyncTimers.clear();
}

function refreshProfileAutoSyncSchedules() {
  clearProfileAutoSyncSchedules();
  const profiles = getPersistedSyncProfiles(settings()).filter(
    (profile) => profile.enabled && profile.updatePeriodically === true && (profile.remoteUrl || profile.taskListUrl || profile.calendarUrl)
  );

  for (const profile of profiles) {
    const minutes = Math.max(1, Number(profile.updateIntervalMinutes || 15) || 15);
    const timerId = window.setInterval(() => {
      void syncProfileEverything(profile).catch((error) => {
        console.warn(`[nextcloud-sync] periodic sync failed (${profile.name})`, error);
      });
    }, minutes * 60 * 1000);
    runtimeState.profileSyncTimers.set(profile.id, timerId);
  }
}

async function ensureCalendarSelectorPage() {
  await replacePageBlocks(
    calendarSelectorPage,
    [
      "Nextcloud calendar selection",
      "Use the picker below to discover and choose a Nextcloud calendar.",
      "{{renderer :nextcloud-calendar-selector}}"
    ],
    "[nextcloud-sync] could not replace calendar selector page content"
  );
}

async function ensureSyncHubPage() {
  await replacePageBlocks(
    syncHubPage,
    [
      "Nextcloud sync",
      "Use the controls below to open the main pages and sync each profile.",
      "{{renderer :nextcloud-sync-hub}}"
    ],
    "[nextcloud-sync] could not replace sync hub page content"
  );
}

async function ensureTaskScopeManagerPage() {
  await replacePageBlocks(
    taskScopeManagerPage,
    [
      "Nextcloud sync profiles",
      "Each profile decides what to sync and where it goes in Nextcloud.",
      "Remote collection URL: the Nextcloud destination used for both tasks and calendar events.",
      "Task page types and task tags: page-level rules that decide which notes belong to this profile.",
      "Ignored tags: tags that always exclude a page or block from this profile.",
      "Only scan matched pages: faster, but it skips pages that only match through block-level overrides.",
      "Calendar property override: optional extra include rule such as calendar:: personal.",
      "Write to journal: default is no. Yes lets dated imported items go to journal pages, no keeps imported items off journals.",
      "Default import page: leave this empty to use the profile inbox. Set it when you want imported tasks and events to land on a specific page instead.",
      "Simple checklist mode: default is no. Yes makes this profile task-only, skips calendar sync/import, and is meant for clean list pages such as shopping lists.",
      "Update periodically: default is no. Yes makes the profile sync automatically on a timer.",
      "Update interval minutes: only used when periodic updates are enabled.",
      "{{renderer :nextcloud-task-scope-manager}}"
    ],
    "[nextcloud-sync] could not replace task scope manager page content"
  );
}

async function ensureProfileEditorPage(profile: SyncProfileConfig) {
  await replacePageBlocks(
    profileEditorPage,
    [
      "Nextcloud sync profile editor",
      "Use the Save Profile button below after editing these blocks.",
      "About remote collection URL: this is the Nextcloud destination for both tasks and calendar events.",
      "About task page types and task tags: page-level rules for including notes in this profile.",
      "About ignored tags: tags that always exclude a page or block from this profile.",
      "About only scan matched pages: yes is faster, no keeps searching the whole graph for block-level overrides.",
      "About calendar property override: optional extra include rule such as calendar:: personal.",
      "About write to journal: default is no. Yes sends dated imported items to journal pages, no sends imported items to the default import page instead.",
      "About default import page: where imported tasks and events land when they are undated or when write to journal is set to no.",
      "About simple checklist mode: yes makes this profile task-only, skips calendar sync/import, and is meant for clean reusable checklists on a specific page.",
      "About update periodically: yes runs sync for this profile on a repeating timer.",
      "About update interval minutes: the number of minutes between automatic sync runs for this profile.",
      "{{renderer :nextcloud-profile-editor}}",
      `Profile id: ${profile.id}`,
      `Name: ${profile.name}`,
      `Remote collection URL: ${profile.remoteUrl || profile.taskListUrl || profile.calendarUrl || ""}`,
      `Task page types: ${profile.filterPageTypes || ""}`,
      `Task tags: ${profile.filterTags || ""}`,
      `Ignored tags: ${profile.ignoredTags || ""}`,
      `Only scan matched pages: ${profile.prefilterPagesOnly ? "yes" : "no"}`,
      `Calendar property key override: ${profile.propertyKey || "calendar"}`,
      `Calendar property value override: ${profile.propertyValue || ""}`,
      `Write to journal: ${profile.writeToJournal === true ? "yes" : "no"}`,
      `Default import page: ${profile.defaultImportPage || ""}`,
      `Simple checklist mode: ${profile.simpleChecklistMode === true ? "yes" : "no"}`,
      `Update periodically: ${profile.updatePeriodically === true ? "yes" : "no"}`,
      `Update interval minutes: ${Math.max(1, Number(profile.updateIntervalMinutes || 15) || 15)}`,
      `Enabled: ${profile.enabled ? "yes" : "no"}`
    ],
    "[nextcloud-sync] could not replace profile editor page content"
  );
}

async function importRemoteTaskItem(
  profile: SyncProfileConfig,
  item: RemoteTaskImportItem,
  importedUidIndex: Map<string, any>,
  journalTitleFormat: string
) {
  const editor = logseq.Editor as any;
  if (item.sourceBlockUuid && (await editor.getBlock?.(item.sourceBlockUuid, { includeChildren: false }))) {
    return false;
  }

  const marker = item.completed ? "DONE" : "TODO";
  const due = item.date ? ` DEADLINE: ${formatLogseqTimestamp(item.date, item.time, item.allDay)}` : "";
  const content = `${marker} ${item.title}${due}`.trim();
  const simpleChecklistMode = profile.simpleChecklistMode === true;
  const properties = {
    "nextcloud-profile": profile.id,
    "nextcloud-remote-uid": item.uid
  } as Record<string, string>;
  if (item.remoteResourceUrl) {
    properties["nextcloud_resource_url"] = item.remoteResourceUrl;
  }

  let pageName = importTargetPageName(profile, item.date, journalTitleFormat);
  const fallbackPageName = importFallbackPageName(profile);
  const existing =
    importedUidIndex.get(item.uid) ||
    (await findImportedBlockByUidOnPage(pageName, item.uid, profile.id)) ||
    (fallbackPageName !== pageName ? await findImportedBlockByUidOnPage(fallbackPageName, item.uid, profile.id) : null);

  if (!existing?.uuid && item.date && profile.writeToJournal !== false && (await journalMarkdownContainsRemoteUid(item.date, item.uid))) {
    return false;
  }

  if (simpleChecklistMode) {
    const plainExisting =
      (await findSimpleChecklistTaskByTitleOnPage(pageName, item.title)) ||
      (fallbackPageName !== pageName ? await findSimpleChecklistTaskByTitleOnPage(fallbackPageName, item.title) : null);

    if (plainExisting?.uuid) {
      const existingContent = String(plainExisting.content || "").trim();
      if (existingContent === content) {
        return false;
      }
      await editor.updateBlock?.(plainExisting.uuid, content);
      return true;
    }

    let targetPageName = pageName;
    try {
      if (item.date && profile.writeToJournal !== false) {
        await ensureJournalPageForImport(targetPageName);
      } else {
        await editor.createPage?.(targetPageName, {}, { redirect: false, createFirstBlock: false });
      }
    } catch (error) {
      console.warn(`[nextcloud-sync] simple checklist task import fell back from ${targetPageName} to ${fallbackPageName}`, error);
      targetPageName = fallbackPageName;
      await editor.createPage?.(targetPageName, {}, { redirect: false, createFirstBlock: false });
    }

    const insertParent = await findSimpleChecklistInsertParent(targetPageName);
    const created = insertParent?.uuid
      ? await editor.insertBlock?.(insertParent.uuid, content, { sibling: false, isPageBlock: false })
      : await editor.appendBlockInPage?.(targetPageName, content);
    if (created?.uuid) {
      return true;
    }
    return false;
  }

  if (existing?.uuid) {
    const existingProperties = await readImportedBlockPropertiesAsync(existing);
    const existingContent = String(existing.content || "").trim();
    const sameContent = existingContent === content;
    const sameProfile = getImportedProfileId(existingProperties) === profile.id;
    const sameUid = getImportedRemoteUid(existingProperties) === item.uid;
    const sameResourceUrl = getImportedProperty(existingProperties, "nextcloud_resource_url", "nextcloud-resource-url") === String(item.remoteResourceUrl || "").trim();
    if (sameContent && sameProfile && sameUid && sameResourceUrl) {
      return false;
    }
    await editor.updateBlock?.(existing.uuid, content);
    await editor.upsertBlockProperty?.(existing.uuid, "nextcloud-profile", profile.id);
    await editor.upsertBlockProperty?.(existing.uuid, "nextcloud-remote-uid", item.uid);
    if (item.remoteResourceUrl) {
      await editor.upsertBlockProperty?.(existing.uuid, "nextcloud_resource_url", item.remoteResourceUrl);
      await editor.removeBlockProperty?.(existing.uuid, "nextcloud-resource-url");
    }
    return false;
  }

  let parentUuid = "";
  try {
    if (item.date && profile.writeToJournal !== false) {
      await ensureJournalPageForImport(pageName);
    }
    parentUuid = await ensureImportSection(pageName, profile, "Tasks");
  } catch (error) {
    console.warn(`[nextcloud-sync] task import fell back from ${pageName} to ${fallbackPageName}`, error);
    pageName = fallbackPageName;
    parentUuid = await ensureImportSection(pageName, profile, "Tasks");
  }
  const created = await editor.insertBlock?.(parentUuid, content, { sibling: false, isPageBlock: false, properties });
  if (created?.uuid) {
    importedUidIndex.set(item.uid, created);
    return true;
  }
  return false;
}

async function importRemoteCalendarItem(
  profile: SyncProfileConfig,
  item: RemoteCalendarImportItem,
  importedUidIndex: Map<string, any>,
  journalTitleFormat: string,
  knownImportedUids: Set<string>
) {
  const editor = logseq.Editor as any;
  const sourceBlock =
    item.sourceBlockUuid
      ? await editor.getBlock?.(item.sourceBlockUuid, { includeChildren: true })
      : null;

  const properties: Record<string, string> = {
    "nextcloud-profile": profile.id,
    "nextcloud-remote-uid": item.uid
  };
  if (item.remoteResourceUrl) {
    properties["nextcloud_resource_url"] = item.remoteResourceUrl;
  }

  if (profile.propertyKey && profile.propertyValue) {
    properties[profile.propertyKey] = profile.propertyValue;
  }

  if (item.allDay) {
    properties.date = formatPropertyDate(item.date);
  } else {
    properties.start = formatPropertyDate(item.date, item.time);
    if (item.endDate) {
      properties.end = formatPropertyDate(item.endDate, item.endTime);
    }
  }

  let pageName = importTargetPageName(profile, item.date, journalTitleFormat);
  const fallbackPageName = importFallbackPageName(profile);
  const existingFromIndex = importedUidIndex.get(item.uid);
  const existingOnPage = await findImportedBlockByUidOnPage(pageName, item.uid, profile.id);
  const existingByIdentity = await findImportedEventByIdentityOnPage(pageName, profile.id, item.title, item.date, item.time);
  const existingInSection = await findImportedEventInSection(pageName, profile, item);
  const existingOnFallback =
    fallbackPageName !== pageName ? await findImportedBlockByUidOnPage(fallbackPageName, item.uid, profile.id) : null;
  const existing = sourceBlock || existingFromIndex || existingOnPage || existingByIdentity || existingInSection || existingOnFallback;
  const fileHasRemoteUid =
    !existing?.uuid && item.date && profile.writeToJournal !== false
      ? await journalMarkdownContainsRemoteUid(item.date, item.uid)
      : false;

  if (fileHasRemoteUid) {
    knownImportedUids.add(item.uid);
    return false;
  }
  if (existing?.uuid) {
    const existingProperties = await readImportedBlockPropertiesAsync(existing);
    const existingTitle = String(existing.content || "").trim();
    const expectedDate = item.allDay ? formatPropertyDate(item.date) : "";
    const expectedStart = item.allDay ? "" : formatPropertyDate(item.date, item.time);
    const expectedEnd = item.allDay ? "" : String(item.endDate ? formatPropertyDate(item.endDate, item.endTime) : "");
    const sameTitle = existingTitle === String(item.title || "").trim();
    const sameProfile = getImportedProfileId(existingProperties) === profile.id;
    const sameUid = getImportedRemoteUid(existingProperties) === item.uid;
    const sameResourceUrl = getImportedProperty(existingProperties, "nextcloud_resource_url", "nextcloud-resource-url") === String(item.remoteResourceUrl || "").trim();
    const sameDate = String(getImportedProperty(existingProperties, "date") || "").trim() === expectedDate;
    const sameStart = String(getImportedProperty(existingProperties, "start") || "").trim() === expectedStart;
    const sameEnd = String(getImportedProperty(existingProperties, "end") || "").trim() === expectedEnd;

    if (sameTitle && sameProfile && sameUid && sameResourceUrl && sameDate && sameStart && sameEnd) {
      knownImportedUids.add(item.uid);
      return false;
    }

    await editor.updateBlock?.(existing.uuid, item.title);
    await editor.upsertBlockProperty?.(existing.uuid, "nextcloud-profile", profile.id);
    await editor.upsertBlockProperty?.(existing.uuid, "nextcloud-remote-uid", item.uid);
    if (item.remoteResourceUrl) {
      await editor.upsertBlockProperty?.(existing.uuid, "nextcloud_resource_url", item.remoteResourceUrl);
      await editor.removeBlockProperty?.(existing.uuid, "nextcloud-resource-url");
    }
    if (item.allDay) {
      await editor.upsertBlockProperty?.(existing.uuid, "date", expectedDate);
      await editor.removeBlockProperty?.(existing.uuid, "start");
      await editor.removeBlockProperty?.(existing.uuid, "end");
    } else {
      await editor.upsertBlockProperty?.(existing.uuid, "start", expectedStart);
      if (item.endDate) {
        await editor.upsertBlockProperty?.(existing.uuid, "end", expectedEnd);
      }
      await editor.removeBlockProperty?.(existing.uuid, "date");
    }
    knownImportedUids.add(item.uid);
    return true;
  }

  if (knownImportedUids.has(item.uid)) {
    return false;
  }

  let parentUuid = "";
  try {
    if (item.date && profile.writeToJournal !== false) {
      await ensureJournalPageForImport(pageName);
    }
    parentUuid = await ensureImportSection(pageName, profile, "Events");
  } catch (error) {
    console.warn(`[nextcloud-sync] calendar import fell back from ${pageName} to ${fallbackPageName}`, error);
    pageName = fallbackPageName;
    parentUuid = await ensureImportSection(pageName, profile, "Events");
  }
  const created = await editor.insertBlock?.(parentUuid, item.title, { sibling: false, isPageBlock: false, properties });
  if (created?.uuid) {
    importedUidIndex.set(item.uid, created);
    knownImportedUids.add(item.uid);
    return true;
  }
  return false;
}

async function deleteImportedCalendarItem(
  uid: string,
  importedUidIndex: Map<string, any>
) {
  const editor = logseq.Editor as any;
  const existing = importedUidIndex.get(uid);
  if (!existing?.uuid) return false;
  await editor.removeBlock?.(existing.uuid);
  importedUidIndex.delete(uid);
  return true;
}

async function importRemoteTasksForProfile(profile: SyncProfileConfig) {
  if (!profile.remoteUrl) {
    logseq.UI.showMsg?.(`Select a remote collection for ${profile.name} first.`, "warning", { timeout: 6000 });
    return;
  }

  const currentSettings = settings();
  const graphConfig = await logseq.App.getCurrentGraphConfigs?.();
  const journalTitleFormat = readGraphConfigJournalTitleFormat(graphConfig) || "MMM do, yyyy";
  const importedUidIndex = await buildImportedUidIndex();
  const scopeSettings = withProfileSettings(currentSettings, profile);
  const remoteTasks = await fetchRemoteTasksForImport(scopeSettings, profile.remoteUrl);

  let importedTasks = 0;

  for (const item of remoteTasks) {
    if (await importRemoteTaskItem(profile, item, importedUidIndex, journalTitleFormat)) {
      importedTasks += 1;
    }
  }

  logseq.UI.showMsg?.(
    `${profile.name}: imported ${importedTasks} tasks from Nextcloud.`,
    importedTasks ? "success" : "warning",
    { timeout: 7000 }
  );
}

async function importRemoteItemsForProfile(profile: SyncProfileConfig) {
  if (!profile.remoteUrl) {
    logseq.UI.showMsg?.(`Select a remote collection for ${profile.name} first.`, "warning", { timeout: 6000 });
    return;
  }

  await normalizeImportedBlockProperties();
  const currentSettings = settings();
  const graphConfig = await logseq.App.getCurrentGraphConfigs?.();
  const journalTitleFormat = readGraphConfigJournalTitleFormat(graphConfig) || "MMM do, yyyy";
  const importedUidIndex = await buildImportedUidIndex();
  const knownImportedCalendarUids = readImportedCalendarUidCache(profile.id, currentSettings);
  const scopeSettings = withProfileSettings(currentSettings, profile);
  const simpleChecklistMode = profile.simpleChecklistMode === true;

  const remoteTasks = await fetchRemoteTasksForImport(scopeSettings, profile.remoteUrl);
  const remoteEvents = simpleChecklistMode ? [] : await fetchRemoteCalendarEventsForImport(scopeSettings, profile.remoteUrl);

  let importedTasks = 0;
  let importedEvents = 0;

  for (const item of remoteTasks) {
    if (await importRemoteTaskItem(profile, item, importedUidIndex, journalTitleFormat)) {
      importedTasks += 1;
    }
  }

  for (const item of remoteEvents) {
    if (await importRemoteCalendarItem(profile, item, importedUidIndex, journalTitleFormat, knownImportedCalendarUids)) {
      importedEvents += 1;
    }
  }

  if (!simpleChecklistMode) {
    await writeImportedCalendarUidCache(profile.id, knownImportedCalendarUids, currentSettings);
  }

  if (simpleChecklistMode) {
    logseq.UI.showMsg?.(
      `${profile.name}: imported ${importedTasks} tasks from Nextcloud.`,
      importedTasks ? "success" : "warning",
      { timeout: 7000 }
    );
  } else {
    logseq.UI.showMsg?.(
      `${profile.name}: imported ${importedTasks} tasks and ${importedEvents} events from Nextcloud.`,
      importedTasks || importedEvents ? "success" : "warning",
      { timeout: 7000 }
    );
  }
}

async function importRemoteItems(scope = getActiveSyncProfile(settings())) {
  if (!scope) {
    logseq.UI.showMsg?.("Create or select a sync profile first.", "warning", { timeout: 5000 });
    return;
  }

  try {
    await importRemoteItemsForProfile(scope);
  } catch (error) {
    console.error("[nextcloud-sync] remote import failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Remote import failed.", "error", { timeout: 7000 });
  }
}

async function importRemoteItemsForAllProfiles() {
  const profiles = getSyncProfiles(settings()).filter((profile) => profile.enabled);
  if (!profiles.length) {
    logseq.UI.showMsg?.("No enabled sync profiles found.", "warning", { timeout: 5000 });
    return;
  }

  for (const profile of profiles) {
    await importRemoteItemsForProfile(profile);
  }
}

async function refreshTasks(scope = getActiveTaskScope(settings())) {
  try {
    await normalizeImportedBlockProperties();
    const current = settings();
    const tasks = scope ? await collectLogseqTasksForScope(scope, current) : await collectLogseqTasksWithSettings(current);
    state.tasks = tasks;
    state.activeTaskProfileId = scope?.id || getActiveSyncProfile(current)?.id || "";
    await writeTaskOverviewPage(tasks, scope ?? undefined);
    logseq.UI.showMsg?.(
      `Indexed ${tasks.length} scoped Logseq tasks${scope ? ` for ${scope.name}` : ""}.`,
      tasks.length ? "success" : "warning",
      { timeout: 5000 }
    );
    return tasks;
  } catch (error) {
    console.error("[nextcloud-sync] refresh failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Task refresh failed.", "error", { timeout: 6000 });
    return [];
  }
}

async function refreshCalendar(scope = getActiveCalendarScope(settings())) {
  try {
    await normalizeImportedBlockProperties();
    const collectorScope = scope ? { ...scope, prefilterPagesOnly: false } : undefined;
    const events = collectorScope ? await collectLogseqCalendarEventsForScope(collectorScope) : await collectLogseqCalendarEvents();
    state.events = events;
    state.activeCalendarProfileId = scope?.id || getActiveSyncProfile(settings())?.id || "";
    await writeCalendarOverviewPage(events, scope ?? undefined);
    logseq.UI.showMsg?.(
      `Indexed ${events.length} Logseq calendar events${scope ? ` for ${scope.name}` : ""}.`,
      events.length ? "success" : "warning",
      { timeout: 5000 }
    );
    return events;
  } catch (error) {
    console.error("[nextcloud-sync] calendar refresh failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Calendar refresh failed.", "error", { timeout: 6000 });
    return [];
  }
}

async function discoverCalendars() {
  const result = await discoverCalDavCalendars(settings());
  state.calendarDiscovery = result;
  await ensureSyncHubPage();
  await ensureCalendarSelectorPage();
  syncCalendarSelectorUI();
  syncHubUI();
  openPage(calendarSelectorPage);
  logseq.UI.showMsg?.(result.message, result.ok ? "success" : "warning", { timeout: 7000 });
  return result;
}

async function discoverTaskLists() {
  const result = await discoverCalDavTaskLists(settings());
  state.taskListDiscovery = result;
  await ensureSyncHubPage();
  await ensureTaskScopeManagerPage();
  syncTaskScopeManagerUI();
  syncHubUI();
  openPage(taskScopeManagerPage);
  logseq.UI.showMsg?.(result.message, result.ok ? "success" : "warning", { timeout: 7000 });
  return result;
}

async function exportTasks(scope = getActiveTaskScope(settings())) {
  const tasks =
    state.tasks.length && state.activeTaskProfileId === (scope?.id || "") ? state.tasks : await refreshTasks(scope);
  if (!tasks.length) {
    logseq.UI.showMsg?.("No tasks found to export.", "warning");
    return;
  }

  const { filename } = await exportTaskListIcs(tasks, settings().calendarTimezone);
  logseq.UI.showMsg?.(`Exported ${tasks.length} tasks to ${filename}${scope ? ` for ${scope.name}` : ""}`, "success", { timeout: 4000 });
}

async function exportCalendar() {
  const activeScope = getActiveCalendarScope(settings());
  const events = state.events.length && state.activeCalendarProfileId === (activeScope?.id || "") ? state.events : await refreshCalendar(activeScope);
  if (!events.length) {
    logseq.UI.showMsg?.("No calendar events found to export.", "warning");
    return;
  }

  const { filename } = await exportCalendarIcs(events, settings().calendarTimezone);
  logseq.UI.showMsg?.(
    `Exported ${events.length} calendar events to ${filename}${activeScope ? ` for ${activeScope.name}` : ""}`,
    "success",
    { timeout: 4000 }
  );
}

async function syncTasks(scope = getActiveTaskScope(settings())) {
  if (!scope) {
    logseq.UI.showMsg?.("Create or select a sync profile first.", "warning", { timeout: 5000 });
    return;
  }
  if (!scope.taskListUrl) {
    logseq.UI.showMsg?.(`Select or create a remote collection for ${scope.name} first.`, "warning", { timeout: 6000 });
    return;
  }

  const tasks = await refreshTasks(scope);
  if (!tasks.length) {
    logseq.UI.showMsg?.("No tasks found to sync.", "warning");
    return;
  }

  try {
    console.warn(`[nextcloud-sync] task sync start (${scope.name})`);
    const currentSettings = settings();
    const linkedProfile = getPersistedSyncProfiles(currentSettings).find((item) => item.id === scope.id);
    const syncSettings = linkedProfile ? withProfileSettings(currentSettings, linkedProfile) : withScopeSettings(currentSettings, scope);
    const result = await syncTasksToCalDav(tasks, syncSettings);
    const summary = `Synced ${result.synced}, verified ${result.verified}, deleted ${result.deleted}, mirrored ${result.completedRemote}, failed ${result.failed}`;
    state.tasks = result.tasks;
    state.activeTaskProfileId = scope.id;
    await writeTaskOverviewPage(result.tasks, scope);
    console.warn(`[nextcloud-sync] task sync finish (${scope.name})`, {
      synced: result.synced,
      verified: result.verified,
      deleted: result.deleted,
      mirrored: result.completedRemote,
      failed: result.failed,
      target: result.calendarUrl
    });
    if (result.errors.length) {
      logseq.UI.showMsg?.(`${scope.name}: ${summary}. ${result.errors.join(" | ")}`, "warning", { timeout: 10000 });
    } else {
      logseq.UI.showMsg?.(`${scope.name}: ${summary}. Target: ${result.calendarUrl}`, "success", { timeout: 10000 });
    }
  } catch (error) {
    console.error("[nextcloud-sync] sync failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Task sync failed.", "error", { timeout: 6000 });
  }
}

async function syncProfileEverything(profile: SyncProfileConfig) {
  if (runtimeState.syncingProfileIds.has(profile.id)) {
    console.warn(`[nextcloud-sync] skipped overlapping profile sync (${profile.name})`);
    return;
  }
  runtimeState.syncingProfileIds.add(profile.id);
  console.warn(`[nextcloud-sync] profile sync start (${profile.name})`);
  try {
    await syncTasks({
      id: profile.id,
      name: profile.name,
      taskListUrl: profile.remoteUrl || profile.taskListUrl,
      filterPageTypes: profile.filterPageTypes,
      filterTags: profile.filterTags,
      ignoredTags: profile.ignoredTags,
      prefilterPagesOnly: profile.prefilterPagesOnly === true,
      enabled: profile.enabled
    });

    if (profile.simpleChecklistMode !== true && (profile.remoteUrl || profile.calendarUrl)) {
      await syncCalendar({
        id: profile.id,
        name: profile.name,
        calendarUrl: profile.remoteUrl || profile.calendarUrl,
        filterPageTypes: profile.filterPageTypes,
        filterTags: profile.filterTags,
        ignoredTags: profile.ignoredTags,
        prefilterPagesOnly: profile.prefilterPagesOnly === true,
        propertyKey: profile.propertyKey,
        propertyValue: profile.propertyValue,
        enabled: profile.enabled
      });
    }

    if (profile.simpleChecklistMode === true) {
      await importRemoteTasksForProfile(profile);
      console.warn(`[nextcloud-sync] profile sync finish (${profile.name})`);
      return;
    }

    await importRemoteTasksForProfile(profile);
    console.warn(`[nextcloud-sync] profile sync finish (${profile.name})`);
  } finally {
    runtimeState.syncingProfileIds.delete(profile.id);
  }
}

async function syncAllTaskScopes() {
  const profiles = getSyncProfiles(settings()).filter((profile) => profile.enabled);
  if (!profiles.length) {
    logseq.UI.showMsg?.("No enabled sync profiles found.", "warning", { timeout: 5000 });
    return;
  }

  for (const profile of profiles) {
    await syncProfileEverything(profile);
  }
  syncTaskScopeManagerUI();
  syncHubUI();
}

async function syncCalendar(scope = getActiveCalendarScope(settings())) {
  if (!scope) {
    logseq.UI.showMsg?.("Create or select a sync profile first.", "warning", { timeout: 5000 });
    return;
  }
  const profileForMode = getPersistedSyncProfiles(settings()).find((item) => item.id === scope.id);
  if (profileForMode?.simpleChecklistMode === true) {
    logseq.UI.showMsg?.(`${scope.name}: simple checklist mode skips calendar sync and only uses VTODO tasks.`, "warning", {
      timeout: 6000
    });
    return;
  }
  if (!scope.calendarUrl) {
    logseq.UI.showMsg?.(`Select a remote collection for ${scope.name} first.`, "warning", { timeout: 6000 });
    return;
  }

  const currentSettings = settings();
  const profile = profileForMode ?? getPersistedSyncProfiles(currentSettings).find((item) => item.id === scope.id);
  const graphConfig = await logseq.App.getCurrentGraphConfigs?.();
  const journalTitleFormat = readGraphConfigJournalTitleFormat(graphConfig) || "MMM do, yyyy";
  const importedUidIndex = await buildImportedUidIndex();
  const knownImportedCalendarUids = profile ? readImportedCalendarUidCache(profile.id, currentSettings) : new Set<string>();
  const remoteItems =
    profile?.remoteUrl
      ? await fetchRemoteCalendarEventsForImport(withProfileSettings(currentSettings, profile), profile.remoteUrl)
      : [];
  const snapshotState = readCalendarSyncState(currentSettings);
  const profileSnapshots = { ...(snapshotState[scope.id] || {}) };

  let events = await refreshCalendar(scope);
  const localByUid = new Map(events.map((event) => [event.uid, event]));
  const remoteByUid = new Map(remoteItems.map((item) => [item.uid, item]));
  let importedRemoteChanges = 0;
  let removedRemoteDeletes = 0;
  let conflicts = 0;
  const skipPutUids = new Set<string>();

  if (profile) {
    for (const [uid, remoteItem] of remoteByUid) {
      const localEvent = localByUid.get(uid);
      const lastSnapshot = profileSnapshots[uid];
      const remoteSnapshot = calendarSnapshotFromRemote(remoteItem);

      if (!localEvent) {
        if (await importRemoteCalendarItem(profile, remoteItem, importedUidIndex, journalTitleFormat, knownImportedCalendarUids)) {
          importedRemoteChanges += 1;
        }
        profileSnapshots[uid] = remoteSnapshot;
        continue;
      }

      const localSnapshot = calendarSnapshotFromLocal(localEvent);
      const localChanged = !lastSnapshot || !sameCalendarSnapshot(localSnapshot, lastSnapshot);
      const remoteChanged = !lastSnapshot || !sameCalendarSnapshot(remoteSnapshot, lastSnapshot);

      if (!lastSnapshot && !sameCalendarSnapshot(localSnapshot, remoteSnapshot)) {
        conflicts += 1;
        continue;
      }

      if (remoteChanged && !localChanged) {
        const imported = await importRemoteCalendarItem(
          profile,
          remoteItem,
          importedUidIndex,
          journalTitleFormat,
          knownImportedCalendarUids
        );
        let reconciled = false;
        if (imported) {
          importedRemoteChanges += 1;
          events = await refreshCalendar(scope);
          localByUid.clear();
          for (const event of events) {
            localByUid.set(event.uid, event);
          }
          const refreshedLocal = localByUid.get(uid);
          reconciled = Boolean(refreshedLocal) && sameCalendarSnapshot(calendarSnapshotFromLocal(refreshedLocal), remoteSnapshot);
        }
        if (reconciled) {
          profileSnapshots[uid] = remoteSnapshot;
        } else {
          skipPutUids.add(uid);
          conflicts += 1;
        }
        continue;
      }

      if (remoteChanged && localChanged && !sameCalendarSnapshot(localSnapshot, remoteSnapshot)) {
        conflicts += 1;
        continue;
      }
    }

    for (const [uid, localEvent] of localByUid) {
      if (remoteByUid.has(uid)) continue;
      const lastSnapshot = profileSnapshots[uid];
      if (!lastSnapshot) continue;
      const localSnapshot = calendarSnapshotFromLocal(localEvent);
      const localChanged = !sameCalendarSnapshot(localSnapshot, lastSnapshot);

      if (localChanged) {
        skipPutUids.add(uid);
        conflicts += 1;
        continue;
      }

      const removed = await deleteImportedCalendarItem(uid, importedUidIndex);
      if (removed) {
        removedRemoteDeletes += 1;
        delete profileSnapshots[uid];
        knownImportedCalendarUids.delete(uid);
      } else {
        skipPutUids.add(uid);
        conflicts += 1;
      }
    }

    if (importedRemoteChanges || removedRemoteDeletes) {
      await writeImportedCalendarUidCache(profile.id, knownImportedCalendarUids, currentSettings);
      events = await refreshCalendar(scope);
    }
  }

  if (!events.length) {
    logseq.UI.showMsg?.("No calendar events found to sync.", "warning");
    return;
  }

  try {
    console.warn(`[nextcloud-sync] calendar sync start (${scope.name})`);
    const eventsToSync = events.filter((event) => !skipPutUids.has(event.uid));
    const result = await syncCalendarToCalDav(eventsToSync, withCalendarScopeSettings(settings(), scope), skipPutUids);
    const summary = `Synced ${result.synced}, verified ${result.verified}, deleted ${result.deleted}, failed ${result.failed}`;
    state.events = result.events;
    state.activeCalendarProfileId = scope.id;
    await writeCalendarOverviewPage(result.events, scope);
    const nextSnapshots = { ...profileSnapshots };
    for (const event of result.events) {
      nextSnapshots[event.uid] = calendarSnapshotFromLocal(event);
    }
    const nextState = readCalendarSyncState(settings());
    nextState[scope.id] = nextSnapshots;
    await writeCalendarSyncState(nextState);
    console.warn(`[nextcloud-sync] calendar sync finish (${scope.name})`, {
      synced: result.synced,
      verified: result.verified,
      deleted: result.deleted,
      failed: result.failed,
      target: result.calendarUrl
    });
    if (result.errors.length) {
      logseq.UI.showMsg?.(`${scope.name}: ${summary}. ${result.errors.join(" | ")}`, "warning", { timeout: 10000 });
    } else if (conflicts) {
      logseq.UI.showMsg?.(`${scope.name}: ${summary}. Skipped ${conflicts} calendar conflicts where both Logseq and Nextcloud changed.`, "warning", { timeout: 10000 });
    } else if (importedRemoteChanges || removedRemoteDeletes) {
      const parts = [];
      if (importedRemoteChanges) parts.push(`pulled ${importedRemoteChanges} remote calendar updates into Logseq first`);
      if (removedRemoteDeletes) parts.push(`removed ${removedRemoteDeletes} Logseq events deleted in Nextcloud`);
      logseq.UI.showMsg?.(`${scope.name}: ${summary}. ${parts.join(". ")}.`, "success", { timeout: 10000 });
    } else {
      logseq.UI.showMsg?.(`${scope.name}: ${summary}. Target: ${result.calendarUrl}`, "success", { timeout: 10000 });
    }
  } catch (error) {
    console.error("[nextcloud-sync] calendar sync failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Calendar sync failed.", "error", { timeout: 6000 });
  }
}

async function testConnection(scope = getActiveTaskScope(settings())) {
  if (!scope) {
    logseq.UI.showMsg?.("Create or select a sync profile first.", "warning", { timeout: 5000 });
    return;
  }
  try {
    const result = await testCalDavConnectionForUrl(scope.taskListUrl, settings());
    if (result.ok) {
      logseq.UI.showMsg?.(`Remote collection reachable: ${result.url}`, "success", { timeout: 4000 });
    } else {
      logseq.UI.showMsg?.(result.message, "warning", { timeout: 6000 });
    }
  } catch (error) {
    console.error("[nextcloud-sync] connection test failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Connection test failed.", "error", { timeout: 6000 });
  }
}

async function testCalendarCollection() {
  const activeScope = getActiveCalendarScope(settings());
  if (!activeScope) {
    logseq.UI.showMsg?.("Create or select a sync profile first.", "warning", { timeout: 5000 });
    return;
  }
  try {
    const result = await testCalendarConnectionForUrl(activeScope.calendarUrl, settings());
    if (result.ok) {
      logseq.UI.showMsg?.(`${activeScope.name}: Calendar reachable: ${result.url}`, "success", { timeout: 4000 });
    } else {
      logseq.UI.showMsg?.(result.message, "warning", { timeout: 6000 });
    }
  } catch (error) {
    console.error("[nextcloud-sync] calendar connection test failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Calendar connection test failed.", "error", { timeout: 6000 });
  }
}

async function selectCalendarOption(e?: any) {
  const url =
    e?.dataset?.url ??
    e?.currentTarget?.dataset?.url ??
    e?.target?.dataset?.url ??
    "";
  const selected = String(url ?? "").trim();
  if (!selected) {
    logseq.UI.showMsg?.("No remote collection URL was provided by the selector.", "warning", { timeout: 5000 });
    return;
  }
  const current = settings();
  const scopes = getPersistedCalendarScopes(current);
  const active = getActiveCalendarScope(current);
  if (active && active.id !== "legacy-default") {
    await saveCalendarScopes(
      scopes.map((scope) =>
        scope.id === active.id ? { ...scope, remoteUrl: selected, taskListUrl: selected, calendarUrl: selected } : scope
      ),
      active.id
    );
    syncCalendarSelectorUI();
    logseq.UI.showMsg?.(`Assigned remote collection to ${active.name}.`, "success", { timeout: 6000 });
    return;
  }
  logseq.updateSettings?.({ caldavCalendarUrl: selected, caldavTaskListUrl: selected });
  syncCalendarSelectorUI();
  logseq.UI.showMsg?.(`Saved remote collection URL: ${selected}`, "success", { timeout: 6000 });
}

function readDatasetValue(e: any, key: string) {
  return (
    e?.dataset?.[key] ??
    e?.currentTarget?.dataset?.[key] ??
    e?.target?.dataset?.[key] ??
    ""
  );
}

function getEventElement(e: any): HTMLElement | null {
  const target = e?.currentTarget ?? e?.target ?? null;
  return target instanceof HTMLElement ? target : null;
}

function findProfileCardElement(e: any, scopeId?: string): HTMLElement | null {
  const fromEvent = (getEventElement(e)?.closest?.(".nextcloud-profile-card") as HTMLElement | null | undefined) ?? null;
  if (fromEvent) return fromEvent;
  if (!scopeId || typeof document === "undefined") return null;
  const cards = Array.from(document.querySelectorAll<HTMLElement>(".nextcloud-profile-card"));
  return cards.find((card) => card.dataset?.scopeId === scopeId) ?? null;
}

function readProfileInputValue(e: any, scopeId: string, field: string) {
  const card = findProfileCardElement(e, scopeId);
  const input = card?.querySelector?.(`[data-profile-field="${field}"]`) as HTMLInputElement | null;
  if (!input) return "";
  if (input.type === "checkbox") {
    return input.checked ? "true" : "false";
  }
  return String(input.value ?? "").trim();
}

async function activateTaskScope(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
  const profiles = getPersistedSyncProfiles(settings());
  const target = profiles.find((profile) => profile.id === scopeId);
  if (!target) {
    logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
    return;
  }
  await saveSyncProfiles(profiles, target.id);
  syncTaskScopeManagerUI();
  syncHubUI();
  await refreshTasks({
    id: target.id,
    name: target.name,
    taskListUrl: target.taskListUrl,
    filterPageTypes: target.filterPageTypes,
    filterTags: target.filterTags,
    ignoredTags: target.ignoredTags,
    prefilterPagesOnly: target.prefilterPagesOnly === true,
    enabled: target.enabled
  });
  logseq.UI.showMsg?.(`Active sync profile: ${target.name}`, "success", { timeout: 4000 });
}

async function assignTaskListToActiveScope(e?: any) {
  const url = String(readDatasetValue(e, "url") || "").trim();
  if (!url) {
    logseq.UI.showMsg?.("No remote collection URL was provided by the selector.", "warning", { timeout: 5000 });
    return;
  }

  const current = settings();
  const profiles = getPersistedSyncProfiles(current);
  const active = getActiveSyncProfile(current);
  if (!active || active.id === "legacy-default") {
    logseq.UI.showMsg?.("Create a named sync profile first, then assign a remote collection to it.", "warning", { timeout: 6000 });
    return;
  }

  const updated = profiles.map((profile) =>
    profile.id === active.id ? { ...profile, remoteUrl: url, taskListUrl: url, calendarUrl: url } : profile
  );
  await saveSyncProfiles(updated, active.id);
  syncTaskScopeManagerUI();
  syncHubUI();
  logseq.UI.showMsg?.(`Assigned remote collection to ${active.name}.`, "success", { timeout: 5000 });
}

async function activateCalendarScope(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
  const profiles = getPersistedSyncProfiles(settings());
  const target = profiles.find((profile) => profile.id === scopeId);
  if (!target) {
    logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
    return;
  }
  await saveSyncProfiles(profiles, target.id);
  syncCalendarSelectorUI();
  syncHubUI();
  await refreshCalendar({
    id: target.id,
    name: target.name,
    calendarUrl: target.remoteUrl || target.calendarUrl,
    filterPageTypes: target.filterPageTypes,
    filterTags: target.filterTags,
    ignoredTags: target.ignoredTags,
    prefilterPagesOnly: target.prefilterPagesOnly === true,
    propertyKey: target.propertyKey,
    propertyValue: target.propertyValue,
    enabled: target.enabled
  });
  logseq.UI.showMsg?.(`Active sync profile: ${target.name}`, "success", { timeout: 4000 });
}

async function createCalendarScope() {
  const current = settings();
  const active = getActiveSyncProfile(current);
  const profiles = getPersistedSyncProfiles(current);
  const existingNames = new Set(profiles.map((profile) => normalizeProfileName(profile.name)).filter(Boolean));
  let sequence = profiles.length + 1;
  let name = `Profile ${sequence}`;
  while (existingNames.has(normalizeProfileName(name))) {
    sequence += 1;
    name = `Profile ${sequence}`;
  }

  const taskPageTypes = active?.filterPageTypes || "";
  const taskTags = active?.filterTags || "";
  const ignoredTags = active?.ignoredTags || "";
  const prefilterPagesOnly = active?.prefilterPagesOnly === true;
  const propertyKey = active?.propertyKey || "calendar";
  const propertyValue = slugify(name).replace(/-/g, " ") || name.toLowerCase();
  const writeToJournal = active?.writeToJournal === true;
  const defaultImportPage = active?.defaultImportPage || "";
  const simpleChecklistMode = active?.simpleChecklistMode === true;
  const updatePeriodically = active?.updatePeriodically === true;
  const updateIntervalMinutes = Math.max(1, Number(active?.updateIntervalMinutes || 15) || 15);
  const scope: SyncProfileConfig = {
    id: createScopeId(name),
    name,
    remoteUrl: active?.remoteUrl || active?.taskListUrl || active?.calendarUrl || "",
    taskListUrl: active?.remoteUrl || active?.taskListUrl || active?.calendarUrl || "",
    filterPageTypes: taskPageTypes,
    filterTags: taskTags,
    ignoredTags,
    prefilterPagesOnly,
    calendarUrl: active?.remoteUrl || active?.calendarUrl || active?.taskListUrl || "",
    propertyKey,
    propertyValue,
    writeToJournal,
    defaultImportPage,
    simpleChecklistMode,
    updatePeriodically,
    updateIntervalMinutes,
    enabled: true
  };
  await saveSyncProfiles([...profiles, scope], scope.id);
  await ensureSyncHubPage();
  await ensureCalendarSelectorPage();
  await ensureTaskScopeManagerPage();
  syncTaskScopeManagerUI();
  syncCalendarSelectorUI();
  syncHubUI();
  openPage(taskScopeManagerPage);
  logseq.UI.showMsg?.(`Created sync profile ${name}. Use Edit to refine filters and calendar rules.`, "success", { timeout: 6000 });
}

async function editCalendarScope(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim() || getActiveCalendarScope(settings())?.id || "";
  const profiles = getPersistedSyncProfiles(settings());
  const scope = profiles.find((item) => item.id === scopeId);
  if (!scope) {
    logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
    return;
  }
  await ensureProfileEditorPage(scope);
  openPage(profileEditorPage);
  logseq.UI.showMsg?.(`Editing sync profile ${scope.name}.`, "success", { timeout: 4000 });
}

async function refreshCalendarScopePreview(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
  const scope = getCalendarScopes(settings()).find((item) => item.id === scopeId) ?? getActiveCalendarScope(settings());
  if (!scope) {
    logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
    return;
  }
  const profile = getPersistedSyncProfiles(settings()).find((item) => item.id === scope.id);
  if (profile?.simpleChecklistMode === true) {
    logseq.UI.showMsg?.(`${scope.name}: simple checklist mode skips calendar preview.`, "warning", { timeout: 5000 });
    return;
  }
  await refreshCalendar(scope);
  openPage(settings().calendarPageName || defaultSettings.calendarPageName);
}

async function syncCalendarScope(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
  const scope = getCalendarScopes(settings()).find((item) => item.id === scopeId) ?? getActiveCalendarScope(settings());
  if (!scope) {
    logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
    return;
  }
  await syncCalendar(scope);
}

async function syncAllCalendarScopes() {
  const scopes = getCalendarScopes(settings()).filter((scope) => scope.enabled);
  if (!scopes.length) {
    logseq.UI.showMsg?.("No enabled sync profiles found.", "warning", { timeout: 5000 });
    return;
  }

  for (const scope of scopes) {
    await syncCalendar(scope);
  }
  syncCalendarSelectorUI();
  syncHubUI();
}

async function createTaskScope() {
  await createCalendarScope();
}

async function editTaskScope(e?: any) {
  await editCalendarScope(e);
}

function readProfileEditorValue(lines: string[], label: string) {
  const prefix = `${label}:`;
  const line = [...lines].reverse().find((item) => item.startsWith(prefix));
  return line ? line.slice(prefix.length).trim() : "";
}

function flattenBlockContents(blocks: any[]): string[] {
  const lines: string[] = [];

  const visit = (block: any) => {
    const content = String(block?.content || "").trim();
    if (content) lines.push(content);
    if (Array.isArray(block?.children)) {
      for (const child of block.children) visit(child);
    }
  };

  for (const block of blocks) visit(block);
  return lines;
}

async function saveProfileEditor() {
  const editor = logseq.Editor as any;
  const blocks = (await editor.getPageBlocksTree?.(profileEditorPage)) ?? [];
  const lines = flattenBlockContents(blocks);

  const profileId = readProfileEditorValue(lines, "Profile id");
  if (!profileId) {
    logseq.UI.showMsg?.("Profile editor is missing a profile id.", "warning", { timeout: 5000 });
    return;
  }

  const profiles = getPersistedSyncProfiles(settings());
  const scope = profiles.find((item) => item.id === profileId);
  if (!scope) {
    logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
    return;
  }

  const name = readProfileEditorValue(lines, "Name") || scope.name;
  const remoteUrl = readProfileEditorValue(lines, "Remote collection URL");
  const taskPageTypes = readProfileEditorValue(lines, "Task page types");
  const taskTags = readProfileEditorValue(lines, "Task tags");
  const ignoredTags = readProfileEditorValue(lines, "Ignored tags");
  const prefilterPagesOnly = /^yes$/i.test(readProfileEditorValue(lines, "Only scan matched pages") || "no");
  const propertyKey = readProfileEditorValue(lines, "Calendar property key override") || "calendar";
  const propertyValue = readProfileEditorValue(lines, "Calendar property value override");
  const writeToJournal = /^yes$/i.test(readProfileEditorValue(lines, "Write to journal") || "no");
  const defaultImportPage = readProfileEditorValue(lines, "Default import page");
  const simpleChecklistMode = /^yes$/i.test(readProfileEditorValue(lines, "Simple checklist mode") || "no");
  const updatePeriodically = /^yes$/i.test(readProfileEditorValue(lines, "Update periodically") || "no");
  const updateIntervalMinutes = Math.max(1, Number(readProfileEditorValue(lines, "Update interval minutes") || 15) || 15);
  const enabled = !/^no$/i.test(readProfileEditorValue(lines, "Enabled") || "yes");

  const updated = profiles.map((item) =>
    item.id === scope.id
      ? {
          ...item,
          name,
          remoteUrl,
          taskListUrl: remoteUrl || item.taskListUrl,
          filterPageTypes: taskPageTypes,
          filterTags: taskTags,
          ignoredTags,
          prefilterPagesOnly,
          calendarUrl: remoteUrl || item.calendarUrl,
          propertyKey,
          propertyValue,
          writeToJournal,
          defaultImportPage,
          simpleChecklistMode,
          updatePeriodically,
          updateIntervalMinutes,
          enabled
        }
      : item
  );

  await saveSyncProfiles(updated, scope.id);
  syncTaskScopeManagerUI();
  syncCalendarSelectorUI();
  syncHubUI();
  openPage(taskScopeManagerPage);
  logseq.UI.showMsg?.(`Updated sync profile ${name}.`, "success", { timeout: 5000 });
}

async function refreshTaskScopePreview(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
  const scope = getTaskScopes(settings()).find((item) => item.id === scopeId) ?? getActiveTaskScope(settings());
  if (!scope) {
    logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
    return;
  }
  await refreshTasks(scope);
  openPage(settings().taskPageName || defaultSettings.taskPageName);
}

async function syncTaskScope(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
  const profile = getSyncProfiles(settings()).find((item) => item.id === scopeId) ?? getActiveSyncProfile(settings());
  if (!profile) {
    logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
    return;
  }
  await syncProfileEverything(profile);
}

async function createRemoteTaskListForActiveScope() {
  const active = getActiveSyncProfile(settings());
  if (!active) {
    logseq.UI.showMsg?.("Create or select a sync profile first.", "warning", { timeout: 5000 });
    return;
  }
  if (active.id === "legacy-default") {
    logseq.UI.showMsg?.("Create a named sync profile first so the new remote collection has somewhere to be saved.", "warning", { timeout: 6000 });
    return;
  }

  const displayName = promptValue("Nextcloud remote collection name", active.name);
  if (!displayName) return;

  try {
    const created = await createCalDavTaskList(settings(), displayName);
    const profiles = getPersistedSyncProfiles(settings()).map((profile) =>
      profile.id === active.id ? { ...profile, remoteUrl: created.url, taskListUrl: created.url, calendarUrl: created.url } : profile
    );
  await saveSyncProfiles(profiles, active.id);
  state.taskListDiscovery = await discoverCalDavTaskLists(settings());
  syncTaskScopeManagerUI();
  syncHubUI();
  logseq.UI.showMsg?.(`Created and assigned remote collection ${displayName}.`, "success", { timeout: 7000 });
  } catch (error) {
    console.error("[nextcloud-sync] remote collection creation failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Remote collection creation failed.", "error", { timeout: 7000 });
  }
}

async function setTaskCollectionUrl() {
  const active = getActiveTaskScope(settings());
  const current = active?.taskListUrl || settings().caldavTaskListUrl || "https://your-host/remote.php/dav/calendars/username/task-list/";
  const manual = promptValue(
    "Paste the exact Nextcloud remote collection URL.\n\nExample:\nhttps://host/remote.php/dav/calendars/username/personal/",
    current
  );
  if (!manual) {
    logseq.UI.showMsg?.("No remote collection URL saved.", "warning", { timeout: 4000 });
    return;
  }
  const scopes = getPersistedTaskScopes(settings());
  if (active && active.id !== "legacy-default") {
    await saveTaskScopes(
      scopes.map((scope) => (scope.id === active.id ? { ...scope, taskListUrl: manual, remoteUrl: manual, calendarUrl: manual } : scope)),
      active.id
    );
    syncTaskScopeManagerUI();
    syncHubUI();
    logseq.UI.showMsg?.(`Saved remote collection URL for ${active.name}.`, "success", { timeout: 5000 });
    return;
  }
  logseq.updateSettings?.({ caldavTaskListUrl: manual, caldavCalendarUrl: manual });
  logseq.UI.showMsg?.("Saved remote collection URL.", "success", { timeout: 5000 });
}

async function setTaskCollectionFromClipboard() {
  try {
    const text = await navigator.clipboard.readText();
    const manual = String(text ?? "").trim();
    if (!manual) {
      logseq.UI.showMsg?.("Clipboard is empty. Copy the remote collection URL first.", "warning", { timeout: 4000 });
      return;
    }
    const active = getActiveTaskScope(settings());
    const scopes = getPersistedTaskScopes(settings());
    if (active && active.id !== "legacy-default") {
      await saveTaskScopes(
        scopes.map((scope) => (scope.id === active.id ? { ...scope, taskListUrl: manual, remoteUrl: manual, calendarUrl: manual } : scope)),
        active.id
      );
      syncTaskScopeManagerUI();
      syncHubUI();
      logseq.UI.showMsg?.(`Saved remote collection URL from clipboard for ${active.name}.`, "success", { timeout: 5000 });
      return;
    }
    logseq.updateSettings?.({ caldavTaskListUrl: manual, caldavCalendarUrl: manual });
    logseq.UI.showMsg?.("Saved remote collection URL from clipboard.", "success", { timeout: 5000 });
  } catch (error) {
    console.error("[nextcloud-sync] clipboard import failed", error);
    logseq.UI.showMsg?.("Couldn't read the clipboard. Use the manual URL command instead.", "warning", { timeout: 5000 });
  }
}

async function setCalendarCollectionUrl() {
  const active = getActiveCalendarScope(settings());
  const current =
    active?.calendarUrl || settings().caldavCalendarUrl || "https://your-host/remote.php/dav/calendars/username/calendar/";
  const manual = promptValue(
    "Paste the exact Nextcloud remote collection URL.\n\nExample:\nhttps://host/remote.php/dav/calendars/username/personal/",
    current
  );
  if (!manual) {
    logseq.UI.showMsg?.("No remote collection URL saved.", "warning", { timeout: 4000 });
    return;
  }
  const scopes = getPersistedCalendarScopes(settings());
  if (active && active.id !== "legacy-default") {
    await saveCalendarScopes(
      scopes.map((scope) => (scope.id === active.id ? { ...scope, calendarUrl: manual, remoteUrl: manual, taskListUrl: manual } : scope)),
      active.id
    );
      syncCalendarSelectorUI();
      syncHubUI();
      logseq.UI.showMsg?.(`Saved remote collection URL for ${active.name}.`, "success", { timeout: 5000 });
    return;
  }
  logseq.updateSettings?.({ caldavCalendarUrl: manual, caldavTaskListUrl: manual });
  syncCalendarSelectorUI();
  logseq.UI.showMsg?.("Saved remote collection URL.", "success", { timeout: 5000 });
}

async function setCalendarCollectionFromClipboard() {
  try {
    const text = await navigator.clipboard.readText();
    const manual = String(text ?? "").trim();
    if (!manual) {
      logseq.UI.showMsg?.("Clipboard is empty. Copy the remote collection URL first.", "warning", { timeout: 4000 });
      return;
    }
    const active = getActiveCalendarScope(settings());
    const scopes = getPersistedCalendarScopes(settings());
    if (active && active.id !== "legacy-default") {
      await saveCalendarScopes(
        scopes.map((scope) => (scope.id === active.id ? { ...scope, calendarUrl: manual, remoteUrl: manual, taskListUrl: manual } : scope)),
        active.id
      );
      syncCalendarSelectorUI();
      syncHubUI();
      logseq.UI.showMsg?.(`Saved remote collection URL from clipboard for ${active.name}.`, "success", { timeout: 5000 });
      return;
    }
    logseq.updateSettings?.({ caldavCalendarUrl: manual, caldavTaskListUrl: manual });
    syncCalendarSelectorUI();
    logseq.UI.showMsg?.("Saved remote collection URL from clipboard.", "success", { timeout: 5000 });
  } catch (error) {
    console.error("[nextcloud-sync] calendar clipboard import failed", error);
    logseq.UI.showMsg?.("Couldn't read the clipboard. Use the manual URL command instead.", "warning", { timeout: 5000 });
  }
}

async function setDavRootUrl() {
  const current = settings().nextcloudDavUrl || "https://your-host/remote.php/dav";
  const manual = promptValue("Paste the Nextcloud DAV root URL.\n\nExample:\nhttps://host/remote.php/dav", current);
  if (!manual) {
    logseq.UI.showMsg?.("No DAV root URL saved.", "warning", { timeout: 4000 });
    return;
  }
  logseq.updateSettings?.({ nextcloudDavUrl: manual });
  logseq.UI.showMsg?.("Saved Nextcloud DAV root URL.", "success", { timeout: 5000 });
}

async function syncOnStartup() {
  const current = settings();
  if (!current.syncOnStartup) return;
  const persistedScopes = getPersistedTaskScopes(current).filter((scope) => scope.enabled && scope.taskListUrl);
  if (persistedScopes.length && current.caldavUsername && current.caldavPassword) {
    for (const scope of persistedScopes) {
      await syncTasks(scope);
    }
  } else if (current.caldavTaskListUrl && current.caldavUsername && current.caldavPassword) {
    await syncTasks(getActiveTaskScope(current) ?? undefined);
  }
  const persistedCalendarScopes = getPersistedCalendarScopes(current).filter((scope) => scope.enabled && scope.calendarUrl);
  if (persistedCalendarScopes.length && current.caldavUsername && current.caldavPassword) {
    for (const scope of persistedCalendarScopes) {
      await syncCalendar(scope);
    }
  } else if (current.caldavCalendarUrl && current.caldavUsername && current.caldavPassword) {
    await syncCalendar(getActiveCalendarScope(current) ?? undefined);
  }
}

logseq.ready(async () => {
  try {
    if (runtimeState.initialized) {
      console.warn("[nextcloud-sync] skipped duplicate startup");
      return;
    }
    runtimeState.initialized = true;

    configureSettings();
    providePluginStyle();
    mountInlineUi();

    logseq.provideModel?.({
      refreshCalendarDiscovery: async () => {
        await discoverCalendars();
      },
      createCalendarScope: async () => {
        await createCalendarScope();
      },
      activateCalendarScope: async (e: any) => {
        await activateCalendarScope(e);
      },
      editCalendarScope: async (e: any) => {
        await editCalendarScope(e);
      },
      refreshCalendarScopePreview: async (e: any) => {
        await refreshCalendarScopePreview(e);
      },
      syncCalendarScope: async (e: any) => {
        await syncCalendarScope(e);
      },
      syncAllCalendarScopes: async () => {
        await syncAllCalendarScopes();
      },
      selectCalendarOption: async (e: any) => {
        await selectCalendarOption(e);
      },
      discoverTaskLists: async () => {
        await discoverTaskLists();
      },
      createTaskScope: async () => {
        await createTaskScope();
      },
      activateTaskScope: async (e: any) => {
        await activateTaskScope(e);
      },
      editTaskScope: async (e: any) => {
        await editTaskScope(e);
      },
      refreshTaskScopePreview: async (e: any) => {
        await refreshTaskScopePreview(e);
      },
      syncTaskScope: async (e: any) => {
        await syncTaskScope(e);
      },
      assignTaskListToActiveScope: async (e: any) => {
        await assignTaskListToActiveScope(e);
      },
      createRemoteTaskListForActiveScope: async () => {
        await createRemoteTaskListForActiveScope();
      },
      syncAllTaskScopes: async () => {
        await syncAllTaskScopes();
      },
      importRemoteProfileItems: async (e: any) => {
        const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
        const scope = getSyncProfiles(settings()).find((item) => item.id === scopeId) ?? getActiveSyncProfile(settings());
        if (!scope) {
          logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
          return;
        }
        await importRemoteItems(scope);
      },
      cleanupImportedProfileItems: async (e: any) => {
        const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
        const scope = getSyncProfiles(settings()).find((item) => item.id === scopeId) ?? getActiveSyncProfile(settings());
        if (!scope) {
          logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
          return;
        }
        await cleanupImportedItems(scope);
      },
      dedupeImportedProfileItems: async (e: any) => {
        const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
        const scope = getSyncProfiles(settings()).find((item) => item.id === scopeId) ?? getActiveSyncProfile(settings());
        if (!scope) {
          logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
          return;
        }
        await dedupeImportedItems(scope);
      },
      openProfileInboxPage: async (e: any) => {
        const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
        const scope = getSyncProfiles(settings()).find((item) => item.id === scopeId) ?? getActiveSyncProfile(settings());
        if (!scope) {
          logseq.UI.showMsg?.("Could not find that sync profile.", "warning", { timeout: 5000 });
          return;
        }
        await openInboxPage(scope);
      },
      importAllRemoteItems: async () => {
        await importRemoteItemsForAllProfiles();
      },
      cleanupAllImportedItems: async () => {
        await cleanupImportedItemsForAllProfiles();
      },
      dedupeAllImportedItems: async () => {
        await dedupeImportedItemsForAllProfiles();
      },
      saveProfileEditorFromUi: async () => {
        await saveProfileEditor();
      },
      openSyncProfilesPage: async () => {
        await ensureTaskScopeManagerPage();
        syncTaskScopeManagerUI();
        openPage(taskScopeManagerPage);
      },
      openTaskOverviewPage: async () => {
        openPage(settings().taskPageName || defaultSettings.taskPageName);
      },
      openInboxOverviewPage: async () => {
        await openInboxPage();
      },
      openCalendarOverviewPage: async () => {
        openPage(settings().calendarPageName || defaultSettings.calendarPageName);
      },
      openCalendarPickerPage: async () => {
        await ensureCalendarSelectorPage();
        syncCalendarSelectorUI();
        openPage(calendarSelectorPage);
      }
    });

    registerCommand("Nextcloud: Open sync", async () => {
      await ensureSyncHubPage();
      syncHubUI();
      openPage(syncHubPage);
    });
    registerCommand("Nextcloud: Open tasks", async () => {
      openPage(settings().taskPageName || defaultSettings.taskPageName);
    });
    registerCommand("Nextcloud: Open inbox", async () => {
      await openInboxPage();
    });
    registerCommand("Nextcloud: Open profiles", async () => {
      await ensureTaskScopeManagerPage();
      syncTaskScopeManagerUI();
      openPage(taskScopeManagerPage);
    });
    registerCommand("Nextcloud: Open sync profiles", async () => {
      await ensureTaskScopeManagerPage();
      syncTaskScopeManagerUI();
      openPage(taskScopeManagerPage);
    });
    registerCommand("Nextcloud: Open calendar", async () => {
      openPage(settings().calendarPageName || defaultSettings.calendarPageName);
    });
    registerCommand("Nextcloud: Open calendars", async () => {
      await ensureCalendarSelectorPage();
      syncCalendarSelectorUI();
      openPage(calendarSelectorPage);
    });
    registerCommand("Nextcloud: Create profile", async () => {
      await createCalendarScope();
    });
    registerCommand("Nextcloud: Import remote items", async () => {
      await importRemoteItems();
    });
    registerCommand("Nextcloud: Import all remote items", async () => {
      await importRemoteItemsForAllProfiles();
    });
    registerCommand("Nextcloud: Cleanup imported items", async () => {
      await cleanupImportedItems();
    });
    registerCommand("Nextcloud: Cleanup all imported items", async () => {
      await cleanupImportedItemsForAllProfiles();
    });
    registerCommand("Nextcloud: Dedupe imported items", async () => {
      await dedupeImportedItems();
    });
    registerCommand("Nextcloud: Dedupe all imported items", async () => {
      await dedupeImportedItemsForAllProfiles();
    });
    registerCommand("Nextcloud: Discover calendars", async () => {
      await discoverCalendars();
    });
    registerCommand("Nextcloud: Discover collections", async () => {
      await discoverTaskLists();
    });
    registerCommand("Nextcloud: Create profile (legacy)", async () => {
      await createTaskScope();
    });
    registerCommand("Nextcloud: Create sync profile", async () => {
      await createTaskScope();
    });
    registerCommand("Nextcloud: Create remote collection", async () => {
      await createRemoteTaskListForActiveScope();
    });
    registerCommand("Nextcloud: Save profile editor", async () => {
      await saveProfileEditor();
    });
    registerCommand("Nextcloud: Refresh tasks", async () => {
      await refreshTasks();
    });
    registerCommand("Nextcloud: Refresh calendar", async () => {
      await refreshCalendar();
    });
    registerCommand("Nextcloud: Export tasks", async () => {
      await exportTasks();
    });
    registerCommand("Nextcloud: Export calendar", async () => {
      await exportCalendar();
    });
    registerCommand("Nextcloud: Sync tasks", async () => {
      await syncTasks();
    });
    registerCommand("Nextcloud: Sync all", async () => {
      await syncAllTaskScopes();
    });
    registerCommand("Nextcloud: Sync all profiles", async () => {
      await syncAllTaskScopes();
    });
    registerCommand("Nextcloud: Sync calendar", async () => {
      await syncCalendar();
    });
    registerCommand("Nextcloud: Sync calendars", async () => {
      await syncAllCalendarScopes();
    });
    registerCommand("Nextcloud: Test tasks connection", async () => {
      await testConnection();
    });
    registerCommand("Nextcloud: Test calendar connection", async () => {
      await testCalendarCollection();
    });
    registerCommand("Nextcloud: Set remote URL", async () => {
      await setTaskCollectionUrl();
    });
    registerCommand("Nextcloud: Set DAV root", async () => {
      await setDavRootUrl();
    });
    registerCommand("Nextcloud: Set remote URL (calendar)", async () => {
      await setCalendarCollectionUrl();
    });
    registerCommand("Nextcloud: Use clipboard URL", async () => {
      await setTaskCollectionFromClipboard();
    });
    registerCommand("Nextcloud: Use clipboard URL (calendar)", async () => {
      await setCalendarCollectionFromClipboard();
    });

    registerSlashCommand("Nextcloud tasks refresh", async () => {
      await refreshTasks();
    });
    registerSlashCommand("Nextcloud tasks export", async () => {
      await exportTasks();
    });
    registerSlashCommand("Nextcloud tasks sync", async () => {
      await syncTasks();
    });
    registerSlashCommand("Nextcloud sync", async () => {
      const activeProfile = getActiveSyncProfile(settings());
      if (!activeProfile) {
        logseq.UI.showMsg?.("Create or select a sync profile first.", "warning", { timeout: 5000 });
        return;
      }

      await syncTaskScope({ dataset: { scopeId: activeProfile.id } });
    });
    registerSlashCommand("Nextcloud import", async () => {
      await importRemoteItems();
    });
    registerSlashCommand("Nextcloud cleanup", async () => {
      await cleanupImportedItems();
    });
    registerSlashCommand("Nextcloud dedupe", async () => {
      await dedupeImportedItems();
    });
    registerSlashCommand("Nextcloud sync all", async () => {
      await syncAllTaskScopes();
    });
    registerSlashCommand("Nextcloud profiles", async () => {
      await ensureTaskScopeManagerPage();
      syncTaskScopeManagerUI();
      openPage(taskScopeManagerPage);
    });
    registerSlashCommand("Nextcloud inbox", async () => {
      await openInboxPage();
    });
    registerSlashCommand("Nextcloud open", async () => {
      await ensureSyncHubPage();
      syncHubUI();
      openPage(syncHubPage);
    });
    registerSlashCommand("Nextcloud calendar refresh", async () => {
      await refreshCalendar();
    });
    registerSlashCommand("Nextcloud calendar export", async () => {
      await exportCalendar();
    });
    registerSlashCommand("Nextcloud calendar sync", async () => {
      await syncCalendar();
    });

    scheduleStartupSync();
    refreshProfileAutoSyncSchedules();

    logseq.UI.showMsg?.("Logseq Nextcloud Sync loaded.", "success");
  } catch (error) {
    runtimeState.initialized = false;
    console.error("[nextcloud-sync] startup failed", error);
    logseq.UI.showMsg?.("Logseq Nextcloud Sync failed to start. Check the dev console.", "error");
  }
});
