import "@logseq/libs";
import {
  collectLogseqTasksForScope,
  collectLogseqTasksWithSettings,
  createCalDavTaskList,
  discoverCalDavTaskLists,
  exportTaskListIcs,
  syncTasksToCalDav,
  testCalDavConnection,
  testCalDavConnectionForUrl
} from "./caldav";
import { collectLogseqCalendarEvents, discoverCalDavCalendars, exportCalendarIcs, syncCalendarToCalDav, testCalendarConnection } from "./calendar";
import type { CalendarDiscoveryResult } from "./calendar";
import type { DiscoveredTaskList, LogseqCalendarEvent, LogseqTaskItem, NextcloudSyncSettings, TaskScopeConfig } from "./types";

declare const logseq: any;

const defaultSettings: NextcloudSyncSettings = {
  syncOnStartup: true,
  calendarTimezone: "Europe/Stockholm",
  nextcloudDavUrl: "",
  caldavTaskListUrl: "",
  caldavCalendarUrl: "",
  caldavUsername: "",
  caldavPassword: "",
  taskPageName: "Nextcloud Tasks",
  calendarPageName: "Nextcloud Calendar",
  taskFilterPageTypes: "",
  taskFilterTags: "",
  taskScopesJson: "",
  activeTaskScopeId: ""
};

const state = {
  tasks: [] as LogseqTaskItem[],
  events: [] as LogseqCalendarEvent[],
  calendarDiscovery: null as CalendarDiscoveryResult | null,
  taskListDiscovery: null as TaskListDiscoveryResult | null
};

const uiState = {
  calendarSelectorSlot: "",
  taskScopeManagerSlot: ""
};

const calendarSelectorPage = "Nextcloud Calendar Picker";
const taskScopeManagerPage = "Nextcloud Tasklist Scopes";

type TaskListDiscoveryResult = Awaited<ReturnType<typeof discoverCalDavTaskLists>>;

function settings(): NextcloudSyncSettings {
  return { ...defaultSettings, ...(logseq?.settings ?? {}) };
}

function registerCommand(label: string, handler: () => Promise<void> | void) {
  if (typeof logseq.App?.registerCommandPalette === "function") {
    logseq.App.registerCommandPalette({ key: label, label }, handler);
  }
}

function registerSlashCommand(label: string, handler: () => Promise<void> | void) {
  const editor = logseq.Editor as any;
  if (typeof editor.registerSlashCommand === "function") {
    editor.registerSlashCommand(label, handler);
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
    enabled: true
  };
}

function readStoredTaskScopes(current: NextcloudSyncSettings) {
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
        enabled: scope?.enabled !== false
      }))
      .filter((scope) => scope.id && scope.name);
  } catch {
    return [];
  }
}

function getTaskScopes(current = settings()) {
  const stored = readStoredTaskScopes(current);
  return stored.length ? stored : [buildLegacyTaskScope(current)];
}

function getPersistedTaskScopes(current = settings()) {
  return readStoredTaskScopes(current);
}

function getActiveTaskScope(current = settings()) {
  const scopes = getTaskScopes(current);
  if (!scopes.length) return null;
  const activeId = String(current.activeTaskScopeId || "").trim();
  return scopes.find((scope) => scope.id === activeId) ?? scopes[0];
}

function withScopeSettings(current: NextcloudSyncSettings, scope: TaskScopeConfig): NextcloudSyncSettings {
  return {
    ...current,
    caldavTaskListUrl: scope.taskListUrl,
    taskFilterPageTypes: scope.filterPageTypes,
    taskFilterTags: scope.filterTags
  };
}

async function saveTaskScopes(scopes: TaskScopeConfig[], activeScopeId?: string) {
  await logseq.updateSettings?.({
    taskScopesJson: JSON.stringify(scopes),
    activeTaskScopeId: activeScopeId ?? scopes[0]?.id ?? ""
  });
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

function renderTaskLine(task: LogseqTaskItem) {
  const pageRef = `[[${task.pageName}]]`;
  const blockRef = task.sourceBlockUuid ? `((${task.sourceBlockUuid}))` : "";
  const state = task.taskState ? `[${task.taskState}]` : "";
  const due = formatTaskDate(task);
  const scopeParts = [
    task.pageType ? `page-type:${task.pageType}` : "",
    task.pageTags.length ? `tags:${task.pageTags.join(",")}` : ""
  ].filter(Boolean);
  const scope = scopeParts.length ? ` {${scopeParts.join(" | ")}}` : "";
  return `${pageRef} :: ${task.title} ${state} {${due}}${scope} ${blockRef}`.replace(/\s+/g, " ").trim();
}

function renderEventLine(event: LogseqCalendarEvent) {
  const pageRef = `[[${event.pageName}]]`;
  const blockRef = event.sourceBlockUuid ? `((${event.sourceBlockUuid}))` : "";
  return `${pageRef} :: ${event.title} {${event.kind} · ${formatEventDate(event)}} ${blockRef}`.replace(/\s+/g, " ").trim();
}

async function writeTaskOverviewPage(tasks: LogseqTaskItem[], scope = getActiveTaskScope(settings())) {
  const pageName = settings().taskPageName || defaultSettings.taskPageName;
  const editor = logseq.Editor as any;
  const timestamp = new Date().toISOString();
  const scopeSummary = describeTaskScope(scope);
  const lines = [
    `- Nextcloud task sync snapshot`,
    `  - Updated: ${timestamp}`,
    `  - Active scope: ${scope?.name ?? "None"}`,
    `  - Tasks found: ${tasks.length}`,
    `  - Scope: ${scopeSummary}`,
    `  - Task list: ${scope?.taskListUrl || "Not selected"}`,
    ...tasks.map((task) => `  - ${renderTaskLine(task)}`)
  ];

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
    await editor.appendBlockInPage?.(pageName, lines.join("\n"));
  } catch (error) {
    console.warn("[nextcloud-sync] could not update task overview page", error);
  }
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

  if (!pageTypes.length && !tags.length) return "Entire graph";
  const parts = [
    pageTypes.length ? `page types: ${pageTypes.join(", ")}` : "",
    tags.length ? `tags: ${tags.join(", ")}` : ""
  ].filter(Boolean);
  return parts.join(" OR ");
}

async function writeCalendarOverviewPage(events: LogseqCalendarEvent[]) {
  const pageName = settings().calendarPageName || defaultSettings.calendarPageName;
  const editor = logseq.Editor as any;
  const timestamp = new Date().toISOString();
  const lines = [
    `- Nextcloud calendar sync snapshot`,
    `  - Updated: ${timestamp}`,
    `  - Events found: ${events.length}`,
    ...events.map((event) => `  - ${renderEventLine(event)}`)
  ];

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
    await editor.appendBlockInPage?.(pageName, lines.join("\n"));
  } catch (error) {
    console.warn("[nextcloud-sync] could not update calendar overview page", error);
  }
}

async function replacePageWithContent(pageName: string, content: string) {
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
    await editor.appendBlockInPage?.(pageName, content);
  } catch (error) {
    console.warn("[nextcloud-sync] could not replace page content", error);
  }
}

function calendarSelectorTemplate() {
  const discovery = state.calendarDiscovery;
  const items = discovery?.calendars ?? [];
  const status = discovery
    ? `<div class="nextcloud-selector__status">${escapeHtml(discovery.message)}</div>`
    : `<div class="nextcloud-selector__status">Run discovery to load calendars from your Nextcloud account.</div>`;

  const cards = items.length
    ? items
        .map(
          (calendar, index) => `
            <div class="nextcloud-selector__card">
              <div class="nextcloud-selector__title">${escapeHtml(calendar.displayName || `Calendar ${index + 1}`)}</div>
              <div class="nextcloud-selector__meta">${escapeHtml(calendar.url)}</div>
              <div class="nextcloud-selector__meta">${escapeHtml(calendar.componentSet.join(", ") || "VEVENT")}</div>
              <div class="nextcloud-selector__actions">
                <button data-on-click="selectCalendarOption" data-url="${escapeHtml(calendar.url)}">Use This Calendar</button>
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
          <div class="nextcloud-selector__hint">Discover VEVENT calendar collections and save one directly into plugin settings.</div>
        </div>
        <div class="nextcloud-selector__actions">
          <button data-on-click="refreshCalendarDiscovery">Discover Calendars</button>
        </div>
      </div>
      ${status}
      <div class="nextcloud-selector__list">${cards}</div>
    </div>
  `;
}

function syncCalendarSelectorUI() {
  if (!uiState.calendarSelectorSlot || typeof logseq.provideUI !== "function") return;
  logseq.provideUI({
    key: `nextcloud-calendar-selector-${uiState.calendarSelectorSlot}`,
    slot: uiState.calendarSelectorSlot,
    reset: true,
    template: calendarSelectorTemplate()
  });
}

function taskScopeManagerTemplate() {
  const current = settings();
  const scopes = getTaskScopes(current);
  const activeScope = getActiveTaskScope(current);
  const discovered = state.taskListDiscovery?.taskLists ?? [];

  const scopeCards = scopes
    .map(
      (scope) => `
        <div class="nextcloud-selector__card">
          <div class="nextcloud-selector__title">
            ${escapeHtml(scope.name)}${scope.id === activeScope?.id ? " (active)" : ""}
          </div>
          <div class="nextcloud-selector__meta">Scope: ${escapeHtml(describeTaskScope(scope))}</div>
          <div class="nextcloud-selector__meta">Task list: ${escapeHtml(scope.taskListUrl || "Not selected")}</div>
          <div class="nextcloud-selector__actions">
            <button data-on-click="activateTaskScope" data-scope-id="${escapeHtml(scope.id)}">Activate</button>
            <button data-on-click="refreshTaskScopePreview" data-scope-id="${escapeHtml(scope.id)}">Preview</button>
            <button data-on-click="syncTaskScope" data-scope-id="${escapeHtml(scope.id)}">Sync</button>
            <button data-on-click="editTaskScope" data-scope-id="${escapeHtml(scope.id)}">Edit</button>
          </div>
        </div>
      `
    )
    .join("");

  const discoveredCards = discovered.length
    ? discovered
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
    : `<div class="nextcloud-selector__empty">No task lists discovered yet.</div>`;

  const status = state.taskListDiscovery
    ? `<div class="nextcloud-selector__status">${escapeHtml(state.taskListDiscovery.message)}</div>`
    : `<div class="nextcloud-selector__status">Create scopes, then discover or create Nextcloud task lists for them.</div>`;

  return `
    <div class="nextcloud-selector">
      <div class="nextcloud-selector__header">
        <div>
          <div class="nextcloud-selector__headline">Nextcloud Tasklist Scopes</div>
          <div class="nextcloud-selector__hint">Each scope can target different page filters and a different Nextcloud task list.</div>
        </div>
        <div class="nextcloud-selector__actions">
          <button data-on-click="createTaskScope">Create Scope</button>
          <button data-on-click="discoverTaskLists">Discover Task Lists</button>
          <button data-on-click="createRemoteTaskListForActiveScope">Create Remote Task List</button>
          <button data-on-click="syncAllTaskScopes">Sync All Enabled Scopes</button>
        </div>
      </div>
      ${status}
      <div class="nextcloud-selector__section-title">Scopes</div>
      <div class="nextcloud-selector__list">${scopeCards}</div>
      <div class="nextcloud-selector__section-title">Discovered Task Lists</div>
      <div class="nextcloud-selector__list">${discoveredCards}</div>
    </div>
  `;
}

function syncTaskScopeManagerUI() {
  if (!uiState.taskScopeManagerSlot || typeof logseq.provideUI !== "function") return;
  logseq.provideUI({
    key: `nextcloud-task-scope-manager-${uiState.taskScopeManagerSlot}`,
    slot: uiState.taskScopeManagerSlot,
    reset: true,
    template: taskScopeManagerTemplate()
  });
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
  `);
}

function mountInlineUi() {
  if (typeof logseq.App?.onMacroRendererSlotted !== "function" || typeof logseq.provideUI !== "function") return;
  logseq.App.onMacroRendererSlotted(({ slot, payload }: any) => {
    const args = Array.isArray(payload?.arguments) ? payload.arguments : [];
    const macroName = args[0];
    if (macroName === ":nextcloud-calendar-selector") {
      uiState.calendarSelectorSlot = slot;
      syncCalendarSelectorUI();
      return;
    }
    if (macroName === ":nextcloud-task-scope-manager") {
      uiState.taskScopeManagerSlot = slot;
      syncTaskScopeManagerUI();
    }
  });
}

async function loadTasksSilently() {
  try {
    const activeScope = getActiveTaskScope(settings());
    state.tasks = activeScope ? await collectLogseqTasksForScope(activeScope, settings()) : await collectLogseqTasksWithSettings(settings());
    state.events = await collectLogseqCalendarEvents();
  } catch (error) {
    console.warn("[nextcloud-sync] silent refresh failed", error);
  }
}

async function ensureCalendarSelectorPage() {
  await replacePageWithContent(
    calendarSelectorPage,
    `- Nextcloud calendar selection\n  - Use the picker below to discover and choose a Nextcloud calendar.\n  - {{renderer :nextcloud-calendar-selector}}`
  );
}

async function ensureTaskScopeManagerPage() {
  await replacePageWithContent(
    taskScopeManagerPage,
    `- Nextcloud tasklist scopes\n  - Create multiple scoped task sync profiles and bind each one to a Nextcloud task list.\n  - {{renderer :nextcloud-task-scope-manager}}`
  );
}

async function refreshTasks(scope = getActiveTaskScope(settings())) {
  try {
    const current = settings();
    const tasks = scope ? await collectLogseqTasksForScope(scope, current) : await collectLogseqTasksWithSettings(current);
    state.tasks = tasks;
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

async function refreshCalendar() {
  try {
    const events = await collectLogseqCalendarEvents();
    state.events = events;
    await writeCalendarOverviewPage(events);
    logseq.UI.showMsg?.(`Indexed ${events.length} Logseq calendar events.`, events.length ? "success" : "warning", { timeout: 5000 });
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
  await ensureCalendarSelectorPage();
  syncCalendarSelectorUI();
  openPage(calendarSelectorPage);
  logseq.UI.showMsg?.(result.message, result.ok ? "success" : "warning", { timeout: 7000 });
  return result;
}

async function discoverTaskLists() {
  const result = await discoverCalDavTaskLists(settings());
  state.taskListDiscovery = result;
  await ensureTaskScopeManagerPage();
  syncTaskScopeManagerUI();
  openPage(taskScopeManagerPage);
  logseq.UI.showMsg?.(result.message, result.ok ? "success" : "warning", { timeout: 7000 });
  return result;
}

async function exportTasks(scope = getActiveTaskScope(settings())) {
  const tasks = state.tasks.length && scope?.id === getActiveTaskScope(settings())?.id ? state.tasks : await refreshTasks(scope);
  if (!tasks.length) {
    logseq.UI.showMsg?.("No tasks found to export.", "warning");
    return;
  }

  const { filename } = await exportTaskListIcs(tasks, settings().calendarTimezone);
  logseq.UI.showMsg?.(`Exported ${tasks.length} tasks to ${filename}${scope ? ` for ${scope.name}` : ""}`, "success", { timeout: 4000 });
}

async function exportCalendar() {
  const events = state.events.length ? state.events : await refreshCalendar();
  if (!events.length) {
    logseq.UI.showMsg?.("No calendar events found to export.", "warning");
    return;
  }

  const { filename } = await exportCalendarIcs(events, settings().calendarTimezone);
  logseq.UI.showMsg?.(`Exported ${events.length} calendar events to ${filename}`, "success", { timeout: 4000 });
}

async function syncTasks(scope = getActiveTaskScope(settings())) {
  if (!scope) {
    logseq.UI.showMsg?.("Create or select a task scope first.", "warning", { timeout: 5000 });
    return;
  }
  if (!scope.taskListUrl) {
    logseq.UI.showMsg?.(`Select or create a Nextcloud task list for ${scope.name} first.`, "warning", { timeout: 6000 });
    return;
  }

  const tasks = state.tasks.length && scope.id === getActiveTaskScope(settings())?.id ? state.tasks : await refreshTasks(scope);
  if (!tasks.length) {
    logseq.UI.showMsg?.("No tasks found to sync.", "warning");
    return;
  }

  try {
    const result = await syncTasksToCalDav(tasks, withScopeSettings(settings(), scope));
    const summary = `Synced ${result.synced}, verified ${result.verified}, deleted ${result.deleted}, mirrored ${result.completedRemote}, failed ${result.failed}`;
    state.tasks = result.tasks;
    await writeTaskOverviewPage(result.tasks, scope);
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

async function syncAllTaskScopes() {
  const scopes = getTaskScopes(settings()).filter((scope) => scope.enabled);
  if (!scopes.length) {
    logseq.UI.showMsg?.("No enabled task scopes found.", "warning", { timeout: 5000 });
    return;
  }

  for (const scope of scopes) {
    await syncTasks(scope);
  }
  syncTaskScopeManagerUI();
}

async function syncCalendar() {
  const events = state.events.length ? state.events : await refreshCalendar();
  if (!events.length) {
    logseq.UI.showMsg?.("No calendar events found to sync.", "warning");
    return;
  }

  try {
    const result = await syncCalendarToCalDav(events, settings());
    const summary = `Synced ${result.synced}, verified ${result.verified}, deleted ${result.deleted}, failed ${result.failed}`;
    await writeCalendarOverviewPage(result.events);
    if (result.errors.length) {
      logseq.UI.showMsg?.(`${summary}. ${result.errors.join(" | ")}`, "warning", { timeout: 10000 });
    } else {
      logseq.UI.showMsg?.(`${summary}. Target: ${result.calendarUrl}`, "success", { timeout: 10000 });
    }
  } catch (error) {
    console.error("[nextcloud-sync] calendar sync failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Calendar sync failed.", "error", { timeout: 6000 });
  }
}

async function testConnection(scope = getActiveTaskScope(settings())) {
  if (!scope) {
    logseq.UI.showMsg?.("Create or select a task scope first.", "warning", { timeout: 5000 });
    return;
  }
  try {
    const result = await testCalDavConnectionForUrl(scope.taskListUrl, settings());
    if (result.ok) {
      logseq.UI.showMsg?.(`Task list reachable: ${result.url}`, "success", { timeout: 4000 });
    } else {
      logseq.UI.showMsg?.(result.message, "warning", { timeout: 6000 });
    }
  } catch (error) {
    console.error("[nextcloud-sync] connection test failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Connection test failed.", "error", { timeout: 6000 });
  }
}

async function testCalendarCollection() {
  try {
    const result = await testCalendarConnection(settings());
    if (result.ok) {
      logseq.UI.showMsg?.(`Calendar reachable: ${result.url}`, "success", { timeout: 4000 });
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
    logseq.UI.showMsg?.("No calendar URL was provided by the selector.", "warning", { timeout: 5000 });
    return;
  }
  logseq.updateSettings?.({ caldavCalendarUrl: selected });
  logseq.UI.showMsg?.(`Saved calendar URL: ${selected}`, "success", { timeout: 6000 });
}

function readDatasetValue(e: any, key: string) {
  return (
    e?.dataset?.[key] ??
    e?.currentTarget?.dataset?.[key] ??
    e?.target?.dataset?.[key] ??
    ""
  );
}

async function activateTaskScope(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
  const scopes = getPersistedTaskScopes(settings());
  const target = scopes.find((scope) => scope.id === scopeId);
  if (!target) {
    logseq.UI.showMsg?.("Could not find that task scope.", "warning", { timeout: 5000 });
    return;
  }
  await saveTaskScopes(scopes, target.id);
  syncTaskScopeManagerUI();
  await refreshTasks(target);
  logseq.UI.showMsg?.(`Active task scope: ${target.name}`, "success", { timeout: 4000 });
}

async function assignTaskListToActiveScope(e?: any) {
  const url = String(readDatasetValue(e, "url") || "").trim();
  if (!url) {
    logseq.UI.showMsg?.("No task list URL was provided by the selector.", "warning", { timeout: 5000 });
    return;
  }

  const current = settings();
  const scopes = getPersistedTaskScopes(current);
  const active = getActiveTaskScope(current);
  if (!active || active.id === "legacy-default") {
    logseq.UI.showMsg?.("Create a named task scope first, then assign a task list to it.", "warning", { timeout: 6000 });
    return;
  }

  const updated = scopes.map((scope) => (scope.id === active.id ? { ...scope, taskListUrl: url } : scope));
  await saveTaskScopes(updated, active.id);
  syncTaskScopeManagerUI();
  logseq.UI.showMsg?.(`Assigned task list to ${active.name}.`, "success", { timeout: 5000 });
}

async function createTaskScope() {
  const name = promptValue("Task scope name", "Work Tasks");
  if (!name) return;
  const pageTypes = promptValue("Comma-separated page-type values for this scope (optional)", "");
  const tags = promptValue("Comma-separated page tags for this scope (optional)", "");
  const scopes = getPersistedTaskScopes(settings());
  const scope: TaskScopeConfig = {
    id: createScopeId(name),
    name,
    taskListUrl: "",
    filterPageTypes: pageTypes,
    filterTags: tags,
    enabled: true
  };
  await saveTaskScopes([...scopes, scope], scope.id);
  await ensureTaskScopeManagerPage();
  syncTaskScopeManagerUI();
  openPage(taskScopeManagerPage);
  logseq.UI.showMsg?.(`Created task scope ${name}.`, "success", { timeout: 5000 });
}

async function editTaskScope(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim() || getActiveTaskScope(settings())?.id || "";
  const scopes = getPersistedTaskScopes(settings());
  const scope = scopes.find((item) => item.id === scopeId);
  if (!scope) {
    logseq.UI.showMsg?.("Could not find that task scope.", "warning", { timeout: 5000 });
    return;
  }

  const name = promptValue("Task scope name", scope.name) || scope.name;
  const pageTypes = promptValue("Comma-separated page-type values", scope.filterPageTypes);
  const tags = promptValue("Comma-separated page tags", scope.filterTags);
  const enabled = promptValue("Enabled? yes/no", scope.enabled ? "yes" : "no");
  const updated = scopes.map((item) =>
    item.id === scope.id
      ? {
          ...item,
          name,
          filterPageTypes: pageTypes,
          filterTags: tags,
          enabled: !/^no$/i.test(enabled.trim())
        }
      : item
  );
  await saveTaskScopes(updated, scope.id);
  syncTaskScopeManagerUI();
  await refreshTasks(updated.find((item) => item.id === scope.id));
  logseq.UI.showMsg?.(`Updated task scope ${name}.`, "success", { timeout: 5000 });
}

async function refreshTaskScopePreview(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
  const scope = getTaskScopes(settings()).find((item) => item.id === scopeId) ?? getActiveTaskScope(settings());
  if (!scope) {
    logseq.UI.showMsg?.("Could not find that task scope.", "warning", { timeout: 5000 });
    return;
  }
  await refreshTasks(scope);
  openPage(settings().taskPageName || defaultSettings.taskPageName);
}

async function syncTaskScope(e?: any) {
  const scopeId = String(readDatasetValue(e, "scopeId") || "").trim();
  const scope = getTaskScopes(settings()).find((item) => item.id === scopeId) ?? getActiveTaskScope(settings());
  if (!scope) {
    logseq.UI.showMsg?.("Could not find that task scope.", "warning", { timeout: 5000 });
    return;
  }
  await syncTasks(scope);
}

async function createRemoteTaskListForActiveScope() {
  const active = getActiveTaskScope(settings());
  if (!active) {
    logseq.UI.showMsg?.("Create or select a task scope first.", "warning", { timeout: 5000 });
    return;
  }
  if (active.id === "legacy-default") {
    logseq.UI.showMsg?.("Create a named task scope first so the new task list has somewhere to be saved.", "warning", { timeout: 6000 });
    return;
  }

  const displayName = promptValue("Nextcloud task list name", active.name);
  if (!displayName) return;

  try {
    const created = await createCalDavTaskList(settings(), displayName);
    const scopes = getPersistedTaskScopes(settings()).map((scope) =>
      scope.id === active.id ? { ...scope, taskListUrl: created.url } : scope
    );
    await saveTaskScopes(scopes, active.id);
    state.taskListDiscovery = await discoverCalDavTaskLists(settings());
    syncTaskScopeManagerUI();
    logseq.UI.showMsg?.(`Created and assigned task list ${displayName}.`, "success", { timeout: 7000 });
  } catch (error) {
    console.error("[nextcloud-sync] task list creation failed", error);
    logseq.UI.showMsg?.(error instanceof Error ? error.message : "Task list creation failed.", "error", { timeout: 7000 });
  }
}

async function setTaskCollectionUrl() {
  const active = getActiveTaskScope(settings());
  const current = active?.taskListUrl || settings().caldavTaskListUrl || "https://your-host/remote.php/dav/calendars/username/task-list/";
  const manual = promptValue(
    "Paste the exact Nextcloud task list collection URL.\n\nExample:\nhttps://host/remote.php/dav/calendars/username/task-list/",
    current
  );
  if (!manual) {
    logseq.UI.showMsg?.("No task list URL saved.", "warning", { timeout: 4000 });
    return;
  }
  const scopes = getPersistedTaskScopes(settings());
  if (active && active.id !== "legacy-default") {
    await saveTaskScopes(
      scopes.map((scope) => (scope.id === active.id ? { ...scope, taskListUrl: manual } : scope)),
      active.id
    );
    syncTaskScopeManagerUI();
    logseq.UI.showMsg?.(`Saved Nextcloud task list URL for ${active.name}.`, "success", { timeout: 5000 });
    return;
  }
  logseq.updateSettings?.({ caldavTaskListUrl: manual });
  logseq.UI.showMsg?.("Saved Nextcloud task list URL.", "success", { timeout: 5000 });
}

async function setTaskCollectionFromClipboard() {
  try {
    const text = await navigator.clipboard.readText();
    const manual = String(text ?? "").trim();
    if (!manual) {
      logseq.UI.showMsg?.("Clipboard is empty. Copy the task list URL first.", "warning", { timeout: 4000 });
      return;
    }
    const active = getActiveTaskScope(settings());
    const scopes = getPersistedTaskScopes(settings());
    if (active && active.id !== "legacy-default") {
      await saveTaskScopes(
        scopes.map((scope) => (scope.id === active.id ? { ...scope, taskListUrl: manual } : scope)),
        active.id
      );
      syncTaskScopeManagerUI();
      logseq.UI.showMsg?.(`Saved Nextcloud task list URL from clipboard for ${active.name}.`, "success", { timeout: 5000 });
      return;
    }
    logseq.updateSettings?.({ caldavTaskListUrl: manual });
    logseq.UI.showMsg?.("Saved Nextcloud task list URL from clipboard.", "success", { timeout: 5000 });
  } catch (error) {
    console.error("[nextcloud-sync] clipboard import failed", error);
    logseq.UI.showMsg?.("Couldn't read the clipboard. Use the manual URL command instead.", "warning", { timeout: 5000 });
  }
}

async function setCalendarCollectionUrl() {
  const current = settings().caldavCalendarUrl || "https://your-host/remote.php/dav/calendars/username/calendar/";
  const manual = promptValue(
    "Paste the exact Nextcloud calendar collection URL.\n\nExample:\nhttps://host/remote.php/dav/calendars/username/calendar/",
    current
  );
  if (!manual) {
    logseq.UI.showMsg?.("No calendar URL saved.", "warning", { timeout: 4000 });
    return;
  }
  logseq.updateSettings?.({ caldavCalendarUrl: manual });
  logseq.UI.showMsg?.("Saved Nextcloud calendar URL.", "success", { timeout: 5000 });
}

async function setCalendarCollectionFromClipboard() {
  try {
    const text = await navigator.clipboard.readText();
    const manual = String(text ?? "").trim();
    if (!manual) {
      logseq.UI.showMsg?.("Clipboard is empty. Copy the calendar URL first.", "warning", { timeout: 4000 });
      return;
    }
    logseq.updateSettings?.({ caldavCalendarUrl: manual });
    logseq.UI.showMsg?.("Saved Nextcloud calendar URL from clipboard.", "success", { timeout: 5000 });
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
  if (current.caldavCalendarUrl && current.caldavUsername && current.caldavPassword) {
    await syncCalendar();
  }
}

logseq.ready(async () => {
  try {
    configureSettings();
    providePluginStyle();
    mountInlineUi();
    await ensureCalendarSelectorPage();
    await ensureTaskScopeManagerPage();

    logseq.provideModel?.({
      refreshCalendarDiscovery: async () => {
        await discoverCalendars();
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
      }
    });

    registerCommand("Nextcloud Sync: Open task overview", async () => {
      openPage(settings().taskPageName || defaultSettings.taskPageName);
    });
    registerCommand("Nextcloud Sync: Open task scope manager", async () => {
      await ensureTaskScopeManagerPage();
      syncTaskScopeManagerUI();
      openPage(taskScopeManagerPage);
    });
    registerCommand("Nextcloud Sync: Open calendar overview", async () => {
      openPage(settings().calendarPageName || defaultSettings.calendarPageName);
    });
    registerCommand("Nextcloud Sync: Open calendar selector", async () => {
      await ensureCalendarSelectorPage();
      syncCalendarSelectorUI();
      openPage(calendarSelectorPage);
    });
    registerCommand("Nextcloud Sync: Discover calendars", async () => {
      await discoverCalendars();
    });
    registerCommand("Nextcloud Sync: Discover task lists", async () => {
      await discoverTaskLists();
    });
    registerCommand("Nextcloud Sync: Create task scope", async () => {
      await createTaskScope();
    });
    registerCommand("Nextcloud Sync: Create remote task list for active scope", async () => {
      await createRemoteTaskListForActiveScope();
    });
    registerCommand("Nextcloud Sync: Refresh task list", async () => {
      await refreshTasks();
    });
    registerCommand("Nextcloud Sync: Refresh calendar", async () => {
      await refreshCalendar();
    });
    registerCommand("Nextcloud Sync: Export task list ICS", async () => {
      await exportTasks();
    });
    registerCommand("Nextcloud Sync: Export calendar ICS", async () => {
      await exportCalendar();
    });
    registerCommand("Nextcloud Sync: Sync tasks to Nextcloud", async () => {
      await syncTasks();
    });
    registerCommand("Nextcloud Sync: Sync all task scopes to Nextcloud", async () => {
      await syncAllTaskScopes();
    });
    registerCommand("Nextcloud Sync: Sync calendar to Nextcloud", async () => {
      await syncCalendar();
    });
    registerCommand("Nextcloud Sync: Test task list connection", async () => {
      await testConnection();
    });
    registerCommand("Nextcloud Sync: Test calendar connection", async () => {
      await testCalendarCollection();
    });
    registerCommand("Nextcloud Sync: Set Nextcloud task list URL", async () => {
      await setTaskCollectionUrl();
    });
    registerCommand("Nextcloud Sync: Set Nextcloud DAV root URL", async () => {
      await setDavRootUrl();
    });
    registerCommand("Nextcloud Sync: Set Nextcloud calendar URL", async () => {
      await setCalendarCollectionUrl();
    });
    registerCommand("Nextcloud Sync: Use clipboard task list URL", async () => {
      await setTaskCollectionFromClipboard();
    });
    registerCommand("Nextcloud Sync: Use clipboard calendar URL", async () => {
      await setCalendarCollectionFromClipboard();
    });

    registerSlashCommand("Nextcloud task sync refresh", async () => {
      await refreshTasks();
    });
    registerSlashCommand("Nextcloud task sync export", async () => {
      await exportTasks();
    });
    registerSlashCommand("Nextcloud task scope manager", async () => {
      await ensureTaskScopeManagerPage();
      syncTaskScopeManagerUI();
      openPage(taskScopeManagerPage);
    });
    registerSlashCommand("Nextcloud calendar sync refresh", async () => {
      await refreshCalendar();
    });
    registerSlashCommand("Nextcloud calendar sync export", async () => {
      await exportCalendar();
    });

    await loadTasksSilently();
    await syncOnStartup();

    logseq.UI.showMsg?.("Logseq Nextcloud Sync loaded.", "success");
  } catch (error) {
    console.error("[nextcloud-sync] startup failed", error);
    logseq.UI.showMsg?.("Logseq Nextcloud Sync failed to start. Check the dev console.", "error");
  }
});
