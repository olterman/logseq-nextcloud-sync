export interface NextcloudSyncSettings {
  syncOnStartup: boolean;
  calendarTimezone: string;
  nextcloudDavUrl: string;
  graphRootPath?: string;
  simpleChecklistMode?: boolean;
  syncProfilesJson: string;
  importedCalendarUidCacheJson?: string;
  calendarSyncStateJson?: string;
  activeSyncProfileId: string;
  caldavTaskListUrl: string;
  caldavCalendarUrl: string;
  calendarScopesJson: string;
  activeCalendarScopeId: string;
  caldavUsername: string;
  caldavPassword: string;
  taskPageName: string;
  calendarPageName: string;
  taskFilterPageTypes: string;
  taskFilterTags: string;
  taskIgnoreTags: string;
  prefilterPagesOnly?: boolean;
  taskScopesJson: string;
  activeTaskScopeId: string;
}

export interface LogseqTaskItem {
  uid: string;
  pageName: string;
  title: string;
  description: string;
  date?: string;
  time?: string;
  allDay?: boolean;
  unscheduled?: boolean;
  sourceBlockUuid?: string;
  sourceBlockContent?: string;
  taskState?: string;
  marker: string;
  pageType?: string;
  pageTags: string[];
  remoteResourceUrl?: string;
}

export type CalendarEventKind = "scheduled" | "deadline" | "date" | "start";

export interface LogseqCalendarEvent {
  uid: string;
  kind: CalendarEventKind;
  pageName: string;
  title: string;
  description: string;
  date: string;
  time?: string;
  allDay?: boolean;
  scopePropertyKey?: string;
  scopePropertyValue?: string;
  sourceBlockUuid?: string;
  sourceBlockContent?: string;
  remoteResourceUrl?: string;
}

export interface DiscoveredCalendar {
  url: string;
  href: string;
  displayName: string;
  componentSet: string[];
  isCalendarCollection: boolean;
}

export interface TaskScopeConfig {
  id: string;
  name: string;
  taskListUrl: string;
  filterPageTypes: string;
  filterTags: string;
  ignoredTags: string;
  prefilterPagesOnly?: boolean;
  enabled: boolean;
}

export interface SyncProfileConfig {
  id: string;
  name: string;
  remoteUrl?: string;
  taskListUrl: string;
  filterPageTypes: string;
  filterTags: string;
  ignoredTags: string;
  prefilterPagesOnly?: boolean;
  calendarUrl: string;
  propertyKey: string;
  propertyValue: string;
  writeToJournal?: boolean;
  defaultImportPage?: string;
  simpleChecklistMode?: boolean;
  updatePeriodically?: boolean;
  updateIntervalMinutes?: number;
  enabled: boolean;
}

export interface CalendarScopeConfig {
  id: string;
  name: string;
  calendarUrl: string;
  filterPageTypes?: string;
  filterTags?: string;
  ignoredTags?: string;
  prefilterPagesOnly?: boolean;
  propertyKey: string;
  propertyValue: string;
  enabled: boolean;
}

export interface DiscoveredTaskList {
  url: string;
  href: string;
  displayName: string;
  componentSet: string[];
  isTaskListCollection: boolean;
}
