export interface NextcloudSyncSettings {
  syncOnStartup: boolean;
  calendarTimezone: string;
  nextcloudDavUrl: string;
  caldavTaskListUrl: string;
  caldavCalendarUrl: string;
  caldavUsername: string;
  caldavPassword: string;
  taskPageName: string;
  calendarPageName: string;
  taskFilterPageTypes: string;
  taskFilterTags: string;
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
}

export type CalendarEventKind = "scheduled" | "deadline" | "date" | "start" | "end";

export interface LogseqCalendarEvent {
  uid: string;
  kind: CalendarEventKind;
  pageName: string;
  title: string;
  description: string;
  date: string;
  time?: string;
  allDay?: boolean;
  sourceBlockUuid?: string;
  sourceBlockContent?: string;
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
  enabled: boolean;
}

export interface DiscoveredTaskList {
  url: string;
  href: string;
  displayName: string;
  componentSet: string[];
  isTaskListCollection: boolean;
}
