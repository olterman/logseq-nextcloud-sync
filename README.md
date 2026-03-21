# Logseq Nextcloud Sync

Standalone Logseq plugin for syncing Logseq tasks and calendar events with Nextcloud over CalDAV.

## What it does

- syncs Logseq task blocks to Nextcloud task lists
- supports multiple saved task scopes, each with its own filters and target task list
- discovers existing Nextcloud task lists and calendars
- can create a new remote Nextcloud task list for the active task scope
- syncs Logseq calendar events to a selected Nextcloud calendar
- exports both tasks and calendar events as ICS
- mirrors completed Nextcloud tasks back into Logseq blocks

## Plugin pages

- `Nextcloud Tasks`
- `Nextcloud Calendar`
- `Nextcloud Calendar Picker`
- `Nextcloud Tasklist Scopes`

## Most useful commands

- `Nextcloud Sync: Open task scope manager`
- `Nextcloud Sync: Discover task lists`
- `Nextcloud Sync: Create task scope`
- `Nextcloud Sync: Create remote task list for active scope`
- `Nextcloud Sync: Sync all task scopes to Nextcloud`
- `Nextcloud Sync: Open calendar selector`
- `Nextcloud Sync: Discover calendars`

## Important settings

- `nextcloudDavUrl`
  Usually `https://your-host/remote.php/dav`
- `caldavUsername`
- `caldavPassword`
- `caldavCalendarUrl`
- `calendarTimezone`

Task scopes are mainly managed through the UI, but the underlying settings are:

- `taskScopesJson`
- `activeTaskScopeId`

Legacy single-scope settings still exist for compatibility:

- `caldavTaskListUrl`
- `taskFilterPageTypes`
- `taskFilterTags`

## Task scopes

Each saved task scope has:

- a name
- a comma-separated `page-type` filter
- a comma-separated tag filter
- a selected Nextcloud task-list URL
- an enabled flag

If both page types and tags are set, a page is included when it matches either list.

## Development

Build:

```bash
cd ~/Projects/logseq-nextcloud-sync
npm run build
```

Typecheck:

```bash
cd ~/Projects/logseq-nextcloud-sync
./node_modules/.bin/tsc -p tsconfig.json --noEmit
```

The Logseq plugin symlink points here:

```text
~/.logseq/plugins/logseq-nextcloud-sync
```
