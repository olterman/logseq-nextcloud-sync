# Logseq Nextcloud Sync

Standalone Logseq plugin for syncing Logseq tasks and calendar events with Nextcloud over CalDAV.

## What it does

- syncs Logseq task blocks to Nextcloud
- syncs Logseq calendar events to Nextcloud
- supports two-way task and calendar sync for linked items
- supports multiple saved sync profiles, each with its own filters and remote collection
- discovers existing Nextcloud collections and calendars
- can create a new remote Nextcloud collection for the active profile
- exports both tasks and calendar events as ICS
- can mirror task state changes both ways for simple checklist profiles

## Plugin pages

- `Nextcloud Tasks`
- `Nextcloud Calendar`
- `Nextcloud Calendar Picker`
- `Nextcloud Sync Profiles`
- `Nextcloud Profile Editor`

## Most useful commands

- `Nextcloud: Open profiles`
- `Nextcloud: Discover collections`
- `Nextcloud: Create profile`
- `Nextcloud: Create remote collection`
- `Nextcloud: Sync all`
- `Nextcloud: Open calendars`

## Important settings

- `nextcloudDavUrl`
  Usually `https://your-host/remote.php/dav`
- `caldavUsername`
- `caldavPassword`
- `caldavCalendarUrl`
- `calendarTimezone`

Profiles are mainly managed through the UI. Legacy single-scope settings still exist for compatibility.

## Sync Profiles

Each saved sync profile has:

- a name
- one remote collection URL
- a comma-separated `page-type` filter
- a comma-separated tag filter
- an optional comma-separated ignore-tag list
- an optional page prefilter mode
- an optional calendar property override such as `calendar:: personal`
- an optional `write to journal` flag for dated imports
- an optional default import page
- an optional `simple checklist mode`
- an optional periodic sync toggle and interval
- an enabled flag

If both page types and tags are set, a page matches when it matches either list.

`Only scan matched pages: yes` is faster, but it skips pages that only match through block-level overrides.

## Import Modes

Profiles currently work in two broad styles:

- Structured sync
  Best for journal-driven tasks and calendar/event sync.
- Simple checklist mode
  Best for shared list pages such as shopping lists.
  This mode:
  - skips calendar sync/import entirely
  - syncs only VTODO tasks
  - imports remote tasks as normal checklist items on the target page
  - tries to match existing local items by normalized title instead of relying on visible sync metadata

Simple checklist mode is intended for pages where readability matters more than preserving rich visible sync metadata.

## Tasks vs Events

Tasks are blocks that start with a task marker such as:

- `TODO`
- `DOING`
- `NOW`
- `WAITING`
- `DONE`

Example task:

```text
TODO Submit funding application DEADLINE: <2026-04-01>
calendar:: academia
```

Calendar events that are not tasks are normal blocks or pages with date fields or timestamps, but without a task marker.

Examples:

```text
Meeting with supervisor
date:: 2026-03-30 14:00
calendar:: academia
```

```text
Conference session
start:: 2026-04-02 09:00
end:: 2026-04-02 11:00
calendar:: academia
```

```text
Doctor appointment SCHEDULED: <2026-03-31 10:00>
calendar:: personal
```

A task block with `SCHEDULED` or `DEADLINE` is synced as a task and is not also exported as a separate calendar event by default.

## Periodic Sync

Each profile can optionally enable:

- `Update periodically: yes/no`
- `Update interval minutes: XX`

When enabled, the plugin runs that profile's sync on a repeating timer. Overlapping runs for the same profile are skipped.

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
