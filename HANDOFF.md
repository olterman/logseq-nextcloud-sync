# Handoff

## Project

`/home/olterman/Projects/logseq-nextcloud-sync`

This is a standalone Logseq plugin project, moved out from the earlier monorepo workspace. It is now symlinked into:

`/home/olterman/.logseq/plugins/logseq-nextcloud-sync`

## Current architecture

- `src/main.ts`
  Plugin entry point, settings schema, command registration, renderer-backed UI pages, task-scope manager flow, calendar selector flow.
- `src/caldav.ts`
  Task collection, task sync, task-list discovery, task-list creation, task completion mirroring, scope-aware task filtering.
- `src/calendar.ts`
  Calendar event collection, calendar sync, calendar discovery, calendar selector support.
- `src/types.ts`
  Shared plugin settings and domain types.

## Implemented features

- Generic task sync, not tied to the academia plugin structure.
- Scopeable task filtering by page type and/or page tags.
- Multiple saved task scopes, each with:
  - name
  - `page-type` filter string
  - tag filter string
  - bound Nextcloud task-list URL
  - enabled flag
- Active scope selection.
- Sync-one-scope and sync-all-enabled-scopes flows.
- Discovery of Nextcloud VEVENT calendars.
- Discovery of Nextcloud VTODO task lists.
- Creation of remote Nextcloud VTODO task lists for the active scope.
- Selector/manager pages rendered inside Logseq.
- ICS export for tasks and calendar events.
- Mirroring of remote completed tasks back into matching Logseq blocks.

## Important UI pages

- `Nextcloud Tasklist Scopes`
  Main task scope manager.
- `Nextcloud Calendar Picker`
  Calendar discovery and selection page.
- `Nextcloud Tasks`
  Snapshot page for the currently refreshed/active scope.
- `Nextcloud Calendar`
  Snapshot page for discovered calendar events.

## Important compatibility note

There is a compatibility fallback from the older single-scope model:

- If `taskScopesJson` is empty, the plugin synthesizes a legacy default scope from:
  - `caldavTaskListUrl`
  - `taskFilterPageTypes`
  - `taskFilterTags`

Once named scopes are created and saved, those become the primary source of truth.

## Suggested next steps

1. Add delete/duplicate controls for task scopes in the manager UI.
2. Add per-scope test-connection UI and per-scope export UI buttons.
3. Consider adding `AND` vs `OR` filter logic for page-type/tag matching.
4. Improve task-list and calendar discovery error presentation on the manager pages.
5. Consider replacing prompt-based scope editing with a richer in-page form.
6. Add persistent preview summaries per scope instead of only the active scope snapshot page.

## Verification completed

- `./node_modules/.bin/tsc -p tsconfig.json --noEmit`
- `node esbuild.mjs`

Both were passing before handoff.
