# Excidian-dev validation

## 0.6.0-dev: Excel-style formatting and multiline cells

Validated on 2026-11-11 in `my-new-wiki`, Obsidian 1.13.7. Production build,
all 110 tests and `git diff --check` passed.

- Removed the local sort, filter, range selection, reset and address lookup
  panel. Search and preview wrap are separate icons. Both worksheet and named
  table views show the saved row order and omit Excel-hidden rows.
- Added bold, italic, underline, strikethrough, fill and font color controls in
  edit mode. Each selection command updates cell styles in one undoable write.
  Format-only writes preserve formulas, cached results, number formats, and
  untouched parts. Style records are reused; a missing stylesheet is created.
- A multiline editor handles Alt+Enter through a temporary Obsidian shortcut
  scope, released on commit/cancel. Enter saves, Tab navigates, and an explicit
  line break is stored with Excel wrap formatting. List-validation choices open
  immediately on clicking the cell in edit mode.
- Created `样式与换行验证.md` with a disposable workbook containing saved filter
  results on a sheet and on a named table. Through the installed Live Preview
  DOM, verified hidden rows, search/wrap, Alt+Enter followed by Enter, the
  default dropdown with five options, and a 2×2 selection with bold, underline,
  fill and text color, including undo/redo. Independently reread the saved XLSX
  and checked its line break, styles, dropdown choice, numeric value and saved
  filter/sort metadata. Restored only the generated workbook.
- Unit tests cover saved order and hidden rows, format/line-break round trips,
  style reuse, missing styles, protected editing, and atomic undo.

## 0.5.2-dev: open with the local default application

Validated on 2026-11-10 in `my-new-wiki`, Obsidian 1.13.7. Production build,
all 107 tests and `git diff --check` passed.

- Added a rightmost spreadsheet icon with the tooltip “用本地默认程序打开（Excel / WPS）”.
  Desktop files use Electron `shell.openPath` with the original absolute path;
  the OS chooses the default application. Platforms without this capability hide it.
- Before handoff, active editors for the same workbook are committed, all their
  previews become readonly, and pending saves are awaited. A failed save prevents
  launching. Duplicate clicks are blocked while pending; an OS launch failure is
  reported and the previews stay readonly.
- Tests cover readonly launch, active-editor save ordering, save failure, launch
  failure, duplicate clicks, unload cancellation, unavailable platforms, multiple
  previews and exact Unicode/space/quote-containing paths passed to Electron.
- In the installed preview, edited B2 without pressing Enter and clicked the new
  icon. A temporary spy on the OS API confirmed that it received the original
  workbook path only after the input was saved and all editors were gone in
  readonly mode. Independently read the saved XLSX to verify the new value, then
  restored the generated fixture from its exact pre-test bytes.
- Restored the real OS API and clicked the icon again. It successfully opened
  `常用功能样例.xlsx` in the installed WPS Office default application. Observed
  the matching document window and closed only that verification workbook.
  The Obsidian preview remained readonly. Existing user workbooks were not edited.

## 0.5.1-dev: block editing while Excel owns the workbook

Validated on 2026-11-10 in `my-new-wiki`, Obsidian 1.13.7. Production build,
all 98 tests and `git diff --check` passed.

- Mode changes check Excel's owner file before exposing editable cells. While
  checking, the preview stays read-only and its mode button is busy/disabled.
- Detected ownership keeps the preview read-only with a notice asking the user
  to close that workbook in Excel. Editable defaults and `mode: edit` use the
  same initial check. Inspection errors fail closed; stale asynchronous mode
  results cannot re-enable a stopped preview or override a newer readonly request.
- Owner-file creation/deletion is watched. Opening the workbook while the preview
  is editable returns it to readonly; closing it does not automatically enable
  editing. Save-time ownership checks remain as a fallback.
- Added controller coverage for blocked/allowed toggles, default/block overrides,
  inspection failures, pending-check cancellation/unload, owner-file notifications
  and Excel opening just before a save. Vault watcher tests cover both owner names.
- Installed `0.5.1-dev` and reopened the verification note. Simulated Excel ownership
  with a disposable `~$常用功能样例.xlsx`: confirmed the readonly/busy state,
  the user notice, successful manual retry after removal, and automatic return
  to readonly on recreation. Removed only that temporary owner file; workbook
  bytes remained identical. No Microsoft Excel installation was available for
  an actual application-open check on this machine.

## 0.5.0-dev: reliable everyday editing

Validated on 2026-11-10 in `my-new-wiki`, Obsidian 1.13.7. Production build,
all 91 tests and `git diff --check` passed.

- Added `常用功能验证.md` and a separate generated `常用功能样例.xlsx` with
  identifiers, numbers, booleans, dates, multiline notes and frozen leading panes.
  `test/manual-common-fixture.ts` reproduces this fixture and refuses to overwrite files.
- Exercised the installed plugin through its DOM events: read-only 2×2 selection
  and TSV copy, arrow navigation, mode toggle, 2×2 paste, one-step undo/redo,
  F2/Enter editing, search/next, original-address lookup and local wrapping.
- Selected the data region, sorted numeric values, edited C21 while its row was
  displayed elsewhere, and independently read the saved workbook: C21 changed,
  C2 stayed unchanged, zero-prefixed/long IDs remained strings, and TRUE remained
  a boolean. A rejected paste with a later invalid category left the whole file
  byte-for-byte unchanged; read-only paste was also blocked. Local contains
  filtering showed only matches plus the header;
  resetting restored original order.
- Used trusted mouse events via Obsidian's CDP developer command to drag a column
  boundary; measured the increased width. Inspected screenshots of collapsed
  and expanded tools and long-text wrapping.
- New tests cover TSV quoting and multiline cells, batch bounds and gaps,
  merged/read-only guards, atomic validation rejection, stable numeric sorting,
  composed filters, value/type fidelity, protected/array formula cells, invalid
  XML/Excel values, stale-cell rejection, save-time external conflicts, and an
  external change arriving immediately after an own save.
- Restored only this generated workbook from its own pre-edit backup after checks.
  Plugin reload clears temporary view settings/history and leaves the default
  read-only mode in effect. Existing user workbooks were not changed.

Scope remains ordinary list maintenance: no formula calculation engine,
physical row sorting, rich clipboard formatting, arbitrary cell movement,
chart/pivot authoring or concurrent editing. Filtering/sorting uses displayed
cached values and is local to one preview. Large/hidden/merged regions have
explicit batch-edit limits documented in README. A roughly 30,000-cell render
budget keeps wide workbooks from creating an excessive DOM. Reloaded plugin
previews were reopened before final checks.

## 0.4.0-dev: worksheet structure and frozen panes

Validated on 2026-11-10 in `my-new-wiki`, Obsidian 1.13.7. Build, all 71 tests,
and `git diff --check` passed.

- Generated a separate 80-row, 12-column fixture with the first two rows and
  columns frozen: `Excidian-dev 验证/行列与冻结验证.md` and its workbook.
- In Live Preview, selected rows 4–6 with Shift-click and right-clicked row 5.
  One worksheet menu appeared, without Obsidian editor items. Inserted three
  rows above the selection; independently read the saved workbook and confirmed
  `A4` was blank and original `A4` moved to `A7`.
- Exercised toolbar Undo and Redo. Selected column C, inserted a column to its
  left, deleted that inserted column, cleared column C, then used Cmd+Z to undo
  the clear. Verified clearing C did not move D in the saved file.
- Scrolled the grid 350 px horizontally and 700 px vertically. Measured the
  first two rows/columns at unchanged coordinates; non-frozen cells moved by
  the scroll offsets. Inspected a captured screenshot and fixed sticky border
  artifacts using separate table borders.
- Tests cover shared/absolute/cross-sheet formulas, names, table metadata and
  calculated columns, validation extensions, merges, hidden columns, frozen
  boundaries, blank extents, hyperlinks, read-only guards, save sequencing,
  exact-byte undo/redo, external-change rejection and unsupported dependencies.
- Restored only the generated structure fixture after validation. Pre-existing
  user workbooks were not edited.

Whole-sheet menus are intentionally available in sheet previews; named-table
previews retain their existing cell editing behavior. Unsupported dependencies
and table boundaries are described in README and rejected before writing.

## 0.3.0-dev: embeds and modes

Validated on 2026-11-10 in Obsidian 1.13.7, vault `my-new-wiki`.

- Plugin enabled as `excidian-dev`, display name `Excidian-dev`, version `0.3.0-dev`.
- Installed files: `.obsidian/plugins/excidian-dev/{main.js,manifest.json,styles.css}`.
- Test note: `Excidian-dev 验证/功能验证.md`.
- Test workbook: `Excidian-dev 验证/多工作表样例.xlsx` (generated fixture, three worksheets).
- `npm run build` passed; `npm test` passed all 56 tests; `git diff --check` passed.

## In-app checks

1. Reading view rendered `![[多工作表样例.xlsx]]` and all three worksheet tabs.
2. `![[多工作表样例.xlsx#2026|交易记录]]` selected the named initial worksheet.
3. Clicking a cell in read-only mode did not open an input.
4. Switching to editable mode displayed the append row and enabled cell inputs.
5. Edited `Liste!A2`, then clicked another sheet before the save completed. The
   edit was saved to `Liste!A2`; `Riepilogo!A2` stayed empty. A separate read-only
   preview of the same workbook refreshed to show the saved value.
6. A newly opened Live Preview rendered both wiki embeds with all worksheet tabs.
   Switched sheets, enabled editing, changed `2026!C2` from 1500 to 1600 and saved
   with Enter. The displayed value and the underlying `.xlsx` both contained 1600.
7. Turned off **默认只读模式** in the settings UI. Existing embed previews became
   editable after closing settings; the code block with `mode: readonly` stayed
   read-only. Restored the setting to read-only and confirmed both embeds followed.
8. Restored the generated workbook from its original session backup after checking
   the saved values from disk. No pre-existing workbook was edited.

Obsidian caches existing Live Preview widgets. Embeds that were already open
before enabling the plugin may retain their old attachment appearance until the
note is reopened. New tabs and Reading view use the registered workbook renderer.

The embed registry is an internal Obsidian API, isolated behind feature detection.
The public Markdown post-processor fallback supports Reading view only. Mobile
and older Obsidian versions were not exercised in this desktop validation.
