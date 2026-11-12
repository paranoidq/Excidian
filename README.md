# Excidian-dev

## Development preview (0.6.0-dev)

Click an `.xlsx` file in the file explorer or follow a link such as
`[[budget.xlsx]]` to open it in its own Obsidian tab. The tab has the same
worksheet controls and editing features as an embed. To open it externally, use
Obsidian's **Open in default app** context menu item or the toolbar's
spreadsheet icon.

Embed a workbook directly with Obsidian's attachment syntax:

```markdown
![[budget.xlsx]]
![[Finance/budget.xlsx#2026|Budget]]
```

Reading view and Live Preview show worksheet tabs below the grid. Click a tab to
switch sheets; an optional `#Sheet name` selects the initial sheet. Short file
names, aliases and link updates on rename use Obsidian's normal link resolution.
Click **在标签页打开** in the embed toolbar to open the current worksheet in a
workbook tab or switch to the one already open.

Previews start **read-only**. The toolbar button shows **预览模式** (preview mode)
or **编辑模式** (edit mode). Click it to switch modes. Settings → **Excidian-dev** → **默认只读模式** changes the default
on this device. Before entering edit mode, Excidian checks Excel's owner file.
If the workbook is already open in Excel, the preview stays read-only and prompts
**请先在 Excel 中关闭该文件，再切换到编辑模式**. This also applies to the
configured editable default and `mode: edit` blocks. Close the workbook in Excel,
then click the mode button again. If Excel opens it during editing, owner-file
notifications return the preview to read-only; save-time checks remain in place.
The rightmost **spreadsheet icon** opens the original workbook with the desktop
system's default application (for example Excel or WPS). Before opening, active
cell editors are committed and pending saves finish; every preview of the same
workbook switches to read-only. A failed save cancels opening, and a launch
failure is reported while the previews remain read-only. This icon is available
on desktop only and does not change the system's default file association.

Read-only mode disables cell editing and hides the append row.
An edit is saved when you press Enter/Tab, leave the cell, or change sheet/mode;
Escape cancels it. Worksheet/mode choices are local to each preview and do not
modify the note. Reloading the note uses the configured defaults again.

Existing `excidian` blocks still work, including `table`, `range`, `cols` and
`headings`. Add `mode: readonly` or `mode: edit` to override the initial mode for
one block. Sheet blocks have worksheet tabs; named table blocks retain their
focused table view.

### Everyday editing

- Single-click to select a cell; drag or Shift-click to select a rectangle, including
  in read-only mode. Double-click, F2 or Enter edits a selected cell. Typing replaces
  its content. Arrow keys navigate, Shift+arrows extend selection, Home/End navigate
  a row, and Cmd/Ctrl+Home/End navigate the preview. Enter/Tab commit and select
  the next cell; Escape cancels.
- Cmd/Ctrl+C copies displayed values as plain TSV. Cmd/Ctrl+V pastes a rectangular
  region; a single pasted value fills the selected rectangle. Quoted tabs, multiline
  cells and empty cells are supported. Delete/Backspace clears a rectangle without
  moving its cells or changing formatting. A paste/clear saves once and undoes once.
- A batch is limited to 10,000 cells and the currently displayed, consecutive region.
  Insert sufficient rows/columns first. Merged cells, table headers, protected cells,
  hidden/reordered row/column gaps and invalid Stop-validation values reject the
  whole batch. Paste does not transfer rich formatting or adjust relative formulas.
- Existing text, Text-formatted cells, leading-zero codes, long identifiers and
  intentional spaces stay text. Numbers use at most 15 significant decimal digits;
  longer input stays text. TRUE/FALSE become booleans in non-text cells. Prefix with
  an apostrophe to force text. Date/time input preserves hours, seconds and milliseconds.
- Use the magnifying glass or Cmd/Ctrl+F to search the visible worksheet. The
  preview follows the row order and hidden rows saved by Excel, including saved
  filter results. Change sorting or filtering in Excel and save the workbook.
- In an active cell editor, **Alt+Enter** (Option+Enter on macOS) inserts a line
  break. Enter saves and moves to the next cell; multiline values keep their
  breaks and use Excel's wrap formatting. Click a dropdown cell in edit mode
  to show all its choices right away.
- Select cells in edit mode to use the formatting toolbar: bold, italic,
  underline, strikethrough, fill and text color. With the grid focused,
  **Cmd/Ctrl+B** toggles bold and **Cmd/Ctrl+I** toggles italic on the selection.
  Color chevrons open the palette. The main color button reapplies its last color.
  Each command saves to the
  workbook as one undoable change, preserving values, formulas and number formats.
  Merged cells and protected sheets are excluded from formatting.
- The preview starts with the workbook's wrap settings. **自动换行** overrides
  them for the current preview and keeps line breaks entered in cells. Drag a
  column header's right edge to resize it locally or double-click the edge to
  restore its width.
  Cell tooltips show complete text. Web/email hyperlinks and plain URLs offer an
  **打开链接** action in the cell's context menu; local/executable links are not exposed.
- In edit mode, select adjacent worksheet cells and right-click **合并单元格**.
  The merged cell keeps the upper-left value. Excidian asks before clearing
  contents from the other selected cells. Right-click a merged cell and choose
  **取消合并** to split it. Both actions save to the workbook and support undo.

Formulas display Excel's **last saved calculation**, including cached results that
may become stale after inputs change. Formula tooltips and the preview note make
this explicit. Excidian does not contain a calculation engine. Protected sheets
and array/data-table formula cells must be edited in Excel. An Excel owner file
blocks saving while Excel has the workbook open. Save-time byte checks and
per-cell expected values reject detected external conflicts; this is not a
cross-application locking or collaborative editing system.

In editable **sheet previews**, click a row number or column letter to select it.
Drag across headers or Shift-click to select a consecutive range. Right-click
within that selection for **Insert**, **Delete** or **Clear contents**. Insert
adds the same number of rows above, or columns to the left. These actions apply
to entire rows/columns in the original worksheet, including cells outside a
`range`/`cols` preview. Clear keeps formatting and Excel table header names.
Delete/Backspace clears the selected rows/columns; Escape clears the selection.
An unrestricted editable sheet includes an empty entry row and column for
appending data. The viewport displays at most 5,000 rows, 256 existing columns and roughly
30,000 cells (wide sheets show fewer rows);
use a range or column selection to view other areas.

The toolbar has workbook **Undo / Redo** buttons. Cmd/Ctrl+Z and
Cmd/Ctrl+Shift+Z act on the workbook while its preview has focus. An input with
unsaved text keeps normal input undo. History is shared by previews of the same
file, stays in memory for the plugin session, and is limited to 20 changes per
file and 32 MiB total. Restoring a snapshot refuses to overwrite an external edit.

Existing Excel **frozen panes** are honored in both preview modes: leading rows
stay at the top and leading columns stay at the left while scrolling inside the
grid. Inserting/deleting before a frozen boundary updates that boundary too.

Structural edits update A1 formulas (including absolute and cross-sheet refs),
named ranges, merges, validation ranges/formulas, filters and table metadata.
Calculated table columns fill inserted data rows. Excel recalculates formulas
when the workbook is next opened; Excidian does not calculate formulas itself.
Named-table previews retain cell editing but do not offer whole-sheet menus.

Some structures require Excel: protected sheets, charts/drawings, comments,
pivot/query/external dependencies, array/data-table formulas, dynamic
INDIRECT/OFFSET and 3D references block structural edits with an explanation.
Literal INDIRECT references to a table column remain supported. Deleting a
table's header/totals row, its last data row or all its columns is blocked, as is
deleting a table column used by structured formulas/names. No file is written
when a structural edit is rejected. These limits do not disable frozen panes.

For local installation, run `npm ci && npm run build`, then copy `main.js`,
`manifest.json` and `styles.css` into `<vault>/.obsidian/plugins/excidian-dev/`
and enable **Excidian-dev**. Its plugin ID and device settings are separate from
the released plugin. The embed integration feature-detects Obsidian's internal
embed registry; if unavailable, the public fallback supports Reading view only.

---

**Your spreadsheets, alive in your notes.** Excidian shows Excel tables and sheets inside Obsidian, and every edit you make is saved straight back into the `.xlsx` file. No exports, no copies, no sync scripts.

![A note showing an Excel table of transactions, with a category dropdown open](https://raw.githubusercontent.com/FrancescoUmberto/ObXcel/main/docs/images/table-mode.png)

- **Tables** that grow as you type, with calculated columns filled in
- **Sheets** laid out like in Excel: widths, merged cells, fills, red negatives
- **Dropdowns and data validation** from your Excel rules
- **OneDrive-aware**: one note works on every computer; conflict copies are flagged
- **Safe**: saves check for changes from other programs before replacing the workbook

Website: [francescoumberto.github.io/ObXcel](https://francescoumberto.github.io/ObXcel/) · Desktop (macOS, Windows, Linux) and mobile (iOS, Android).

## Installation

- **From Obsidian:** Settings → Community plugins → Browse → search **Excidian** → Install → Enable.
- **Manually:** download `main.js`, `manifest.json` and `styles.css` from the [latest release](https://github.com/FrancescoUmberto/ObXcel/releases/latest) into `<vault>/.obsidian/plugins/obxcel/`, reload Obsidian and enable **Excidian**.

Requires Obsidian 1.13 or later.

## Desktop and mobile

| | Desktop | Mobile (iOS, Android) |
| --- | --- | --- |
| Workbooks inside the vault (`./bank.xlsx`, `Finance/bank.xlsx`) | ✓ | ✓ |
| Workbooks elsewhere (`~/…`, `/Users/…`, `C:\…`, `onedrive:/…`) | ✓ | — |
| Tables, sheets, editing, dropdowns, validation | ✓ | ✓ |
| Reload when the file changes, conflict-copy warnings | ✓ | ✓ |
| "Show file" in Finder / Explorer | ✓ | — |

Mobile apps can only reach files inside the vault. **To use the same note on your computer and your phone, keep the workbook inside the vault** and point to it with a relative path. Sync it with the vault (Obsidian Sync, iCloud, or similar) like any other file. A block that points outside the vault shows a message on mobile explaining this, and keeps working on desktop.

## File access and privacy

Excidian reads and writes the Excel files that your notes point to. **On desktop these can be outside your vault** (for example in your OneDrive folder); for this it uses Node's file system (`fs`), because Obsidian's vault API only covers files inside the vault. On mobile it uses only the vault API, so it only reaches files inside the vault. In detail, it:

- reads the workbook named in each `excidian` code block, and writes to it only when you edit a cell;
- briefly creates a hidden temporary file next to the workbook when saving on desktop, then uses it to replace the workbook;
- looks at the workbook's folder to notice changes and sync-conflict copies, and lists the OneDrive folders in your home folder to resolve `onedrive:` paths.

It touches no other files, makes **no network requests**, and collects no data. The OneDrive folder setting and dismissed warnings are stored on your device only.

## Usage

A block shows either an **Excel table** or a **sheet**.

**Table** (Insert → Table in Excel), by its name. Best for lists you keep adding to, like your transactions:

````markdown
```excidian
file: ~/Documents/Finance/bank.xlsx
table: Movimenti
```
````

The table's column names are the header, rows follow the saved order and Excel-hidden rows are omitted, and the empty row at the bottom adds a row to the table. A table always fits the width of the note: its columns share the width in proportion to their Excel widths.

**Sheet**, whole or part of it. Best for summaries, budgets and dashboards:

````markdown
```excidian
file: ~/Documents/Finance/bank.xlsx
sheet: Riepilogo
cols=[B, C, D]
headings: false
```
````

![A budget sheet with a merged title, coloured header row and an over-budget value in red](https://raw.githubusercontent.com/FrancescoUmberto/ObXcel/main/docs/images/sheet-mode.png)

A sheet is laid out the way Excel shows it:

- columns get their Excel widths; a sheet wider than the note scrolls sideways, with the row numbers kept in view;
- hidden rows and columns stay hidden;
- merged cells span their rows and columns;
- formatting is kept: bold, italic, underline, strikethrough, font and fill colours, alignment, wrapped text, and coloured numbers from formats like `#,##0.00;[Red]-#,##0.00`. Excel's plain black text and white background are left to your Obsidian theme, so sheets stay readable in dark mode;
- Excel tables inside the sheet get a highlighted header row, and typing in the row just below a table adds a row to it (like in Excel).

| Key        | Meaning                                                                                              |
| ---------- | ---------------------------------------------------------------------------------------------------- |
| `file`     | Required. Where the workbook is (see the table below).                                               |
| `table`    | Name of an Excel table (as in Table Design → Table Name). Table names are unique, so no sheet needed. |
| `sheet`    | Sheet name or 1-based position. Defaults to the first sheet.                                         |
| `range`    | With `sheet`: `A1:F200` for a fixed area, `A1:F` or `A:F` for "these columns, down to the last row". |
| `cols`     | With `sheet`: only these columns, in this order: `[B, C, D]`, `[B:D, F]`. Letters aren't case-sensitive; columns listed here show even if hidden in Excel. Overrides the columns of `range` (whose rows still apply). |
| `headings` | `false` hides column letters and row numbers (a table keeps its column names).                       |

Keys can be written `key: value` or `key=value`.

If the table or sheet isn't found, the error lists the ones that exist.

`file:` accepts:

| Example                              | Resolved from                                   |
| ------------------------------------ | ----------------------------------------------- |
| `./Sheets/bank.xlsx`, `../bank.xlsx` | the folder of the note, like a relative link    |
| `Sheets/bank.xlsx`                   | the root of the vault                           |
| `~/Documents/bank.xlsx`              | your home folder                                |
| `/Users/me/bank.xlsx`, `C:\…\bank.xlsx` | an absolute path, anywhere on the computer   |
| `onedrive:/Finance/bank.xlsx`        | this device's OneDrive folder (see below)       |

### Files in OneDrive

Write the path relative to your OneDrive folder, and the same note works on every computer:

````markdown
```excidian
file: onedrive:/Finance/bank.xlsx
table: Movimenti
```
````

Each computer finds its own OneDrive folder: `~/Library/CloudStorage/OneDrive-…` on macOS, the `%OneDrive%` folder on Windows, `~/OneDrive…` elsewhere. With several OneDrive accounts, the first one that contains the file is used. To choose a folder yourself, set **Settings → Excidian → OneDrive folder**. This setting is stored on each device and not in the vault, so syncing the vault doesn't copy one computer's path to another.

**Conflict copies.** When the same file is changed on two devices before they sync, OneDrive keeps both versions and saves one as a copy named after the device, e.g. `bank-MacBook-Pro.xlsx`. Excidian shows a warning above the table when such a copy appears next to your file. **Show file** opens it in Finder/Explorer so you can merge the entries in Excel and delete the copy. **Dismiss** hides that warning on this device (a newer copy warns again). Names like `bank (1).xlsx` count too, so an unrelated file such as `bank-2025.xlsx` will also trigger the warning; dismiss it once.

Tips:

- Mark the workbook **Always keep on this device** in OneDrive, so it's never an online-only placeholder.
- Avoid editing the file on two devices at the same time, and close Excel on this computer while editing in Obsidian.

### Editing

- Double-click a cell or press F2/Enter to edit it. **Alt+Enter** inserts a line break. **Enter** saves and moves down, **Tab** moves right (add Shift to go the other way), **Esc** cancels.
- The empty row at the bottom adds a new entry. For a table, the table grows by one row, the way typing below a table does in Excel. Calculated columns (like a running balance) are filled in, and the row gets the same formatting as the row above.
- What you type is converted automatically:
  - `-45`, `12,50`, `1500.00` → number
  - `2026-10-03` or `03/10/2026` (day/month/year) → date
  - `=SUM(B2:B40)` → formula
  - anything else → text; an empty cell clears it

### Dropdowns and data validation

Rules set in Excel under **Data → Data Validation** work in Obsidian too:

- **List rules** show a ▾ in the cell. Editing it opens a dropdown: type to filter, use ↑/↓ and Enter (or click) to pick. The list can come from anywhere Excel allows: a typed list (`Grocery,Health,Sport`), a range (`$H$2:$H$10`), a range on another sheet (`Lists!$A$2:$A$10`), a named range, or a table column through `INDIRECT("Categories[Type]")`.
- **Number, date and text-length rules** are checked when you save a cell. While editing, the rule's input message is shown under the cell, or a description like "Decimal between -10000 and 10000".
- **Stop** rules reject invalid values with the rule's error message, like Excel. **Warning** and **Information** rules save the value and show the message as a notice.
- Rules without "Show error alert" are only shown, never enforced. Custom (formula) rules are left for Excel to check.
- When a new row is added to a table, the rules of the table's columns grow with it, so the new row keeps its dropdowns in Excel too.

Tip: to manage a category list in one place, put the categories in their own table (e.g. `Categories`, column `Type`) and use `=INDIRECT("Categories[Type]")` as the list source. New categories added to that table appear in the dropdown automatically.

### Good to know

- **Tables fit the width of the note; sheets keep Excel's column widths** and scroll sideways when wider. Values too long for their column are cut off with "…" (unless the cell wraps text); hover a cell, or edit it, to see it in full.
- With `cols` in a different order than the sheet (e.g. `[D, B]`), merged cells don't span; their value shows in their first cell.
- Not shown: borders, fonts and font sizes, conditional formatting, charts and images.

- **Only the cells you edit change.** The rest of the workbook (other sheets, charts, pivot tables, formatting) is saved untouched.
- **Formulas are not calculated in Obsidian.** Excel recalculates everything the next time it opens the file. Until then, a newly written formula shows in grey italics.
- **Write formulas in English with commas**, e.g. `=SUM(A1;A2)` won't work but `=SUM(A1,A2)` will. That's how .xlsx files store them; Excel shows them in your language.
- **Close the file in Excel while editing in Obsidian.** If Excel has it open and you then save there, it overwrites the Obsidian edits. Excidian warns you when it detects this.
- Rows can't be added to a table with a totals row; turn the totals row off in Excel to use the empty row.
- Table headers can't be edited from Obsidian (rename columns in Excel).
- The table reloads on its own when the file changes on disk.

## Project structure (MVC)

```
src/
├── main.ts                    Plugin entry: picks the platform, wires settings, the settings tab and the code block
├── model/                     Excel data and rules. No Obsidian or Node imports: runs on desktop and mobile, testable in Node.
│   ├── workbook.ts            Workbook: sheets, tables, adding table rows, saving
│   ├── storage.ts             Storage interface (StoredFile), vault paths, conflict-copy names
│   ├── workbook-store.ts      Load, queued saves, undo/redo, conflict copies, "open in Excel" check
│   ├── device-settings.ts     Per-device settings (OneDrive folder, dismissed warnings)
│   ├── sheet.ts               Reading and writing cells in a worksheet
│   ├── table.ts               An Excel table: range, header/totals rows, columns
│   ├── validation.ts          Data validation rules: reading, describing, checking values
│   ├── values.ts              Cell values, parsing typed input, date serials
│   ├── number-format.ts       Displaying values with Excel number formats
│   └── ooxml/                 Low-level .xlsx format: zip package, XML, references, styles, colours, formulas
├── view/                      Drawing only. Knows nothing about files or Excel.
│   ├── grid-data.ts           The data shape the grid renders
│   ├── grid-view.ts           Grid rendering, warning banners and keyboard cell editing
│   ├── cell-popup.ts          Dropdown and help panel under the cell being edited
│   └── settings-tab.ts        The Settings → Excidian page
├── platform/                  Storage implementations
│   ├── vault-file.ts          Files inside the vault, via Obsidian's vault API (used on mobile)
│   └── desktop/               Desktop only, loaded on demand: files anywhere via Node's fs, ~ and onedrive: paths
└── controller/                Glue between the model and the view
    ├── block-config.ts        Parses the code block settings
    ├── grid-builder.ts        Turns the model into grid data
    └── sheet-controller.ts    Lifecycle of one code block: load, show, save edits, watch, warnings
```

The model edits the workbook's XML directly instead of going through a spreadsheet library, so parts it doesn't understand are preserved exactly.

## Development

```bash
npm install
npm run dev     # rebuild on change
npm run build   # type-check + production build
npm test        # model tests (Node, no Obsidian needed)
```

Install into a vault by copying `main.js`, `manifest.json` and `styles.css` to
`<vault>/.obsidian/plugins/obxcel/`, then enable **Excidian** under Settings → Community plugins.
For development, symlink the project folder there instead:

```bash
ln -s "$PWD" "<vault>/.obsidian/plugins/obxcel"
```

## Releasing

Every push to `main` runs [`.github/workflows/ci.yml`](.github/workflows/ci.yml):

1. **Build**: installs, runs the tests, builds `main.js`, and checks that `package.json`, `manifest.json` and `versions.json` agree on the version. The three plugin files are attached to the run as a download.
2. **Release**: if no release exists yet for the version in `manifest.json`, it creates one named and tagged exactly like the version (e.g. `0.1.0`, no `v`, as Obsidian requires) with `main.js`, `manifest.json` and `styles.css` attached. If that version is already released, this step is skipped.
3. **Website**: publishes `docs/` to GitHub Pages, with the current version filled in.

To publish a new version:

```bash
npm version patch   # or minor / major: updates package.json, manifest.json and versions.json, and commits
git push
```

One-time setup on GitHub: **Settings → Pages → Source: GitHub Actions**.

## License

[MIT](LICENSE) © 2026 Umberto Francesco Carolini
