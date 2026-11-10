# Excidian-dev validation

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
