import assert from "node:assert/strict";
import { build } from "esbuild";
import { createRequire } from "node:module";
import { describe, it } from "node:test";
import { DeviceSettings } from "../src/model/device-settings";
import type { StoredFile } from "../src/model/storage";
import { Workbook } from "../src/model/workbook";
import type { GridCallbacks } from "../src/view/grid-view";
import type { GridData } from "../src/view/grid-data";
import { financeWorkbook } from "./fixture";

(globalThis as unknown as { window: typeof globalThis }).window ??= globalThis;

// Exercise the real controller, workbook and save queue with only the Obsidian
// lifecycle and DOM view replaced. This catches edits redirected by tab changes.
async function harness(source = "file: demo.xlsx", defaultReadOnly = true, initiallyOpenInExcel = false) {
	const result = await build({
		entryPoints: ["src/controller/sheet-controller.ts"], bundle: true, write: false,
		platform: "node", format: "cjs", external: ["obsidian"],
		plugins: [{ name: "headless-grid", setup(build) {
			build.onResolve({ filter: /\/grid-view$/ }, () => ({ path: "grid", namespace: "test" }));
			build.onLoad({ filter: /.*/, namespace: "test" }, () => ({ contents: `
				export class GridView {
					constructor(root, callbacks) { root.view = this; this.callbacks = callbacks; this.renders = 0; }
					render(grid) { this.grid = grid; this.renders++; }
					showError(error) { this.error = error; }
					finishEdit() { this.onFinishEdit?.(); } cancelEdit() {} destroy() { this.destroyed = true; } setStatus() {}
				}`, loader: "js" }));
		} }],
	});
	const module = { exports: {} as any };
	const notices: string[] = [];
	const require = createRequire(import.meta.url);
	new Function("require", "module", "exports", result.outputFiles[0].text)(
		(id: string) => id === "obsidian" ? {
			MarkdownRenderChild: class {}, Notice: class { constructor(text: string) { notices.push(text); } },
		} : require(id), module, module.exports,
	);
	let bytes = await financeWorkbook();
	let writes = 0;
	let openCount = 0;
	let onOpen: (() => Promise<void>) | undefined;
	let onWrite: (() => Promise<void>) | undefined;
	let watchers = 0;
	let openInExcel = initiallyOpenInExcel;
	let inspection: (() => Promise<void>) | undefined;
	let notify: (() => void) | undefined;
	const file: StoredFile = {
		id: `test-${Math.random()}`, name: "demo.xlsx", displayPath: "demo.xlsx",
		read: async () => bytes,
		write: async (value) => { await onWrite?.(); bytes = value; writes++; },
		openLocal: async () => { openCount++; await onOpen?.(); },
		writeSibling: async () => {}, siblings: async () => { await inspection?.(); return openInExcel ? [{ name: "~$demo.xlsx", modified: new Date() }] : []; },
		watch: (handler) => { notify = handler; watchers++; return () => { watchers--; }; },
	};
	const data = new Map<string, unknown>();
	const settings = new DeviceSettings({ load: (key) => data.get(key), save: (key, value) => { data.set(key, value); } });
	settings.defaultReadOnly = defaultReadOnly;
	const active = new Set();
	const root = {} as { view: { callbacks: GridCallbacks; grid: GridData; renders: number; error?: string; destroyed?: boolean; onFinishEdit?: () => void } };
	const controller = new module.exports.SheetController(root, source, "note.md", { settings, resolver: { resolve: () => file }, active });
	controller.onload();
	await until(() => (!!root.view.grid && !root.view.grid.busy) || !!root.view.error);
	assert.equal(root.view.error, undefined);
	return { controller, file, view: root.view, settings, active, notices, opens: () => openCount, setOpen: (handler: () => Promise<void>) => { onOpen = handler; }, setWrite: (handler: () => Promise<void>) => { onWrite = handler; }, setExcelOpen: (open: boolean) => { openInExcel = open; }, setInspection: (handler?: () => Promise<void>) => { inspection = handler; }, notify: () => notify?.(), read: () => Workbook.load(bytes), writes: () => writes, watchers: () => watchers, external: async (text: string) => { const wb = await Workbook.load(bytes); wb.sheet("Liste").setCell(2, 1, { kind: "text", value: text }); bytes = await wb.save(); notify?.(); } };
}

async function until(check: () => boolean) {
	for (let i = 0; i < 300; i++) {
		if (check()) return;
		await new Promise((resolve) => setTimeout(resolve, 10));
	}
	throw new Error("Timed out waiting for preview");
}

describe("preview controller", () => {
	it("blocks structural edits while read-only and supports selected-sheet edit/undo/redo", async () => {
		const h = await harness();
		try {
			await h.controller.structure({ axis: "row", kind: "insert", start: 3, count: 2 });
			assert.equal(h.writes(), 0);
			h.view.callbacks.onSheetChange("2026"); await h.view.callbacks.onModeChange(false);
			await h.controller.structure({ axis: "row", kind: "insert", start: 3, count: 2 });
			assert.equal(h.view.grid.busy, false);
			assert.equal((await h.read()).sheet("2026").cell(5, 3).value, -45);
			assert.equal(h.view.grid.canUndo, true);
			await h.controller.history("undo");
			assert.equal((await h.read()).sheet("2026").cell(3, 3).value, -45);
			assert.equal(h.view.grid.canRedo, true);
			await h.controller.history("redo");
			assert.equal((await h.read()).sheet("2026").cell(5, 3).value, -45);
			assert.deepEqual(h.notices, []);
		} finally { h.controller.onunload(); }
	});

	it("finishes a cell save before applying a structural edit", async () => {
		const h = await harness();
		try {
			h.view.callbacks.onSheetChange("Liste"); await h.view.callbacks.onModeChange(false);
			h.view.callbacks.onCommit(h.view.grid.rows[1].cells[0], "saved before insert");
			await h.controller.structure({ axis: "row", kind: "insert", start: 2, count: 1 });
			assert.equal((await h.read()).sheet("Liste").cell(3, 1).value, "saved before insert");
			assert.deepEqual(h.notices, []);
		} finally { h.controller.onunload(); }
	});
	it("enforces read-only at the save boundary, with independent per-preview toggles", async () => {
		const h = await harness();
		try {
			const cell = { ...h.view.grid.rows[0].cells[0], editable: true };
			await h.controller.commit(cell, "blocked");
			assert.equal(h.writes(), 0);
			await h.view.callbacks.onModeChange(false);
			assert.equal(h.view.grid.readOnly, false);
			assert.equal(h.settings.defaultReadOnly, true);
			await h.view.callbacks.onModeChange(true);
			assert.ok(h.view.grid.rows.every((r) => r.cells.every((c) => !c.editable)));
		} finally { h.controller.onunload(); }
		assert.equal(h.watchers(), 0);
		assert.equal(h.active.size, 0);
		assert.equal(h.view.destroyed, true);
	});

	it("saves to the selected sheet even when another tab is selected before the save finishes", async () => {
		const h = await harness();
		try {
			await h.view.callbacks.onModeChange(false);
			h.view.callbacks.onSheetChange("Liste");
			const cell = h.view.grid.rows[1].cells[0];
			const saving = h.controller.commit(cell, "Edited category");
			h.view.callbacks.onSheetChange("Riepilogo");
			await saving;
			assert.equal(h.writes(), 1);
			assert.equal((await h.read()).sheet("Liste").cell(2, 1).value, "Edited category");
			assert.equal((await h.read()).sheet("Riepilogo").cell(2, 1).value, null);
			assert.equal(h.view.grid.activeSheet, "Riepilogo");
			assert.deepEqual(h.notices, []);
		} finally { h.controller.onunload(); }
	});

	it("applies the configured default on restart and honors a block override", async () => {
		const h = await harness("file: demo.xlsx", false);
		try {
			assert.equal(h.view.grid.readOnly, false);
			h.settings.defaultReadOnly = true;
			h.controller.restart();
			await until(() => h.view.grid.readOnly === true);
			assert.equal(h.watchers(), 1);
		} finally { h.controller.onunload(); }
		const override = await harness("file: demo.xlsx\nmode: readonly", false);
		assert.equal(override.view.grid.readOnly, true);
		override.controller.onunload();
	});

	it("does not render a pending reload after the preview is unloaded", async () => {
		const h = await harness();
		const renders = h.view.renders;
		h.view.callbacks.onRefresh();
		h.controller.onunload();
		await new Promise((resolve) => setTimeout(resolve, 30));
		assert.equal(h.view.renders, renders);
	});
});

describe("batch edits", () => {
	it("keeps read-only at the controller boundary", async () => {
		const h = await harness();
		try { await h.controller.batchCommit([{ cell: { ...h.view.grid.rows[1].cells[0], editable: true }, text: "blocked" }]); assert.equal(h.writes(), 0); }
		finally { h.controller.onunload(); }
	});
	it("saves the full region once, preserves IDs, and undoes the region in one step", async () => {
		const h = await harness();
		try {
			await h.view.callbacks.onModeChange(false); h.view.callbacks.onSheetChange("Liste");
			const original = (await h.read()).sheet("Liste").cell(2, 1).value;
			await h.controller.batchCommit([
				{ cell: h.view.grid.rows[1].cells[0], text: "00123" },
				{ cell: h.view.grid.rows[2].cells[0], text: "123456789012345678" },
			]);
			assert.equal(h.writes(), 1);
			assert.equal((await h.read()).sheet("Liste").cell(2, 1).value, "00123");
			assert.equal((await h.read()).sheet("Liste").cell(3, 1).value, "123456789012345678");
			await h.controller.history("undo");
			assert.equal((await h.read()).sheet("Liste").cell(2, 1).value, original);
			assert.equal(h.view.grid.canUndo, false);
			assert.deepEqual(h.notices, []);
		} finally { h.controller.onunload(); }
	});
	it("rejects an entire region when a later cell fails validation", async () => {
		const h = await harness();
		try {
			await h.view.callbacks.onModeChange(false); h.view.callbacks.onSheetChange("2026");
			const row = h.view.grid.rows.find((r) => r.label === "3")!;
			await h.controller.batchCommit([{ cell: row.cells[2], text: "-20" }, { cell: row.cells[3], text: "invalid category" }]);
			assert.equal(h.writes(), 0);
			assert.equal((await h.read()).sheet("2026").cell(3, 3).value, -45);
			assert.ok(h.notices.some((notice) => notice.includes("整个区域未保存")));
		} finally { h.controller.onunload(); }
	});
	it("rejects a stale cell rather than overwriting a newer value", async () => {
		const h = await harness();
		try {
			await h.view.callbacks.onModeChange(false); h.view.callbacks.onSheetChange("Liste");
			const stale = h.view.grid.rows[1].cells[0];
			await h.controller.commit(stale, "new value");
			await h.controller.commit(stale, "stale overwrite");
			assert.equal(h.writes(), 1);
			assert.equal((await h.read()).sheet("Liste").cell(2, 1).value, "new value");
			assert.ok(h.notices.some((notice) => notice.includes("已被其他预览或程序修改")));
		} finally { h.controller.onunload(); }
	});
});


it("reloads an external change that arrives just after its own save", async () => {
	const h = await harness();
	try {
		await h.view.callbacks.onModeChange(false); h.view.callbacks.onSheetChange("Liste");
		await h.controller.commit(h.view.grid.rows[1].cells[0], "ours");
		await h.external("external after save");
		await until(() => h.view.grid.rows[1].cells[0].text === "external after save");
		assert.equal(h.writes(), 1);
	} finally { h.controller.onunload(); }
});

describe("Excel ownership mode guard", () => {
	it("blocks switching to edit, notifies, and permits a new attempt after Excel closes", async () => {
		const h = await harness();
		try {
			h.setExcelOpen(true);
			const switching = h.controller.changeMode(false);
			assert.equal(h.view.grid.readOnly, true);
			assert.equal(h.view.grid.busy, true);
			await switching;
			assert.equal(h.view.grid.readOnly, true);
			assert.equal(h.view.grid.busy, false);
			assert.ok(h.notices.some((text) => text.includes("请先在 Excel 中关闭该文件")));
			assert.equal(h.writes(), 0);
			h.setExcelOpen(false);
			await h.view.callbacks.onModeChange(false);
			assert.equal(h.view.grid.readOnly, false);
		} finally { h.controller.onunload(); }
	});
	it("does not allow default-edit or a mode: edit block to bypass ownership", async () => {
		for (const [source, defaultReadOnly] of [["file: demo.xlsx", false], ["file: demo.xlsx\nmode: edit", true]] as const) {
			const h = await harness(source, defaultReadOnly, true);
			try {
				assert.equal(h.view.grid.readOnly, true);
				assert.ok(h.view.grid.rows.every((row) => row.cells.every((cell) => !cell.editable)));
				assert.equal(h.notices.length, 1);
				assert.equal(h.writes(), 0);
			} finally { h.controller.onunload(); }
		}
	});
	it("keeps readonly when the ownership check fails", async () => {
		const h = await harness();
		try {
			h.setInspection(async () => { throw new Error("inspection failed"); });
			await h.view.callbacks.onModeChange(false);
			assert.equal(h.view.grid.readOnly, true);
			assert.equal(h.view.grid.busy, false);
			assert.ok(h.notices.some((text) => text.includes("无法确认文件是否被 Excel 占用")));
		} finally { h.controller.onunload(); }
	});
	it("discards a pending edit-mode check after returning to readonly or unloading", async () => {
		for (const unload of [false, true]) {
			const h = await harness();
			let resolve!: () => void;
			h.setInspection(() => new Promise<void>((done) => { resolve = done; }));
			const switching = h.controller.changeMode(false);
			if (unload) h.controller.onunload(); else await h.controller.changeMode(true);
			const renders = h.view.renders;
			resolve(); await switching;
			assert.equal(h.view.grid.readOnly, true);
			assert.equal(h.view.renders, renders);
			assert.deepEqual(h.notices, []);
			if (!unload) h.controller.onunload();
		}
	});
	it("returns to readonly if Excel is opened after edit mode was enabled", async () => {
		const h = await harness();
		try {
			await h.view.callbacks.onModeChange(false);
			h.setExcelOpen(true); h.notify();
			await until(() => h.view.grid.readOnly === true);
			assert.ok(h.notices.some((text) => text.includes("Excel 中打开")));
			assert.equal(h.writes(), 0);
		} finally { h.controller.onunload(); }
	});
	it("also blocks saves when Excel opens before its watcher notification arrives", async () => {
		const h = await harness();
		try {
			await h.view.callbacks.onModeChange(false);
			h.view.callbacks.onSheetChange("Liste");
			const cell = h.view.grid.rows[1].cells[0];
			h.setExcelOpen(true);
			await h.controller.commit(cell, "blocked");
			assert.equal(h.writes(), 0);
			assert.equal(h.view.grid.readOnly, true);
			assert.ok(h.notices.some((text) => text.includes("请先在 Excel 中关闭该文件")));
		} finally { h.controller.onunload(); }
	});
});

describe("open workbook in the local default application", () => {
	it("opens readonly workbooks without writing them", async () => {
		const h = await harness();
		try {
			assert.equal(h.view.grid.canOpenLocal, true);
			h.setOpen(async () => { assert.equal(h.view.grid.readOnly, true); assert.equal(h.view.grid.busy, true); });
			await h.controller.openLocally();
			assert.equal(h.opens(), 1); assert.equal(h.writes(), 0);
			assert.equal(h.view.grid.readOnly, true); assert.equal(h.view.grid.busy, false);
		} finally { h.controller.onunload(); }
	});
	it("commits the active editor, waits for its save and only then opens in readonly mode", async () => {
		const h = await harness();
		try {
			await h.view.callbacks.onModeChange(false); h.view.callbacks.onSheetChange("Liste");
			const cell = h.view.grid.rows[1].cells[0];
			h.view.onFinishEdit = () => { h.view.callbacks.onCommit(cell, "saved before local open"); };
			h.setOpen(async () => {
				assert.equal(h.view.grid.readOnly, true);
				assert.equal((await h.read()).sheet("Liste").cell(2, 1).value, "saved before local open");
			});
			await h.controller.openLocally();
			assert.equal(h.opens(), 1); assert.equal(h.writes(), 1);
			assert.equal(h.view.grid.readOnly, true); assert.deepEqual(h.notices, []);
		} finally { h.controller.onunload(); }
	});
	it("does not launch when the pending editor cannot be saved", async () => {
		const h = await harness();
		try {
			await h.view.callbacks.onModeChange(false); h.view.callbacks.onSheetChange("Liste");
			h.setWrite(async () => { throw new Error("disk full"); });
			const cell = h.view.grid.rows[1].cells[0];
			h.view.onFinishEdit = () => { h.view.callbacks.onCommit(cell, "not saved"); };
			await h.controller.openLocally();
			assert.equal(h.opens(), 0); assert.equal(h.writes(), 0);
			assert.equal(h.view.grid.readOnly, true); assert.equal(h.view.grid.busy, false);
			assert.ok(h.notices.some((notice) => notice.includes("有修改未保存成功")));
		} finally { h.controller.onunload(); }
	});
	it("keeps readonly and reports a default-application failure", async () => {
		const h = await harness();
		try {
			await h.view.callbacks.onModeChange(false);
			h.setOpen(async () => { throw new Error("No application associated"); });
			await h.controller.openLocally();
			assert.equal(h.view.grid.readOnly, true); assert.equal(h.view.grid.busy, false);
			assert.ok(h.notices.some((notice) => notice.includes("No application associated")));
		} finally { h.controller.onunload(); }
	});
	it("ignores duplicate clicks while a launch is pending", async () => {
		const h = await harness();
		try {
			let release!: () => void;
			h.setOpen(() => new Promise<void>((resolve) => { release = resolve; }));
			const opening = h.controller.openLocally();
			await until(() => h.opens() === 1);
			await h.controller.openLocally(); assert.equal(h.opens(), 1);
			release(); await opening;
		} finally { h.controller.onunload(); }
	});
	it("stops an in-flight handoff after the preview is unloaded", async () => {
		const h = await harness();
		const opening = h.controller.openLocally(); h.controller.onunload();
		await opening; assert.equal(h.opens(), 0);
	});
	it("hides unsupported platform capability", async () => {
		const h = await harness();
		try {
			delete h.file.openLocal; await h.controller.reload();
			assert.equal(h.view.grid.canOpenLocal, false);
			await h.controller.openLocally(); assert.equal(h.opens(), 0);
		} finally { h.controller.onunload(); }
	});
	it("switches other previews of the same workbook to readonly too", async () => {
		const h = await harness(), peer = await harness();
		try {
			peer.controller.file = h.file;
			h.active.add(peer.controller);
			await h.view.callbacks.onModeChange(false); await peer.view.callbacks.onModeChange(false);
			await h.controller.openLocally();
			assert.equal(h.view.grid.readOnly, true); assert.equal(peer.view.grid.readOnly, true);
			assert.equal(h.opens(), 1);
		} finally { h.controller.onunload(); peer.controller.onunload(); }
	});
});

describe("formatting a selection", () => {
	it("formats a range in one undoable write and leaves values untouched", async () => {
		const h = await harness();
		try {
			await h.view.callbacks.onModeChange(false); h.view.callbacks.onSheetChange("2026");
			const cells = h.view.grid.rows[1].cells.slice(2, 4);
			const before = (await h.read()).sheet("2026");
			await h.controller.formatCells(cells, { bold: true, underline: true, fill: "#FFF2CC", color: "#C00000" });
			assert.equal(h.writes(), 1);
			const after = (await h.read()).sheet("2026");
			for (const cell of cells) {
				assert.equal(after.cell(cell.row, cell.col).value, before.cell(cell.row, cell.col).value);
				assert.equal(after.cell(cell.row, cell.col).look.bold, true);
				assert.equal(after.cell(cell.row, cell.col).look.fill, "#FFF2CC");
			}
			await h.controller.history("undo");
			assert.deepEqual((await h.read()).sheet("2026").cell(2, 3), before.cell(2, 3));
			assert.equal(h.view.grid.canUndo, false);
		} finally { h.controller.onunload(); }
	});
	it("rejects formatting in readonly mode or while Excel owns the file", async () => {
		const h = await harness();
		try {
			await h.controller.formatCells(h.view.grid.rows[0].cells.slice(0, 1), { bold: true });
			assert.equal(h.writes(), 0);
			await h.view.callbacks.onModeChange(false); h.view.callbacks.onSheetChange("Liste");
			h.setExcelOpen(true);
			await h.controller.formatCells(h.view.grid.rows[1].cells, { fill: "#FFFF00" });
			assert.equal(h.writes(), 0);
			assert.equal(h.view.grid.readOnly, true);
		} finally { h.controller.onunload(); }
	});
});
