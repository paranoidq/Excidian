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

// Exercise the real controller, workbook and save queue with only the Obsidian
// lifecycle and DOM view replaced. This catches edits redirected by tab changes.
async function harness(source = "file: demo.xlsx", defaultReadOnly = true) {
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
					cancelEdit() {} destroy() { this.destroyed = true; } setStatus() {}
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
	let watchers = 0;
	const file: StoredFile = {
		id: `test-${Math.random()}`, name: "demo.xlsx", displayPath: "demo.xlsx",
		read: async () => bytes,
		write: async (value) => { bytes = value; writes++; },
		writeSibling: async () => {}, siblings: async () => [],
		watch: () => { watchers++; return () => { watchers--; }; },
	};
	const data = new Map<string, unknown>();
	const settings = new DeviceSettings({ load: (key) => data.get(key), save: (key, value) => { data.set(key, value); } });
	settings.defaultReadOnly = defaultReadOnly;
	const active = new Set();
	const root = {} as { view: { callbacks: GridCallbacks; grid: GridData; renders: number; error?: string; destroyed?: boolean } };
	const controller = new module.exports.SheetController(root, source, "note.md", { settings, resolver: { resolve: () => file }, active });
	controller.onload();
	await until(() => !!root.view.grid || !!root.view.error);
	assert.equal(root.view.error, undefined);
	return { controller, view: root.view, settings, active, notices, read: () => Workbook.load(bytes), writes: () => writes, watchers: () => watchers };
}

async function until(check: () => boolean) {
	for (let i = 0; i < 100; i++) {
		if (check()) return;
		await new Promise((resolve) => setTimeout(resolve, 10));
	}
	throw new Error("Timed out waiting for preview");
}

describe("preview controller", () => {
	it("enforces read-only at the save boundary, with independent per-preview toggles", async () => {
		const h = await harness();
		try {
			const cell = { ...h.view.grid.rows[0].cells[0], editable: true };
			await h.controller.commit(cell, "blocked");
			assert.equal(h.writes(), 0);
			h.view.callbacks.onModeChange(false);
			assert.equal(h.view.grid.readOnly, false);
			assert.equal(h.settings.defaultReadOnly, true);
			h.view.callbacks.onModeChange(true);
			assert.ok(h.view.grid.rows.every((r) => r.cells.every((c) => !c.editable)));
		} finally { h.controller.onunload(); }
		assert.equal(h.watchers(), 0);
		assert.equal(h.active.size, 0);
		assert.equal(h.view.destroyed, true);
	});

	it("saves to the selected sheet even when another tab is selected before the save finishes", async () => {
		const h = await harness();
		try {
			h.view.callbacks.onModeChange(false);
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
