import { Menu, Notice, setIcon } from "obsidian";
import { colName } from "../model/ooxml/refs";
import type { Axis, StructureEdit } from "../model/ooxml/structure-refs";
import { bounds, clipboardText, parseClipboard, pasteChanges } from "./grid-tools";
import type { CellChange, CellPoint, CellRange } from "./grid-tools";
import type { CellStylePatch, FontToggle } from "../model/ooxml/styles";
import { FormatToolbar } from "./format-toolbar";
import { CellPopup } from "./cell-popup";
import type { CellData, CellStyle, GridData, RowData, WarningData } from "./grid-data";

const ROW_NUMBER_PX = 44;

export interface GridCallbacks {
	registerCellShortcut?: (input: HTMLTextAreaElement, insertLineBreak: () => void) => () => void;
	onFormat(cells: CellData[], patch: CellStylePatch): void;
	onBatchCommit(changes: CellChange[]): void;
	onStructure(edit: StructureEdit): void;
	onHistory(direction: "undo" | "redo"): void;
	onSheetChange(sheet: string): void;
	onModeChange(readOnly: boolean): void;
	onCommit(cell: CellData, text: string): void;
	onOpenLocal(): void;
	onRefresh(): void;
	onWarningAction(warningId: string, actionId: string): void;
}

interface ActiveEdit {
	r: number;
	c: number;
	input: HTMLTextAreaElement;
	finish(commit: boolean, move?: [number, number]): void;
}

/**
 * Draws a grid and handles cell editing. Knows nothing about files or Excel:
 * it renders GridData and reports edits through the callbacks.
 */
export class GridView {
	private titleEl?: HTMLElement;
	private sheetsEl?: HTMLElement;
	private modeEl?: HTMLButtonElement;
	private sheetKey = "";
	private readOnly = true;
	private identity = "";
	private statusEl?: HTMLElement;
	private warningsEl?: HTMLElement;
	private tableEl?: HTMLTableElement;
	private tbody?: HTMLTableSectionElement;
	private noteEl?: HTMLElement;
	private columnsKey = "";
	private rowNumbers = true;
	private colEls: HTMLTableColElement[] = [];
	private rowEls: HTMLTableRowElement[] = [];
	private data: CellData[][] = [];
	private editing: ActiveEdit | null = null;
	private pendingFocus: { r: number; c: number } | null = null;
	private statusTimer?: number;
	private grid?: GridData;
	private cells?: CellRange;
	private cellDragging = false;
	private toolsEl?: HTMLElement;
	private toolsOpen = false;
	private queryEl?: HTMLInputElement;
	private findResult?: HTMLElement;
	private formatToolbar?: FormatToolbar;
	private wrapText = false;
	private localWidths = new Map<number, number>();
	private selection?: { axis: Axis; anchor: number; end: number };
	private dragging = false;
	private resizeObserver?: ResizeObserver;
	private freezeFrame?: number;

	constructor(
		private root: HTMLElement,
		private callbacks: GridCallbacks,
	) {
		root.addClass("excidian");
		root.tabIndex = 0;
		root.addEventListener("click", (e) => {
			if (!(e.target as HTMLElement).closest(".excidian-format-color")) this.formatToolbar?.closePalettes();
			e.stopPropagation();
		});
		root.addEventListener("keydown", (e) => this.keydown(e));
		root.addEventListener("keydown", (e) => {
			if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "f") {
				e.preventDefault(); e.stopPropagation(); this.toolsOpen = true;
				this.toolsEl?.toggle(true); this.queryEl?.focus(); this.queryEl?.select();
			}
		}, true);
		root.addEventListener("copy", (e) => this.copy(e));
		root.addEventListener("paste", (e) => this.paste(e));
	}

	/** Draws the grid, updating the existing one in place so an open editor survives refreshes. */
	render(grid: GridData) {
		const identity = `${grid.tooltip}\u0000${grid.activeSheet ?? grid.title}`;
		const changed = identity !== this.identity;
		if (changed) { this.cells = undefined; this.localWidths.clear(); this.toolsOpen = false; }
		const prior = this.cells?.end && this.data[this.cells.end.r]?.[this.cells.end.c];
		if (this.editing && this.data[this.editing.r]?.[this.editing.c]?.row !== grid.rows[this.editing.r]?.cells[this.editing.c]?.row) this.cancelEdit();
		if (identity !== this.identity || grid.readOnly) {
			this.cancelEdit();
			if (identity !== this.identity) { this.selection = undefined; this.reset(); }
		}
		this.grid = grid;
		this.identity = identity;
		this.readOnly = grid.readOnly ?? false;
		const rowNumbers = grid.rowNumbers !== false;
		const columnHeader = grid.columnHeader !== false;
		const columnsKey = [rowNumbers, columnHeader, ...grid.columns].join("\u0000");
		if (!this.tbody || columnsKey !== this.columnsKey) this.build(grid.columns, rowNumbers, columnHeader);
		this.columnsKey = columnsKey;

		this.titleEl!.setText(grid.title);
		this.titleEl!.title = grid.tooltip;
		setIcon(this.modeEl!, this.readOnly ? "eye" : "pencil");
		this.modeEl!.title = this.readOnly ? "当前只读 · 点击切换为可编辑" : "当前可编辑 · 点击切换为只读";
		this.modeEl!.setAttribute("aria-label", this.readOnly ? "切换为可编辑模式" : "切换为只读模式");
		this.modeEl!.setAttribute("aria-pressed", String(!this.readOnly));
		this.modeEl!.disabled = !!grid.busy;
		const openLocal = this.root.querySelector<HTMLButtonElement>(".excidian-open-local")!;
		openLocal.toggle(!!grid.canOpenLocal);
		openLocal.disabled = !!grid.busy;
		this.root.toggleClass("excidian-is-readonly", this.readOnly);
		this.renderSheets(grid);
		this.renderWarnings(grid.warnings ?? []);
		this.sizeColumns(grid);
		const remapped = JSON.stringify(this.data.map((row) => row[0]?.row)) !== JSON.stringify(grid.rows.map((row) => row.cells[0]?.row));
		this.data = grid.rows.map((row) => row.cells);
		if (prior && remapped && !changed) {
			const r = this.data.findIndex((row) => row.some((cell) => cell.row === prior.row));
			const c = this.data[r]?.findIndex((cell) => cell.col === prior.col) ?? -1;
			this.cells = r >= 0 && c >= 0 ? { anchor: { r, c }, end: { r, c } } : undefined;
		}
		grid.rows.forEach((row, i) => this.renderRow(row, i));
		while (this.rowEls.length > grid.rows.length) {
			if (this.editing && this.editing.r === this.rowEls.length - 1) this.editing.finish(false);
			this.rowEls.pop()!.remove();
		}

		this.noteEl!.setText(grid.note ?? "");
		this.noteEl!.toggle(!!grid.note);
		this.paintSelection();
		this.root.toggleClass("excidian-wrap-all", this.wrapText);
		this.scheduleFrozen();
		this.root.setAttribute("aria-busy", String(!!grid.busy));
		for (const direction of ["undo", "redo"] as const) {
			const button = this.root.querySelector<HTMLButtonElement>(`.excidian-${direction}`);
			if (button) button.disabled = this.readOnly || !!grid.busy || !(direction === "undo" ? grid.canUndo : grid.canRedo);
		}

		const focus = this.pendingFocus;
		this.pendingFocus = null;
		if (focus && !this.editing) this.selectCell(focus);
	}

	showError(message: string) {
		this.reset();
		this.root.createDiv({ cls: "excidian-error", text: `Excidian: ${message}` });
	}

	setStatus(text: string) {
		if (!this.statusEl) return;
		this.statusEl.setText(text);
		window.clearTimeout(this.statusTimer);
		this.statusTimer = window.setTimeout(() => this.statusEl?.setText(""), 2000);
	}

	/** Commit the active editor before handing the workbook to another application. */
	finishEdit() { this.editing?.finish(true); }

	/** Stop an unfinished editor before changing sheet, mode or lifecycle. */
	cancelEdit() {
		this.editing?.finish(false);
		this.pendingFocus = null;
	}

	destroy() {
		this.reset();
		this.resizeObserver?.disconnect();
		if (this.freezeFrame !== undefined) this.root.ownerDocument.defaultView!.cancelAnimationFrame(this.freezeFrame);
		window.clearTimeout(this.statusTimer);
	}

	private reset() {
		this.editing?.finish(false);
		this.root.empty();
		this.tbody = undefined;
		this.tableEl = undefined;
		this.colEls = [];
		this.columnsKey = "";
		this.rowEls = [];
		this.editing = null;
		this.pendingFocus = null;
		this.sheetKey = "";
		this.dragging = false;
		this.cellDragging = false;
	}

	private build(columns: string[], rowNumbers: boolean, columnHeader: boolean) {
		this.reset();
		this.rowNumbers = rowNumbers;
		const bar = this.root.createDiv("excidian-toolbar");
		// Toolbar/tab controls belong to the preview, including in Live Preview.
		bar.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); });
		bar.addEventListener("click", (e) => {
			if (!(e.target as HTMLElement).closest(".excidian-format-color")) this.formatToolbar?.closePalettes();
			e.stopPropagation();
		});
		this.titleEl = bar.createSpan("excidian-title");
		this.statusEl = bar.createSpan("excidian-status");
		this.formatToolbar = new FormatToolbar(bar, (property) => this.toggleFormat(property), (patch) => this.applyFormat(patch));
		for (const direction of ["undo", "redo"] as const) {
			const button = bar.createEl("button", { cls: `clickable-icon excidian-${direction}`, attr: { "aria-label": direction === "undo" ? "撤销工作簿修改" : "重做工作簿修改" } });
			setIcon(button, direction === "undo" ? "undo-2" : "redo-2");
			button.onclick = () => { this.editing?.finish(true); this.callbacks.onHistory(direction); };
		}
		this.modeEl = bar.createEl("button", { cls: "clickable-icon excidian-mode" });
		this.modeEl.onclick = () => {
			this.editing?.finish(true);
			this.callbacks.onModeChange(!this.readOnly);
		};
		const refresh = bar.createEl("button", { cls: "clickable-icon", attr: { "aria-label": "Reload from disk" } });
		setIcon(refresh, "refresh-cw");
		refresh.onclick = () => this.callbacks.onRefresh();
		const tools = bar.createEl("button", { cls: "clickable-icon excidian-tools-toggle", attr: { "aria-label": "搜索当前工作表", title: "搜索当前工作表" } });
		setIcon(tools, "search");
		tools.onclick = () => {
			this.toolsOpen = !this.toolsOpen;
			this.toolsEl?.toggle(this.toolsOpen);
			if (this.toolsOpen) { this.queryEl?.focus(); this.queryEl?.select(); }
		};
		const wrapButton = bar.createEl("button", { cls: "clickable-icon excidian-wrap-toggle", attr: { title: "自动换行（仅当前预览）", "aria-label": "自动换行", "aria-pressed": String(this.wrapText) } });
		setIcon(wrapButton, "wrap-text");
		wrapButton.onclick = () => {
			this.wrapText = !this.wrapText; wrapButton.setAttribute("aria-pressed", String(this.wrapText));
			this.root.toggleClass("excidian-wrap-all", this.wrapText); this.scheduleFrozen();
		};
		const openLocal = bar.createEl("button", { cls: "clickable-icon excidian-open-local", attr: { "aria-label": "用本地默认程序打开", title: "用本地默认程序打开（Excel / WPS）" } });
		setIcon(openLocal, "file-spreadsheet");
		openLocal.onclick = () => this.callbacks.onOpenLocal();
		this.buildSearch();
		this.warningsEl = this.root.createDiv("excidian-warnings");

		const wrap = this.root.createDiv("excidian-wrap");
		if (!this.resizeObserver) {
			this.resizeObserver = new ResizeObserver(() => this.scheduleFrozen());
			this.resizeObserver.observe(this.root);
		}
		// Keep clicks from moving the editor cursor into the code block source.
		wrap.addEventListener("mousedown", (e) => e.stopPropagation());
		const table = wrap.createEl("table", { cls: "excidian-table" });
		this.tableEl = table;
		const colgroup = table.createEl("colgroup");
		if (rowNumbers) colgroup.createEl("col", { cls: "excidian-col-rownum" });
		this.colEls = columns.map(() => colgroup.createEl("col"));
		if (columnHeader) {
			const head = table.createEl("thead").createEl("tr");
			if (rowNumbers) head.createEl("th", { cls: "excidian-corner" });
			for (const [i, label] of columns.entries()) {
				const th = head.createEl("th", { text: label, attr: { title: label, "data-grid-column": String(i) } });
				const handle = th.createSpan({ cls: "excidian-resize", attr: { title: "拖动调整当前预览列宽；双击恢复", "aria-hidden": "true" } });
				handle.addEventListener("pointerdown", (e) => this.resizeColumn(e, i, th));
				handle.addEventListener("dblclick", (e) => { e.preventDefault(); e.stopPropagation(); this.localWidths.delete(i); this.sizeColumns(this.grid!); this.scheduleFrozen(); });
			}
		}
		table.addEventListener("pointerdown", (e) => { if (this.pointAt(e.target as HTMLElement)) this.selectCellPointer(e); else this.selectHeader(e); });
		table.addEventListener("pointerover", (e) => {
			if (this.cellDragging && e.buttons) {
				const point = this.pointAt(e.target as HTMLElement);
				if (point && this.cells) { this.cells.end = point; this.paintSelection(); }
				return;
			}
			if (!this.dragging || !e.buttons) { this.dragging = false; return; }
			const header = this.headerAt(e.target as HTMLElement);
			if (header && this.selection?.axis === header.axis) {
				this.selection.end = header.index; this.paintSelection();
			}
		});
		table.addEventListener("pointerup", () => { this.dragging = false; this.cellDragging = false; });
		table.addEventListener("contextmenu", (e) => {
			const point = this.pointAt(e.target as HTMLElement);
			if (!point) { this.headerMenu(e); return; }
			e.preventDefault(); e.stopPropagation();
			const area = this.cells && bounds(this.cells);
			if (!area || point.r < area.top || point.r > area.bottom || point.c < area.left || point.c > area.right) this.selectCell(point);
			this.editing?.finish(true);
			const menu = new Menu();
			menu.addItem((item) => item.setTitle("复制选区").setIcon("copy").onClick(() => {
				const selected = bounds(this.cells!);
				if ((selected.bottom - selected.top + 1) * (selected.right - selected.left + 1) > 10_000) { this.report("每次最多复制 10,000 个单元格。"); return; }
				const text = clipboardText(this.data.slice(selected.top, selected.bottom + 1).map((row) => row.slice(selected.left, selected.right + 1).map((cell) => cell.covered ? "" : cell.pending ? cell.editText : cell.text)));
				void this.root.ownerDocument.defaultView!.navigator.clipboard.writeText(text).catch((error) => this.report(error));
			}));
			const link = this.data[point.r]?.[point.c]?.link;
			if (link) menu.addItem((item) => item.setTitle("打开链接").setIcon("external-link").onClick(() => { this.root.ownerDocument.defaultView!.open(link, "_blank", "noopener,noreferrer"); }));
			if (!this.readOnly) menu.addItem((item) => item.setTitle("清除选区内容").setIcon("eraser").setDisabled(!!this.grid?.busy).onClick(() => this.applyPaste("")));
			menu.showAtMouseEvent(e);
		});

		this.tbody = table.createEl("tbody");
		// Columns share the page width, so long values are cut off: show them in full on hover.
		this.tbody.addEventListener("mouseover", (e) => {
			const td = (e.target as HTMLElement).closest("td");
			if (td && !td.hasAttribute("title") && td.scrollWidth > td.clientWidth) td.title = td.textContent ?? "";
		});
		// List-validation cells open their choices on a normal click.
		this.tbody.addEventListener("click", (e) => {
			const point = this.pointAt(e.target as HTMLElement);
			if (!point || e.shiftKey || this.editing || !this.cells) return;
			const { anchor, end } = this.cells;
			if (anchor.r === end.r && anchor.c === end.c && end.r === point.r && end.c === point.c && this.data[point.r]?.[point.c]?.options?.length) this.startEdit(point.r, point.c);
		});
		this.tbody.addEventListener("dblclick", (e) => {
			if ((e.target as HTMLElement).closest("textarea")) return;
			const point = this.pointAt(e.target as HTMLElement);
			if (point) { e.preventDefault(); this.startEdit(point.r, point.c); }
		});
		this.noteEl = this.root.createDiv("excidian-note");
		this.sheetsEl = this.root.createDiv({ cls: "excidian-sheets", attr: { role: "tablist", "aria-label": "工作表" } });
		this.sheetsEl.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); });
		this.sheetsEl.addEventListener("click", (e) => e.stopPropagation());
	}

	private report(error: unknown) { new Notice(`Excidian: ${error instanceof Error ? error.message : String(error)}`, 7000); }

	private pointAt(target: HTMLElement): CellPoint | undefined {
		const td = target.closest("td");
		if (!td || td.closest("table") !== this.tableEl) return;
		const r = this.rowEls.indexOf(td.parentElement as HTMLTableRowElement);
		const c = Array.prototype.indexOf.call(td.parentElement!.children, td) - this.offset;
		return r >= 0 && c >= 0 && !this.data[r]?.[c]?.covered ? { r, c } : undefined;
	}

	private selectCell(point: CellPoint, extend = false, focus = true, scroll = focus) {
		if (!this.data[point.r]?.[point.c] || this.data[point.r][point.c].covered) return;
		this.selection = undefined;
		this.cells = extend && this.cells ? { anchor: this.cells.anchor, end: point } : { anchor: point, end: point };
		this.paintSelection();
		if (focus) this.root.focus({ preventScroll: true });
		if (scroll) {
			const td = this.cellEl(point.r, point.c);
			const wrap = this.tableEl?.parentElement;
			// Scroll only inside this preview, without moving the note's editor.
			if (td && wrap && !td.hasClass("excidian-frozen-row")) {
				const box = td.getBoundingClientRect(), viewport = wrap.getBoundingClientRect();
				const frozenHeight = Array.from(this.tableEl?.querySelectorAll("tr") ?? []).filter((tr) => tr.firstElementChild?.classList.contains("excidian-frozen-row")).reduce((n, tr) => n + tr.getBoundingClientRect().height, this.tableEl?.tHead?.getBoundingClientRect().height ?? 0);
				if (box.top < viewport.top + frozenHeight) wrap.scrollTop -= viewport.top + frozenHeight - box.top;
				if (box.bottom > viewport.bottom) wrap.scrollTop += box.bottom - viewport.bottom;
			}
			if (td && wrap && !td.hasClass("excidian-frozen-column")) {
				const box = td.getBoundingClientRect(), viewport = wrap.getBoundingClientRect();
				const frozenWidth = Array.from(this.rowEls[point.r]?.children ?? []).filter((el) => el.classList.contains("excidian-frozen-column")).reduce((n, el) => n + el.getBoundingClientRect().width, 0);
				if (box.left < viewport.left + frozenWidth) wrap.scrollLeft -= viewport.left + frozenWidth - box.left;
				if (box.right > viewport.right) wrap.scrollLeft += box.right - viewport.right;
			}
		}
	}

	private selectCellPointer(event: PointerEvent) {
		if (event.button !== 0 || this.grid?.busy || (event.target as HTMLElement).closest("textarea, input")) return;
		const point = this.pointAt(event.target as HTMLElement); if (!point) return;
		event.preventDefault(); event.stopPropagation(); this.editing?.finish(true);
		this.selectCell(point, event.shiftKey); this.cellDragging = true;
	}

	private keydown(event: KeyboardEvent) {
		const target = event.target as HTMLElement;
		if (target.closest("input, select, textarea, button")) return;
		const command = event.ctrlKey || event.metaKey;
		const key = event.key.toLowerCase();
		if (event.key === "Escape") {
			this.cells = undefined; this.selection = undefined; this.paintSelection(); this.cancelEdit(); event.stopPropagation(); return;
		}
		if (!this.grid || this.grid.busy) { event.preventDefault(); event.stopPropagation(); return; }
		// Keep grid shortcuts from changing the surrounding Markdown document.
		event.stopPropagation();
		if (command && (key === "z" || key === "y")) {
			event.preventDefault();
			if (!this.readOnly) this.callbacks.onHistory(event.shiftKey || key === "y" ? "redo" : "undo"); return;
		}
		if (command && key === "x") { event.preventDefault(); this.report("请使用复制、粘贴和清除内容；暂不提供剪切移动。"); return; }
		const formatKey = ({ b: "bold", i: "italic", u: "underline" } as const)[key as "b" | "i" | "u"];
		if (command && formatKey) {
			event.preventDefault(); this.toggleFormat(formatKey); return;
		}
		if (command && key === "a") {
			event.preventDefault(); event.stopPropagation();
			if (this.data.length && this.data[0].length) this.cells = { anchor: { r: 0, c: 0 }, end: { r: this.data.length - 1, c: this.data[0].length - 1 } };
			this.selection = undefined; this.paintSelection(); return;
		}
		if (["Delete", "Backspace"].includes(event.key) && !command && !event.altKey) {
			event.preventDefault(); event.stopPropagation();
			if (this.readOnly) return;
			if (this.selection && this.grid.structureEditable) {
				const { axis, anchor, end } = this.selection;
				this.callbacks.onStructure({ axis, kind: "clear", start: Math.min(anchor, end), count: Math.abs(end - anchor) + 1 });
			} else if (this.cells) this.applyPaste("");
			return;
		}
		if (!this.data.length || !this.data[0].length) return;
		const point = this.cells?.end ?? { r: 0, c: 0 };
		if (event.altKey && event.key === "ArrowDown" && this.data[point.r]?.[point.c]?.options?.length) {
			event.preventDefault(); this.startEdit(point.r, point.c); return;
		}
		const steps: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Tab: [0, event.shiftKey ? -1 : 1] };
		const step = steps[event.key];
		if (step || event.key === "Home" || event.key === "End") {
			event.preventDefault(); event.stopPropagation();
			let r = point.r, c = point.c;
			if (step) { r += step[0]; c += step[1]; }
			else { c = event.key === "Home" ? 0 : this.data[r].length - 1; if (command) r = event.key === "Home" ? 0 : this.data.length - 1; }
			while (step && this.data[r]?.[c]?.covered) { r += step[0]; c += step[1]; }
			this.selectCell({ r, c }, event.shiftKey && event.key !== "Tab"); return;
		}
		if (event.key === "Enter" || event.key === "F2") {
			event.preventDefault(); event.stopPropagation();
			if (this.readOnly) this.selectCell({ r: point.r + (event.shiftKey ? -1 : 1), c: point.c });
			else this.startEdit(point.r, point.c); return;
		}
		if (event.keyCode === 229 && !this.readOnly) { this.startEdit(point.r, point.c, ""); return; }
		if (event.key.length === 1 && !command && !event.altKey && !event.isComposing) {
			event.preventDefault(); event.stopPropagation();
			if (!this.readOnly) this.startEdit(point.r, point.c, event.key);
		}
	}

	private copy(event: ClipboardEvent) {
		if ((event.target as HTMLElement).closest("input, textarea")) return;
		if (!this.cells || !event.clipboardData) return;
		const area = bounds(this.cells), rows: string[][] = [];
		if ((area.bottom - area.top + 1) * (area.right - area.left + 1) > 10_000) { event.preventDefault(); this.report("每次最多复制 10,000 个单元格。"); return; }
		for (let r = area.top; r <= area.bottom; r++) rows.push(this.data[r].slice(area.left, area.right + 1).map((cell) => cell.covered ? "" : cell.pending ? cell.editText : cell.text));
		event.preventDefault(); event.stopPropagation(); event.clipboardData.setData("text/plain", clipboardText(rows));
		this.setStatus(`已复制 ${rows.length} × ${rows[0].length}`);
	}

	private paste(event: ClipboardEvent) {
		if ((event.target as HTMLElement).closest("input, textarea")) return;
		if (!this.cells) return;
		event.preventDefault(); event.stopPropagation();
		if (this.readOnly) { this.report("请先切换为可编辑模式。"); return; }
		this.applyPaste(event.clipboardData?.getData("text/plain") ?? "");
	}

	private applyPaste(text: string) {
		if (!this.grid || !this.cells || this.readOnly || this.grid.busy) return;
		try {
			const area = bounds(this.cells);
			const changes = pasteChanges(this.grid, { r: area.top, c: area.left }, parseClipboard(text), this.cells);
			this.callbacks.onBatchCommit(changes);
		} catch (error) { this.report(error); }
	}


	private buildSearch() {
		const tools = this.root.createDiv("excidian-tools"); this.toolsEl = tools; tools.toggle(this.toolsOpen);
		tools.addEventListener("mousedown", (e) => e.stopPropagation());
		tools.addEventListener("keydown", (e) => e.stopPropagation());
		const search = tools.createDiv("excidian-tool-row");
		this.queryEl = search.createEl("input", { type: "search", placeholder: "搜索当前工作表", attr: { "aria-label": "搜索当前工作表" } });
		this.queryEl.addEventListener("input", (event) => {
			if (!(event as InputEvent).isComposing) this.find(false, true);
		});
		this.queryEl.addEventListener("compositionend", () => this.find(false, true));
		this.queryEl.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); this.find(e.shiftKey); } });
		for (const previous of [true, false]) {
			const button = search.createEl("button", { cls: "clickable-icon", attr: { title: previous ? "上一个结果" : "下一个结果", "aria-label": previous ? "上一个结果" : "下一个结果" } });
			setIcon(button, previous ? "chevron-up" : "chevron-down");
			button.onclick = () => this.find(previous);
		}
		this.findResult = search.createSpan("excidian-find-result");
	}

	private formatCells(): CellData[] {
		if (!this.cells || this.readOnly || this.grid?.busy || this.editing) return [];
		const area = bounds(this.cells);
		if ((area.bottom - area.top + 1) * (area.right - area.left + 1) > 10_000) return [];
		const cells = this.data.slice(area.top, area.bottom + 1).flatMap((row) => row.slice(area.left, area.right + 1));
		return cells.some((cell) => cell.covered || cell.span) ? [] : cells;
	}

	private refreshFormatting() { this.formatToolbar?.update(!this.readOnly, this.formatCells()); }

	private toggleFormat(property: FontToggle) {
		const cells = this.formatCells();
		if (cells.length) this.callbacks.onFormat(cells, { [property]: !cells.every((cell) => cell.style?.[property]) });
	}

	private applyFormat(patch: CellStylePatch) {
		const cells = this.formatCells();
		if (cells.length) this.callbacks.onFormat(cells, patch);
	}

	private find(previous = false, first = false) {
		const query = this.queryEl?.value.toLocaleLowerCase() ?? "";
		if (!query) { this.findResult?.setText(""); return; }
		const matches: CellPoint[] = [];
		this.data.forEach((row, r) => row.forEach((cell, c) => { if (!cell.covered && cell.text.toLocaleLowerCase().includes(query)) matches.push({ r, c }); }));
		const point = this.cells?.end;
		const current = matches.findIndex((match) => match.r === point?.r && match.c === point?.c);
		const i = first || current < 0 ? (previous ? matches.length - 1 : 0) : (current + (previous ? -1 : 1) + matches.length) % matches.length;
		this.findResult?.setText(matches.length ? `${i + 1} / ${matches.length}` : "0 个结果");
		if (matches[i]) {
			const focused = this.root.ownerDocument.activeElement;
			this.selectCell(matches[i], false, focused !== this.queryEl, true);
		}
	}

	private resizeColumn(event: PointerEvent, column: number, th: HTMLElement) {
		event.preventDefault(); event.stopPropagation(); this.editing?.finish(true);
		const handle = event.target as HTMLElement, start = event.clientX, width = th.getBoundingClientRect().width;
		handle.setPointerCapture(event.pointerId);
		const move = (e: PointerEvent) => { this.localWidths.set(column, Math.max(40, Math.min(800, width + e.clientX - start))); this.sizeColumns(this.grid!); this.scheduleFrozen(); };
		const finish = () => { handle.removeEventListener("pointermove", move); handle.removeEventListener("pointerup", finish); handle.removeEventListener("lostpointercapture", finish); };
		handle.addEventListener("pointermove", move); handle.addEventListener("pointerup", finish); handle.addEventListener("lostpointercapture", finish);
	}


	private renderSheets(grid: GridData) {
		const key = JSON.stringify([grid.sheets, grid.activeSheet]);
		if (key === this.sheetKey) return;
		this.sheetKey = key;
		this.sheetsEl!.empty();
		this.sheetsEl!.toggle(!!grid.sheets?.length);
		for (const name of grid.sheets ?? []) {
			const active = name === grid.activeSheet;
			const tab = this.sheetsEl!.createEl("button", {
				text: name,
				cls: "excidian-sheet-tab",
				attr: { role: "tab", "aria-selected": String(active), "aria-label": `工作表 ${name}`, tabindex: active ? "0" : "-1" },
			});
			tab.onclick = () => {
				if (active) return;
				this.editing?.finish(true);
				this.callbacks.onSheetChange(name);
			};
			tab.addEventListener("keydown", (e) => {
				const names = grid.sheets!;
				const i = names.indexOf(name);
				const next = e.key === "ArrowRight" ? (i + 1) % names.length
					: e.key === "ArrowLeft" ? (i + names.length - 1) % names.length
					: e.key === "Home" ? 0 : e.key === "End" ? names.length - 1 : -1;
				if (next < 0) return;
				e.preventDefault();
				e.stopPropagation();
				this.editing?.finish(true);
				this.callbacks.onSheetChange(names[next]);
				(this.sheetsEl?.children[next] as HTMLElement | undefined)?.focus();
			});
		}
	}

	private headerAt(target: HTMLElement): { axis: Axis; index: number } | undefined {
		const th = target.closest("th");
		if (!th || th.closest("table") !== this.tableEl) return;
		const col = th.getAttribute("data-grid-column");
		if (col !== null && this.grid?.columnNumbers?.[Number(col)]) return { axis: "column", index: this.grid.columnNumbers[Number(col)] };
		if (th.hasClass("excidian-rownum")) {
			const row = this.rowEls.indexOf(th.parentElement as HTMLTableRowElement);
			const index = Number(this.grid?.rows[row]?.label);
			if (index > 0) return { axis: "row", index };
		}
	}

	private selectHeader(event: PointerEvent) {
		if (event.button !== 0 || !this.grid?.structureEditable || this.grid.busy) return;
		const header = this.headerAt(event.target as HTMLElement); if (!header) return;
		event.preventDefault(); event.stopPropagation();
		this.editing?.finish(true);
		if (event.shiftKey && this.selection?.axis === header.axis) this.selection.end = header.index;
		else this.selection = { axis: header.axis, anchor: header.index, end: header.index };
		this.cells = undefined; this.dragging = true; this.root.focus({ preventScroll: true }); this.paintSelection();
	}

	private headerMenu(event: MouseEvent) {
		const header = this.headerAt(event.target as HTMLElement); if (!header) return;
		// Consume the event even in read-only mode so Obsidian never treats a
		// worksheet header as a Markdown editor context-menu request.
		event.preventDefault(); event.stopPropagation();
		if (!this.grid?.structureEditable || this.grid.busy) return;
		this.editing?.finish(true);
		const selection = this.selection;
		if (!selection || selection.axis !== header.axis || header.index < Math.min(selection.anchor, selection.end) || header.index > Math.max(selection.anchor, selection.end)) {
			this.selection = { axis: header.axis, anchor: header.index, end: header.index };
		}
		this.cells = undefined; this.root.focus({ preventScroll: true }); this.paintSelection();
		const { axis, anchor, end } = this.selection!;
		const start = Math.min(anchor, end), count = Math.abs(end - anchor) + 1;
		const label = axis === "row" ? "行" : "列";
		const from = axis === "row" ? String(start) : colName(start);
		const to = axis === "row" ? String(start + count - 1) : colName(start + count - 1);
		const menu = new Menu();
		menu.addItem((item) => item.setTitle(`工作表整${label}操作 · ${from}${count > 1 ? `–${to}` : ""}`).setDisabled(true));
		menu.addSeparator();
		menu.addItem((item) => item.setTitle(`插入 ${count} ${label}（${axis === "row" ? "上方" : "左侧"}）`).setIcon("plus").onClick(() => this.callbacks.onStructure({ axis, kind: "insert", start, count })));
		menu.addItem((item) => item.setTitle(`删除 ${count} ${label}`).setIcon("trash-2").onClick(() => this.callbacks.onStructure({ axis, kind: "delete", start, count })));
		menu.addSeparator();
		menu.addItem((item) => item.setTitle("清除内容（保留格式和表格标题）").setIcon("eraser").onClick(() => this.callbacks.onStructure({ axis, kind: "clear", start, count })));
		menu.showAtMouseEvent(event);
	}

	private paintSelection() {
		const s = this.grid?.structureEditable ? this.selection : undefined;
		const area = this.cells && bounds(this.cells);
		const contains = (n: number) => !!s && n >= Math.min(s.anchor, s.end) && n <= Math.max(s.anchor, s.end);
		for (const th of Array.from(this.tableEl?.querySelectorAll<HTMLElement>("thead th[data-grid-column]") ?? [])) {
			th.toggleClass("excidian-selected", s?.axis === "column" && contains(this.grid!.columnNumbers![Number(th.dataset.gridColumn)]));
		}
		this.rowEls.forEach((tr, i) => {
			const rowSelected = s?.axis === "row" && contains(Number(this.grid?.rows[i]?.label));
			Array.from(tr.children).forEach((cell, j) => {
				const col = this.grid?.columnNumbers?.[j - this.offset];
				const c = j - this.offset;
				const inside = !!area && i >= area.top && i <= area.bottom && c >= area.left && c <= area.right;
				(cell as HTMLElement).toggleClass("excidian-selected", inside || rowSelected || (s?.axis === "column" && col !== undefined && contains(col)));
				(cell as HTMLElement).toggleClass("excidian-active-cell", !!this.cells && i === this.cells.end.r && c === this.cells.end.c);
			});
		});
		this.refreshFormatting();
	}

	private scheduleFrozen() {
		const win = this.root.ownerDocument.defaultView!;
		if (this.freezeFrame !== undefined) win.cancelAnimationFrame(this.freezeFrame);
		this.freezeFrame = win.requestAnimationFrame(() => { this.freezeFrame = undefined; this.positionFrozen(); });
	}

	private positionFrozen() {
		if (!this.tableEl || !this.grid) return;
		const frozen = this.grid.frozen ?? { rows: 0, columns: 0 };
		const headerHeight = this.tableEl.tHead?.getBoundingClientRect().height ?? 0;
		let top = headerHeight;
		const lefts: number[] = [];
		let left = this.rowNumbers ? this.tableEl.querySelector("th")?.getBoundingClientRect().width ?? ROW_NUMBER_PX : 0;
		for (let i = 0; i < this.colEls.length; i++) {
			lefts.push(left);
			if ((this.grid.columnNumbers?.[i] ?? Infinity) <= frozen.columns) left += this.colEls[i].getBoundingClientRect().width;
		}
		const position = (el: HTMLElement, freezeRow: boolean, freezeCol: boolean, topPx: number, leftPx: number, header = false) => {
			el.toggleClass("excidian-frozen-row", freezeRow);
			el.toggleClass("excidian-frozen-column", freezeCol);
			el.style.top = freezeRow ? `${topPx}px` : "";
			el.style.left = freezeCol ? `${leftPx}px` : "";
			el.style.zIndex = freezeRow && freezeCol ? "8" : header && freezeCol ? "9" : freezeRow ? "5" : freezeCol ? "4" : "";
		};
		Array.from(this.tableEl.tHead?.rows[0]?.children ?? []).forEach((el, j) => {
			const i = j - this.offset;
			if (i >= 0) position(el as HTMLElement, false, (this.grid!.columnNumbers?.[i] ?? Infinity) <= frozen.columns, 0, lefts[i], true);
		});
		this.rowEls.forEach((tr, i) => {
			const freezeRow = Number(this.grid!.rows[i]?.label) <= frozen.rows;
			Array.from(tr.children).forEach((el, j) => {
				const c = j - this.offset;
				position(el as HTMLElement, freezeRow, c < 0 || (this.grid!.columnNumbers?.[c] ?? Infinity) <= frozen.columns, top, c < 0 ? 0 : lefts[c]);
			});
			if (freezeRow) top += tr.getBoundingClientRect().height;
		});
	}

	/**
	 * Sizes the columns from the given widths (Excel's), or else from how long their
	 * contents are. A fitted grid splits the page width in those proportions; an
	 * unfitted one gives each column its real width and scrolls sideways if needed.
	 */
	private sizeColumns(grid: GridData) {
		const widths =
			grid.columnWidths ??
			grid.columns.map((label, j) => {
				let longest = label.length;
				for (const row of grid.rows.slice(0, 200)) longest = Math.max(longest, row.cells[j]?.text.length ?? 0);
				return Math.min(Math.max(longest, 4), 40);
			});

		if (grid.fit !== false && !this.localWidths.size) {
			const total = widths.reduce((a, b) => a + b, 0);
			this.tableEl!.setCssStyles({ width: "100%" });
			widths.forEach((w, j) => this.colEls[j]?.setCssStyles({ width: `${((w / total) * 100).toFixed(2)}%` }));
			return;
		}

		// Excel's pixel width of a column for a width in characters (Calibri 11: 7px per character + padding).
		const pixels = widths.map((w, i) => this.localWidths.get(i) ?? Math.round(w * 7 + 5));
		pixels.forEach((px, j) => this.colEls[j]?.setCssStyles({ width: `${px}px` }));
		const rowNumberWidth = this.rowNumbers ? ROW_NUMBER_PX : 0;
		this.tableEl!.setCssStyles({ width: `${pixels.reduce((a, b) => a + b, rowNumberWidth)}px` });
	}

	private renderWarnings(warnings: WarningData[]) {
		const el = this.warningsEl!;
		el.empty();
		for (const warning of warnings) {
			const box = el.createDiv("excidian-warning");
			setIcon(box.createSpan("excidian-warning-icon"), "alert-triangle");
			box.createSpan({ cls: "excidian-warning-text", text: warning.text });
			for (const action of warning.actions) {
				const button = box.createEl("button", { text: action.label });
				button.onclick = () => this.callbacks.onWarningAction(warning.id, action.id);
			}
		}
	}

	private renderRow(row: RowData, i: number) {
		let tr = this.rowEls[i];
		if (!tr) {
			tr = this.tbody!.createEl("tr");
			if (this.rowNumbers) tr.createEl("th", { cls: "excidian-rownum" });
			row.cells.forEach(() => tr.createEl("td"));
			this.rowEls.push(tr);
		}
		tr.className = `excidian-row-${row.kind}`;
		if (this.rowNumbers) (tr.firstElementChild as HTMLElement).setText(row.label);
		row.cells.forEach((cell, j) => {
			const td = tr.children[j + this.offset] as HTMLTableCellElement | undefined;
			if (td) this.paint(td, cell);
		});
	}

	/** Children of a row before its first cell (the row number). */
	private get offset(): number {
		return this.rowNumbers ? 1 : 0;
	}

	private cellEl(r: number, c: number): HTMLTableCellElement | undefined {
		return this.rowEls[r]?.children[c + this.offset] as HTMLTableCellElement | undefined;
	}

	private paint(td: HTMLTableCellElement, cell: CellData) {
		if (this.editing && this.cellEl(this.editing.r, this.editing.c) === td) return;
		td.className = "";
		td.removeAttribute("style");
		// Every grid position keeps a <td> (so indexes stay simple); merged-over ones are hidden.
		td.colSpan = cell.span?.cols ?? 1;
		td.rowSpan = cell.span?.rows ?? 1;
		if (cell.covered) {
			td.empty();
			td.addClass("excidian-covered");
			return;
		}
		td.setText(cell.pending ? cell.editText : cell.text);
		td.toggleClass("excidian-link", !!cell.link);
		td.toggleClass("excidian-num", !!cell.numeric);
		td.toggleClass("excidian-formula", !!cell.formula);
		td.toggleClass("excidian-pending", !!cell.pending);
		td.toggleClass("excidian-readonly", !cell.editable);
		td.toggleClass("excidian-has-list", !!cell.options?.length && cell.editable);
		td.toggleClass("excidian-table-header", !!cell.tableHeader);
		td.toggleClass("excidian-wrap-text", !!cell.style?.wrap);
		td.toggleClass("excidian-multiline", /[\r\n]/.test(cell.text));
		if (cell.style) applyStyle(td, cell.style);
		td.title = [cell.text, cell.hint, cell.link ? `${cell.link} · 右键打开链接` : undefined].filter(Boolean).join("\n");
	}

	private startEdit(r: number, c: number, replacement?: string) {
		if (this.readOnly || this.grid?.busy) return;
		if (this.editing?.r === r && this.editing.c === c) return;
		this.editing?.finish(true);

		const cell = this.data[r]?.[c];
		const td = this.cellEl(r, c);
		if (!cell || !td || !cell.editable || cell.covered) return;
		this.selectCell({ r, c }, false, false);

		const original = cell.editText;
		td.empty();
		td.addClass("excidian-editing");
		const input = td.createEl("textarea", { attr: { rows: "1", "aria-label": `编辑 ${colName(cell.col)}${cell.row}` } });
		input.value = original;
		const resize = () => {
			input.style.height = "0px";
			input.style.height = `${Math.min(160, Math.max(26, input.scrollHeight))}px`;
			this.scheduleFrozen();
		};
		input.addEventListener("input", resize);
		const popup =
			cell.options?.length || cell.help
				? new CellPopup(td, input, cell.options ?? [], cell.help, (value) => {
						input.value = value;
						finish(true);
					})
				: null;

		let releaseShortcut: (() => void) | undefined;
		const insertLineBreak = () => {
			input.setRangeText("\n", input.selectionStart, input.selectionEnd, "end");
			input.dispatchEvent(new Event("input", { bubbles: true }));
		};
		const finish = (commit: boolean, move?: [number, number]) => {
			if (this.editing?.input !== input) return;
			this.editing = null;
			releaseShortcut?.();
			this.refreshFormatting();
			popup?.destroy();
			const text = input.value;
			// The grid may have been refreshed while editing; use the latest data for this cell.
			const current = this.data[r]?.[c] ?? cell;
			td.empty();
			if (commit && text !== original) {
				td.className = "excidian-saving";
				td.setText(text);
				// Keep the value seen when editing began for optimistic conflict checks.
				this.callbacks.onCommit(cell, text);
			} else {
				this.paint(td, current);
			}

			if (move) {
				let [nr, nc] = [r + move[0], c + move[1]];
				// Step over cells hidden under a merged cell.
				while (this.data[nr]?.[nc]?.covered) [nr, nc] = [nr + move[0], nc + move[1]];
				if (this.cellEl(nr, nc)) this.selectCell({ r: nr, c: nc });
				// The next row may appear once the save finishes (e.g. a new table row).
				else if (move[0] > 0) this.pendingFocus = { r: nr, c: nc };
			}
		};
		this.editing = { r, c, input, finish };
		this.refreshFormatting();

		// Enter/Tab take the highlighted choice of the dropdown, if there is one.
		const takeChoice = () => {
			const choice = popup?.activeValue;
			if (choice !== undefined) input.value = choice;
		};

		input.addEventListener("keydown", (e) => {
			e.stopPropagation();
			if (e.isComposing || e.keyCode === 229) return;
			if (e.key === "Enter" && e.altKey) {
				e.preventDefault();
				insertLineBreak();
				return;
			}
			if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z" && input.value === original) {
				e.preventDefault(); finish(false); this.callbacks.onHistory(e.shiftKey ? "redo" : "undo"); return;
			}
			if (popup?.hasOptions && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
				e.preventDefault();
				popup.move(e.key === "ArrowDown" ? 1 : -1);
			} else if (e.key === "Enter") {
				e.preventDefault();
				takeChoice();
				finish(true, [e.shiftKey ? -1 : 1, 0]);
			} else if (e.key === "Tab") {
				e.preventDefault();
				takeChoice();
				finish(true, [0, e.shiftKey ? -1 : 1]);
			} else if (e.key === "Escape") {
				e.preventDefault();
				finish(false);
			}
		});
		input.addEventListener("blur", () => finish(true));
		input.focus();
		if (input.ownerDocument.activeElement === input) releaseShortcut = this.callbacks.registerCellShortcut?.(input, insertLineBreak);
		if (replacement !== undefined) { input.value = replacement; input.setSelectionRange(replacement.length, replacement.length); }
		else input.select();
		resize();
	}
}

function applyStyle(td: HTMLElement, style: CellStyle) {
	td.toggleClass("excidian-bold", !!style.bold);
	td.toggleClass("excidian-italic", !!style.italic);
	td.toggleClass("excidian-underline", !!style.underline);
	td.toggleClass("excidian-strike", !!style.strike);
	if (style.align) td.addClass(`excidian-align-${style.align}`);
	// Colours come from the workbook, so they can only be set per cell.
	if (style.color) td.setCssStyles({ color: style.color });
	if (style.background) td.setCssStyles({ backgroundColor: style.background });
}
