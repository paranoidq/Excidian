import { MarkdownRenderChild, Notice } from "obsidian";
import type { DeviceSettings } from "../model/device-settings";
import type { CellStylePatch } from "../model/ooxml/styles";
import { formatRef, inRange } from "../model/ooxml/refs";
import type { StructureEdit } from "../model/ooxml/structure-refs";
import type { FileResolver, SiblingFile, StoredFile } from "../model/storage";
import type { ValidationIssue } from "../model/validation";
import { CellInput, editText, parseInput } from "../model/values";
import type { Sheet } from "../model/sheet";
import type { TableLocation, Workbook } from "../model/workbook";
import { findConflictCopies, isOpenInExcel, readWorkbook, restoreWorkbook, updateWorkbook, workbookHistory } from "../model/workbook-store";
import type { CellChange } from "../view/grid-tools";
import type { CellData, GridData, WarningData } from "../view/grid-data";
import { GridView } from "../view/grid-view";
import { BlockConfig, parseBlockConfig } from "./block-config";
import { buildGrid } from "./grid-builder";

const EXCEL_OPEN_MESSAGE = "该文件已在 Excel 中打开，Excidian 已保持只读。请先在 Excel 中关闭该文件，再切换到编辑模式。";

/** Services the plugin shares with every code block. */
export interface ControllerContext {
	settings: DeviceSettings;
	/** Turns `file:` values into files: anywhere on desktop, inside the vault on mobile. */
	resolver: FileResolver;
	/** Controllers currently on screen, so settings changes can restart them. */
	active: Set<SheetController>;
	registerCellShortcut?: (input: HTMLTextAreaElement, insertLineBreak: () => void) => () => void;
}

/** Connects one excidian code block to its workbook: loads it, shows it, saves edits. */
export class SheetController extends MarkdownRenderChild {
	private view: GridView;
	private cfg?: BlockConfig;
	private file?: StoredFile;
	private grid?: GridData;
	private conflicts: SiblingFile[] = [];
	private stopWatching?: () => void;
	private reloadTimer?: number;
	private ignoreChangesUntil = 0;
	private workbook?: Workbook;
	private selectedSheet?: string;
	private readOnly = true;
	private initialEditRequested = false;
	private checkingMode = false;
	private modeGeneration = 0;
	private generation = 0;
	private stopped = true;
	private busy = false;
	private saves = new Set<Promise<void>>();
	private saveFailures = 0;

	constructor(
		containerEl: HTMLElement,
		private source: string | BlockConfig,
		/** Vault-relative path of the note containing the block. */
		private notePath: string,
		private context: ControllerContext,
	) {
		super(containerEl);
		this.view = new GridView(containerEl, {
			onSheetChange: (sheet) => {
				this.selectedSheet = sheet;
				this.render();
			},
			onModeChange: (readOnly) => this.changeMode(readOnly),
			onCommit: (cell, text) => {
				const save = this.commit(cell, text);
				this.saves.add(save);
				void save.finally(() => this.saves.delete(save));
			},
			registerCellShortcut: this.context.registerCellShortcut,
			onFormat: (cells, patch) => void this.formatCells(cells, patch),
			onBatchCommit: (changes) => void this.batchCommit(changes),
			onStructure: (edit) => void this.structure(edit),
			onHistory: (direction) => void this.history(direction),
			onOpenLocal: () => void this.openLocally(),
			onRefresh: () => void this.reload(),
			onWarningAction: (warningId, actionId) => this.onWarningAction(warningId, actionId),
		});
	}

	onload() {
		this.context.active.add(this);
		this.start();
	}

	onunload() {
		this.context.active.delete(this);
		this.stop();
		this.view.destroy();
	}

	/** Re-reads the block settings, e.g. after the OneDrive folder changed. */
	restart() {
		this.stop();
		this.start();
	}

	private start() {
		try {
			this.cfg = typeof this.source === "string" ? parseBlockConfig(this.source) : this.source;
			this.initialEditRequested = this.cfg.mode ? this.cfg.mode === "edit" : !this.context.settings.defaultReadOnly;
			// Never expose editable cells before the initial ownership check finishes.
			this.readOnly = true;
			this.file = this.context.resolver.resolve(this.cfg.file, this.notePath);
		} catch (e) {
			this.cfg = undefined;
			this.file = undefined;
			this.view.showError(errorMessage(e));
			return;
		}
		this.stopped = false;
		this.stopWatching = this.file.watch(() => this.queueReload());
		void this.reload();
	}

	private stop() {
		this.stopped = true;
		this.generation++;
		this.modeGeneration++;
		this.checkingMode = false;
		this.view.cancelEdit();
		window.clearTimeout(this.reloadTimer);
		this.stopWatching?.();
		this.stopWatching = undefined;
	}

	/** Defer our own save notifications, rather than dropping nearby external changes. */
	private queueReload() {
		window.clearTimeout(this.reloadTimer);
		if (this.stopped) return;
		this.reloadTimer = window.setTimeout(() => {
			if (this.busy || this.saves.size || Date.now() < this.ignoreChangesUntil) this.queueReload();
			else void this.reload();
		}, 300);
	}

	private async reload() {
		if (!this.cfg || !this.file || this.stopped) return;
		const generation = ++this.generation;
		try {
			const [wb, conflicts, openInExcel] = await Promise.all([
				readWorkbook(this.file), findConflictCopies(this.file), isOpenInExcel(this.file),
			]);
			if (this.stopped || generation !== this.generation) return;
			this.conflicts = conflicts;
			const wantsEdit = this.initialEditRequested || !this.readOnly;
			this.initialEditRequested = false;
			if (openInExcel && wantsEdit) {
				this.modeGeneration++; this.checkingMode = false;
				this.readOnly = true; this.view.cancelEdit();
				new Notice(`Excidian: ${EXCEL_OPEN_MESSAGE}`, 8000);
			} else if (wantsEdit) this.readOnly = false;
			this.show(wb);
		} catch (e) {
			if (!this.stopped && generation === this.generation) this.view.showError(errorMessage(e));
		}
	}

	private async changeMode(readOnly: boolean) {
		if (this.stopped || !this.file || this.busy) return;
		if (readOnly) {
			this.modeGeneration++; this.checkingMode = false;
			this.initialEditRequested = false;
			this.readOnly = true; this.view.cancelEdit(); this.render();
			return;
		}
		if (!this.readOnly || this.checkingMode) return;
		const file = this.file, generation = ++this.modeGeneration;
		const current = () => !this.stopped && this.file === file && this.modeGeneration === generation;
		this.checkingMode = true; this.render();
		try {
			const openInExcel = await isOpenInExcel(file);
			if (!current()) return;
			if (openInExcel) new Notice(`Excidian: ${EXCEL_OPEN_MESSAGE}`, 8000);
			else this.readOnly = false;
		} catch (error) {
			if (current()) new Notice(`Excidian: 无法确认文件是否被 Excel 占用，已保持只读。${errorMessage(error)}`, 8000);
		} finally {
			if (current()) { this.checkingMode = false; this.render(); }
		}
	}

	private async openLocally() {
		const file = this.file;
		if (this.stopped || this.busy || !file?.openLocal) return;
		const controllers = [...new Set([this, ...this.context.active])].filter((controller) => !controller.stopped && controller.file?.id === file.id);
		if (controllers.some((controller) => controller.busy)) {
			new Notice("Excidian: 文件正在保存，请稍后再用本地程序打开。", 6000);
			return;
		}
		const failures = controllers.map((controller) => controller.saveFailures);
		// Finish every editor for this workbook before locking all of its previews.
		for (const controller of controllers) controller.view.finishEdit();
		for (const controller of controllers) {
			controller.modeGeneration++; controller.checkingMode = false;
			controller.initialEditRequested = false; controller.readOnly = true;
			controller.busy = true; controller.view.cancelEdit(); controller.render();
		}
		try {
			await Promise.all(controllers.flatMap((controller) => [...controller.saves]));
			if (this.stopped || this.file !== file) return;
			if (controllers.some((controller, index) => controller.saveFailures !== failures[index])) {
				throw new Error("有修改未保存成功，未打开本地程序。请处理保存提示后重试。");
			}
			await file.openLocal();
		} catch (error) {
			if (!this.stopped && this.file === file) new Notice(`Excidian: 无法用本地程序打开。${errorMessage(error)}`, 8000);
		} finally {
			for (const controller of controllers) {
				controller.busy = false;
				if (!controller.stopped) controller.render();
			}
		}
	}

	private show(wb: Workbook) {
		this.workbook = wb;
		if (this.cfg?.target.kind === "sheet") {
			// Bind even a default/numeric selection to its name so external reordering
			// cannot send an in-flight edit to a different worksheet.
			if (!this.selectedSheet) this.selectedSheet = wb.sheet(this.cfg.target.sheet).name;
			else if (!wb.sheets.some((s) => s.name === this.selectedSheet)) this.selectedSheet = wb.sheet().name;
		}
		this.render();
	}

	private effectiveConfig(): BlockConfig {
		const cfg = this.cfg!;
		return cfg.target.kind === "sheet" && this.selectedSheet
			? { ...cfg, target: { ...cfg.target, sheet: this.selectedSheet } }
			: cfg;
	}

	private render() {
		if (!this.workbook || this.stopped) return;
		const history = workbookHistory(this.file!);
		this.grid = {
			...buildGrid(this.workbook, this.effectiveConfig(), this.readOnly, true),
			tooltip: this.file!.displayPath, warnings: this.conflictWarnings(), busy: this.busy || this.checkingMode,
			canOpenLocal: !!this.file!.openLocal,
			canUndo: history.undo, canRedo: history.redo,
		};
		this.view.render(this.grid);
	}

	private async structure(edit: StructureEdit) {
		if (this.readOnly || this.busy || this.stopped || !this.file || this.cfg?.target.kind !== "sheet") return;
		const sheet = this.selectedSheet!;
		await this.mutate(async (file) => updateWorkbook(file, (wb) => wb.editStructure(sheet, edit)));
	}

	private async history(direction: "undo" | "redo") {
		if (this.readOnly || this.busy || this.stopped || !this.file) return;
		await this.mutate((file) => restoreWorkbook(file, direction));
	}

	private async mutate(action: (file: StoredFile) => Promise<Workbook>) {
		const file = this.file!;
		this.busy = true; this.view.cancelEdit(); this.render();
		try {
			await Promise.all([...this.saves]);
			if (this.stopped || file !== this.file) return;
			this.ignoreChangesUntil = Date.now() + 10_000;
			if (await isOpenInExcel(file)) { this.readOnly = true; throw new Error(EXCEL_OPEN_MESSAGE); }
			const wb = await action(file);
			if (this.stopped || file !== this.file) return;
			this.generation++;
			this.show(wb);
			// Update other embeds immediately, including those suppressing a recent own save.
			for (const controller of this.context.active) if (controller !== this && controller.file?.id === file.id) void controller.reload();
			this.view.setStatus("已保存");
		} catch (error) {
			new Notice(`Excidian: ${errorMessage(error)}`, 8000);
			if (!this.stopped && file === this.file) await this.reload();
		} finally {
			this.busy = false; this.ignoreChangesUntil = Date.now() + 300;
			if (!this.stopped && file === this.file) this.render();
		}
	}

	private conflictWarnings(): WarningData[] {
		return this.conflicts
			.filter((copy) => !this.context.settings.isDismissed(this.conflictId(copy)))
			.map((copy) => ({
				id: this.conflictId(copy),
				text:
					`Possible sync conflict copy: ${copy.name} (changed ${copy.modified.toLocaleString()}). ` +
					"Some entries may be in that file instead of this one. Merge them in Excel, then delete the copy.",
				// "Show file" needs a file manager, which only desktop has.
				actions: [...(copy.reveal ? [{ id: "reveal", label: "Show file" }] : []), { id: "dismiss", label: "Dismiss" }],
			}));
	}

	/** Includes the modification time, so a new conflict with the same name warns again. */
	private conflictId(copy: SiblingFile): string {
		return `${this.file?.id}|${copy.name}|${copy.modified.getTime()}`;
	}

	private onWarningAction(warningId: string, actionId: string) {
		const copy = this.conflicts.find((c) => this.conflictId(c) === warningId);
		if (!copy) return;
		if (actionId === "reveal") {
			try {
				copy.reveal?.();
			} catch {
				new Notice(`Excidian: ${copy.name}`);
			}
		} else if (actionId === "dismiss" && this.grid) {
			this.context.settings.dismiss(warningId);
			this.grid = { ...this.grid, warnings: this.conflictWarnings() };
			this.view.render(this.grid);
		}
	}

	private async formatCells(cells: CellData[], patch: CellStylePatch) {
		if (this.readOnly || this.busy || this.stopped || !this.file || !this.cfg || !cells.length) return;
		const cfg = this.effectiveConfig();
		await this.mutate((file) => updateWorkbook(file, (wb) => {
			if (cells.length > 10_000) throw new Error("每次最多设置 10,000 个单元格的格式。");
			const sheet = cfg.target.kind === "sheet" ? wb.sheet(cfg.target.sheet) : wb.table(cfg.target.table).sheet;
			const merges = sheet.mergedRanges();
			for (const cell of cells) {
				if (sheet.isRowHidden(cell.row) || merges.some((range) => inRange(range, cell.row, cell.col))) throw new Error("选区包含隐藏行或合并单元格，请刷新后重新选择。");
				assertUnchanged(sheet, cell);
				sheet.formatCell(cell.row, cell.col, patch);
			}
		}, { recalculate: false }));
	}

	private async batchCommit(changes: CellChange[]) {
		if (this.readOnly || this.busy || this.stopped || !this.file || !this.cfg || !changes.length) return;
		const cfg = this.effectiveConfig();
		let warnings = 0;
		await this.mutate((file) => updateWorkbook(file, (wb) => {
			if (changes.length > 10_000) throw new Error("每次最多修改 10,000 个单元格。");
			const seen = new Set<string>();
			// Check all expected values before applying any in-memory table expansion.
			for (const { cell } of changes) assertUnchanged(editTarget(wb, cfg, cell).sheet, cell);
			for (const { cell, text } of changes) {
				if (!cell.editable) throw new Error("区域含只读单元格。");
				const ref = formatRef(cell.row, cell.col);
				if (seen.has(ref)) throw new Error("区域包含重复单元格。");
				seen.add(ref);
				const { sheet } = editTarget(wb, cfg, cell);
				if (sheet.mergedRanges().some((range) => cell.row >= range.top && cell.row <= range.bottom && cell.col >= range.left && cell.col <= range.right)) throw new Error("区域含合并单元格，请在普通数据区域粘贴。");
				if (sheet.isTableHeader(cell.row, cell.col)) throw new Error("请在 Excel 中修改表格标题。");
				sheet.assertWritable(cell.row, cell.col);
				const input = parseInput(text, sheet.cell(cell.row, cell.col));
				const issue = checkEdit(wb, cfg, cell, input);
				if (issue?.blocking) throw new Error(`${ref}: ${issueText(issue)}（整个区域未保存）`);
				if (issue) warnings++;
				// Changes remain in memory until every cell has passed validation.
				applyEdit(wb, cfg, cell, input);
			}
		}));
		if (warnings) new Notice(`Excidian: ${warnings} 个单元格未满足提示性验证规则，请检查。`, 8000);
	}

	private async commit(cell: CellData, text: string) {
		if (this.readOnly || this.busy || !cell.editable || this.stopped || !this.cfg || !this.file) return;
		// Capture the sheet now: switching tabs during an async save must never redirect an edit.
		const cfg = this.effectiveConfig();
		const file = this.file!;


		this.ignoreChangesUntil = Date.now() + 10_000;
		let warning = null as ValidationIssue | null;
		try {
			if (await isOpenInExcel(file)) {
				this.readOnly = true;
				throw new Error(EXCEL_OPEN_MESSAGE);
			}
			const wb = await updateWorkbook(file, (wb) => {
				// Interpret using the fresh cell type/format, right before writing.
				const { sheet } = editTarget(wb, cfg, cell);
				assertUnchanged(sheet, cell);
				const input = parseInput(text, sheet.cell(cell.row, cell.col));
				if (sheet.isTableHeader(cell.row, cell.col)) throw new Error("请在 Excel 中修改表格标题。");
				const issue = checkEdit(wb, cfg, cell, input);
				if (issue?.blocking) throw new Error(issueText(issue));
				warning = issue;
				applyEdit(wb, cfg, cell, input);
			});
			if (this.stopped || file !== this.file) return;
			this.generation++;
			this.show(wb);
			this.view.setStatus(`Saved ${formatRef(cell.row, cell.col)}`);
			for (const controller of this.context.active) if (controller !== this && controller.file?.id === file.id) void controller.reload();
			if (warning) new Notice(`Excidian: saved, but ${issueText(warning)}`, 8000);
		} catch (e) {
			this.saveFailures++;
			new Notice(`Excidian: could not save ${formatRef(cell.row, cell.col)}: ${errorMessage(e)}`);
			if (!this.stopped && file === this.file) await this.reload();
		} finally {
			this.ignoreChangesUntil = Date.now() + 1000;
		}
	}
}

/**
 * The table an edit goes through: the block's table, or in sheet mode the table the
 * cell is in (or just below), so typing under a table grows it like in Excel.
 */
function editTarget(wb: Workbook, cfg: BlockConfig, cell: CellData): { sheet: Sheet; location?: TableLocation } {
	if (cfg.target.kind === "table") {
		const location = wb.table(cfg.target.table);
		return { sheet: location.sheet, location };
	}
	const sheet = wb.sheet(cfg.target.sheet);
	const location = wb.tableAt(sheet, cell.row, cell.col);
	return { sheet, location: location && cell.row >= location.table.dataRange.top ? location : undefined };
}

function checkEdit(wb: Workbook, cfg: BlockConfig, cell: CellData, input: CellInput): ValidationIssue | null {
	const { sheet, location } = editTarget(wb, cfg, cell);
	return wb.checkInput(sheet, cell.row, cell.col, input, location?.table);
}

function issueText(issue: ValidationIssue): string {
	return issue.title ? `${issue.title}: ${issue.message}` : issue.message;
}

function applyEdit(wb: Workbook, cfg: BlockConfig, cell: CellData, input: CellInput) {
	const { sheet, location } = editTarget(wb, cfg, cell);
	if (location) wb.writeTableCell(location, cell.row, cell.col, input);
	else sheet.setCell(cell.row, cell.col, input);
}


function errorMessage(e: unknown): string {
	return e instanceof Error ? e.message : String(e);
}

function assertUnchanged(sheet: Sheet, cell: CellData) {
	const current = sheet.cell(cell.row, cell.col);
	if (editText(current.value, current.formula) !== cell.editText) throw new Error(`${formatRef(cell.row, cell.col)} 已被其他预览或程序修改，请刷新后重试。`);
}
