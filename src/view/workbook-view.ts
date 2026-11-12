import { FileView, WorkspaceLeaf, type OpenViewState, type Plugin, type TFile, type Workspace } from "obsidian";
import { SheetController, type ControllerContext } from "../controller/sheet-controller";

const VIEW_TYPE = "excidian-workbook";
const EXTENSIONS = ["xlsx", "XLSX"];

function workbookLeaf(workspace: Workspace, file: TFile): WorkspaceLeaf {
	return workspace.getLeavesOfType(VIEW_TYPE).find((leaf) => leaf.getViewState().state?.file === file.path)
		?? workspace.getLeaf("tab");
}

export function openWorkbookTab(workspace: Workspace, file: TFile, sheet?: string): Promise<void> {
	return workbookLeaf(workspace, file).openFile(file, {
		active: true,
		eState: sheet ? { subpath: `#${sheet}` } : undefined,
	});
}

class WorkbookView extends FileView {
	icon = "file-spreadsheet";
	private controller?: SheetController;

	constructor(leaf: WorkspaceLeaf, private context: ControllerContext) {
		super(leaf);
		this.contentEl.addClass("excidian-workbook-view");
	}

	getViewType() {
		return VIEW_TYPE;
	}

	async onLoadFile(file: TFile) {
		this.controller = this.addChild(new SheetController(
			this.contentEl.createDiv(),
			{ file: file.path, target: { kind: "sheet" }, headings: true },
			file.path,
			this.context,
		));
	}

	async onUnloadFile() {
		if (this.controller) this.removeChild(this.controller);
		this.controller = undefined;
		this.contentEl.empty();
	}

	async onRename(file: TFile) {
		await super.onRename(file);
		if (file !== this.file) return;
		await this.onUnloadFile();
		if (file === this.file) await this.onLoadFile(file);
	}

	setEphemeralState(state: { subpath?: unknown } | undefined) {
		super.setEphemeralState(state);
		if (typeof state?.subpath === "string" && state.subpath.startsWith("#")) {
			this.controller?.selectSheet(state.subpath.slice(1));
		}
	}
}

export function registerWorkbookView(plugin: Plugin, context: ControllerContext) {
	plugin.registerView(VIEW_TYPE, (leaf) => new WorkbookView(leaf, context));
	plugin.registerExtensions(EXTENSIONS, VIEW_TYPE);

	// The extension registry chooses the view, but Obsidian still opens ordinary
	// explorer and link clicks in the current leaf. Redirect just workbook opens
	// so the source note stays open. An explicitly created empty leaf is used as is.
	const { workspace } = plugin.app;
	const original = WorkspaceLeaf.prototype.openFile;
	let enabled = true;
	const openFile = function (this: WorkspaceLeaf, file: TFile, options?: OpenViewState) {
		if (!enabled || this.view?.app !== plugin.app || !EXTENSIONS.includes(file?.extension) || options?.group) {
			return original.call(this, file, options);
		}
		const state = this.getViewState();
		if (state.type === "empty" || (state.type === VIEW_TYPE && state.state?.file === file.path)) {
			return original.call(this, file, options);
		}
		const target = workbookLeaf(workspace, file);
		return original.call(target, file, { ...options, active: options?.active ?? true });
	};
	WorkspaceLeaf.prototype.openFile = openFile;
	plugin.register(() => {
		enabled = false;
		if (WorkspaceLeaf.prototype.openFile === openFile) WorkspaceLeaf.prototype.openFile = original;
	});
}
