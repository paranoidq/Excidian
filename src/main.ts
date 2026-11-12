import { Modal, Platform, Plugin, Scope } from "obsidian";
import { ControllerContext, SheetController } from "./controller/sheet-controller";
import { DeviceSettings } from "./model/device-settings";
import type { FileResolver } from "./model/storage";
import { VaultResolver } from "./platform/vault-file";
import { ExcidianSettingTab } from "./view/settings-tab";
import { registerWorkbookEmbeds } from "./view/workbook-embed";
import { registerWorkbookView } from "./view/workbook-view";

export default class ExcidianPlugin extends Plugin {
	async onload() {
		const settings = new DeviceSettings({
			load: (key): unknown => {
				const value: unknown = this.app.loadLocalStorage(key);
				return value;
			},
			save: (key, value) => this.app.saveLocalStorage(key, value),
		});

		// Desktop reaches any file on the computer through Node; mobile has no Node,
		// so it works with files inside the vault through Obsidian's vault API.
		// The desktop module is imported dynamically so mobile never loads Node code.
		let resolver: FileResolver;
		let detectOneDriveRoots = (): string[] => [];
		if (Platform.isDesktopApp) {
			const desktop = await import("./platform/desktop");
			resolver = new desktop.DesktopResolver(this.app, settings);
			detectOneDriveRoots = () => desktop.detectOneDriveRoots();
		} else {
			resolver = new VaultResolver(this.app);
		}

		const context: ControllerContext = {
			settings, resolver, active: new Set(),
			confirmMerge: (address, clearedCells) => new Promise<boolean>((resolve) => {
				const modal = new Modal(this.app);
				let confirmed = false;
				modal.setTitle(`合并 ${address}`);
				modal.setContent(`合并后只保留左上角的内容，其余 ${clearedCells} 个单元格的内容会被清除。`);
				const actions = modal.contentEl.createDiv("modal-button-container");
				const cancel = actions.createEl("button", { text: "取消" });
				cancel.onclick = () => modal.close();
				actions.createEl("button", { cls: "mod-warning", text: "合并并清除" }).onclick = () => {
					confirmed = true;
					modal.close();
				};
				modal.onOpen = () => cancel.focus();
				modal.setCloseCallback(() => resolve(confirmed));
				modal.open();
			}),
			registerShortcutScope: (register) => {
				const scope = new Scope(this.app.scope);
				register(scope);
				this.app.keymap.pushScope(scope);
				return () => this.app.keymap.popScope(scope);
			},
		};

		this.addSettingTab(
			new ExcidianSettingTab(this.app, this, {
				settings,
				showOneDrive: Platform.isDesktopApp,
				detectOneDriveRoots,
				onChange: () => context.active.forEach((controller) => controller.restart()),
			}),
		);

		this.registerMarkdownCodeBlockProcessor("excidian", (source, el, ctx) => {
			ctx.addChild(new SheetController(el, source, ctx.sourcePath, context));
		});
		registerWorkbookEmbeds(this, context);
		registerWorkbookView(this, context);
		this.register(() => {
			for (const controller of [...context.active]) controller.unload();
		});
	}
}
