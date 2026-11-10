import { App, Plugin, PluginSettingTab, SettingDefinitionItem } from "obsidian";
import type { DeviceSettings } from "../model/device-settings";

export interface SettingsTabOptions {
	settings: DeviceSettings;
	/** OneDrive paths only work on desktop; mobile reaches files inside the vault only. */
	showOneDrive: boolean;
	detectOneDriveRoots(): string[];
	/** Called when the tab closes after something changed. */
	onChange(): void;
}

const ONEDRIVE_ROOT = "oneDriveRoot";
const DEFAULT_READ_ONLY = "defaultReadOnly";

/**
 * Settings → Excidian, declared with Obsidian's settings API so the options also
 * show up in settings search. Values live in this device's storage, not the vault.
 */
export class ExcidianSettingTab extends PluginSettingTab {
	private changed = false;

	constructor(
		app: App,
		plugin: Plugin,
		private options: SettingsTabOptions,
	) {
		super(app, plugin);
	}

	getSettingDefinitions(): SettingDefinitionItem[] {
		const { settings } = this.options;
		const detected = this.options.detectOneDriveRoots();
		const dismissed = settings.dismissedCount;
		return [
			{
				name: "默认只读模式",
				desc: "预览 Excel 时默认只读。关闭后默认可编辑；每个预览都可以独立切换模式。此设置仅保存在当前设备。",
				aliases: ["readonly", "edit", "只读", "编辑"],
				control: { type: "toggle", key: DEFAULT_READ_ONLY },
			},
			{
				name: "OneDrive folder",
				desc: createFragment((desc) => {
					desc.appendText("Used by ");
					desc.createEl("code", { text: "file: onedrive:/…" });
					desc.appendText(" paths. Saved on this device only, so each computer can point to its own. Leave empty to detect it automatically.");
					desc.createEl("br");
					desc.appendText(detected.length ? `Detected: ${detected.join(", ")}` : "No OneDrive folder detected on this device.");
				}),
				aliases: ["onedrive", "cloud", "sync"],
				visible: this.options.showOneDrive,
				control: { type: "text", key: ONEDRIVE_ROOT, placeholder: detected[0] ?? "/path/to/OneDrive" },
			},
			{
				name: "Show dismissed conflict warnings again",
				desc: dismissed ? `${dismissed} dismissed on this device.` : "None dismissed.",
				disabled: () => this.options.settings.dismissedCount === 0,
				action: () => {
					settings.clearDismissed();
					this.changed = true;
					this.update();
				},
			},
		];
	}

	getControlValue(key: string): unknown {
		if (key === DEFAULT_READ_ONLY) return this.options.settings.defaultReadOnly;
		return key === ONEDRIVE_ROOT ? this.options.settings.oneDriveRoot : undefined;
	}

	setControlValue(key: string, value: unknown) {
		if (key === DEFAULT_READ_ONLY && typeof value === "boolean") {
			this.options.settings.defaultReadOnly = value;
			this.changed = true;
		}
		if (key === ONEDRIVE_ROOT && typeof value === "string") {
			this.options.settings.oneDriveRoot = value;
			this.changed = true;
		}
	}

	hide() {
		super.hide();
		if (this.changed) {
			this.changed = false;
			this.options.onChange();
		}
	}
}
