import { MarkdownRenderChild, Plugin, TFile } from "obsidian";
import { SheetController, type ControllerContext } from "../controller/sheet-controller";
import { VaultFile } from "../platform/vault-file";

interface EmbedContext {
	containerEl: HTMLElement;
	sourcePath?: string;
}

// Obsidian uses this registry for both Reading view and Live Preview. It is an
// internal API, so keep the dependency here and fall back to the public processor.
interface EmbedRegistry {
	registerExtension(extension: string, create: (ctx: EmbedContext, file: TFile, subpath?: string) => MarkdownRenderChild): void;
	unregisterExtension(extension: string): void;
}

class WorkbookEmbed extends MarkdownRenderChild {
	private controller?: SheetController;
	private loaded = false;

	constructor(
		private plugin: Plugin,
		containerEl: HTMLElement,
		private file: TFile,
		private context: ControllerContext,
		private sourcePath = "",
		private subpath?: string,
	) {
		super(containerEl);
	}

	onload() {
		this.loaded = true;
		this.mount();
	}

	onunload() {
		this.loaded = false;
	}

	loadFile(file?: TFile) {
		if (file && file.path !== this.file.path) {
			this.file = file;
			this.subpath = undefined;
			if (this.controller) this.removeChild(this.controller);
			this.controller = undefined;
		}
		if (this.loaded) this.mount();
	}

	private mount() {
		if (this.controller) return;
		this.containerEl.empty();
		this.containerEl.addClass("excidian-embed");
		this.controller = this.addChild(new SheetController(
			this.containerEl.createDiv(),
			{ file: this.file.path, target: { kind: "sheet", sheet: this.subpath?.replace(/^#/, "") || undefined }, headings: true },
			this.sourcePath,
			{
				...this.context,
				// The registry already resolves wiki links (short names, relative links,
				// aliases and renamed files). Use the vault API for the exact target.
				resolver: { resolve: () => new VaultFile(this.plugin.app, this.file.path) },
			},
		));
	}
}

export function registerWorkbookEmbeds(plugin: Plugin, context: ControllerContext) {
	const registry = (plugin.app as unknown as { embedRegistry?: EmbedRegistry }).embedRegistry;
	if (registry?.registerExtension && registry?.unregisterExtension) {
		try {
			registry.registerExtension("xlsx", (ctx, file, subpath) =>
				new WorkbookEmbed(plugin, ctx.containerEl, file, context, ctx.sourcePath, subpath));
			plugin.register(() => registry.unregisterExtension("xlsx"));
			return;
		} catch (error) {
			console.warn("Excidian-dev: native embeds unavailable; using Reading view fallback.", error);
		}
	}
	plugin.registerMarkdownPostProcessor((el, ctx) => {
		const embeds = Array.from(el.querySelectorAll<HTMLElement>(".internal-embed"));
		if (el.matches(".internal-embed")) embeds.unshift(el);
		for (const embed of embeds) {
			if (embed.hasClass("excidian-embed")) continue;
			const src = embed.getAttribute("src") ?? "";
			const hash = src.indexOf("#");
			const path = hash < 0 ? src : src.slice(0, hash);
			if (!/\.xlsx$/i.test(path)) continue;
			const file = plugin.app.metadataCache.getFirstLinkpathDest(path, ctx.sourcePath);
			if (file instanceof TFile) {
				ctx.addChild(new WorkbookEmbed(plugin, embed, file, context, ctx.sourcePath, hash < 0 ? undefined : src.slice(hash + 1)));
			}
		}
	});
}
