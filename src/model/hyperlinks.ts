/** Only explicit web/email links are exposed as actions; other workbook links stay untouched. */
export function safeLink(value: string): string | undefined {
	if (value !== value.trim() || !/^(https?:\/\/|mailto:)/i.test(value) || /[\u0000-\u0020]/.test(value)) return undefined;
	try {
		const url = new URL(value);
		if (url.protocol === "mailto:") return url.pathname ? value : undefined;
		return url.hostname ? value : undefined;
	} catch { return undefined; }
}
