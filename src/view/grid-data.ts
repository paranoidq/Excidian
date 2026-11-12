import type { Range } from "../model/ooxml/refs";
import type { TextRun } from "../model/ooxml/rich-text";
import type { CellFont } from "../model/ooxml/styles";

/** Everything the grid view needs to draw. Built by the controller from the model. */
export interface GridData {
	/** Whole-sheet structural actions are offered only in sheet previews. */
	structureEditable?: boolean;
	columnNumbers?: number[];
	frozen?: { rows: number; columns: number };
	busy?: boolean;
	canOpenLocal?: boolean;
	canUndo?: boolean;
	canRedo?: boolean;
	readOnly?: boolean;
	sheets?: string[];
	activeSheet?: string;
	title: string;
	tooltip: string;
	columns: string[];
	/** Column widths in Excel's unit (about one character each). */
	columnWidths?: number[];
	/**
	 * Fit the page width, splitting it in the proportions of `columnWidths`;
	 * otherwise columns get their real widths and the grid scrolls sideways.
	 */
	fit?: boolean;
	/** Show the column header row (default true). */
	columnHeader?: boolean;
	/** Show row numbers (default true). */
	rowNumbers?: boolean;
	rows: RowData[];
	note?: string;
	warnings?: WarningData[];
}

/** A banner above the grid, with buttons reported back through onWarningAction. */
export interface WarningData {
	id: string;
	text: string;
	actions: { id: string; label: string }[];
}

export type RowKind = "data" | "totals" | "new";

export interface RowData {
	label: string;
	kind: RowKind;
	cells: CellData[];
}

export interface CellData {
	/** Sheet coordinates, handed back to the controller on edit. */
	row: number;
	col: number;
	text: string;
	runs?: TextRun[];
	link?: string;
	editText: string;
	editable: boolean;
	numeric?: boolean;
	formula?: boolean;
	/** A formula with no calculated result yet. */
	pending?: boolean;
	/** Tooltip. */
	hint?: string;
	/** Choices offered in a dropdown while editing (from a list validation rule). */
	options?: { label: string; value: string }[];
	/** Shown under the cell while editing, e.g. the rule's input message. */
	help?: string;
	/** A merged cell's size, in grid columns and rows. */
	span?: { cols: number; rows: number };
	mergeRange?: Range;
	hasContent?: boolean;
	tableCell?: boolean;
	/** Covered by a merged cell to its left or above: not drawn. */
	covered?: boolean;
	/** The header row of an Excel table shown inside a sheet. */
	tableHeader?: boolean;
	style?: CellStyle;
}

/** CSS-ready formatting. */
export interface CellStyle extends CellFont {
	background?: string;
	align?: "left" | "center" | "right" | "justify";
	wrap?: boolean;
}
