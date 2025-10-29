/**
 * Type definitions for extracted DOCX document data
 */

export interface ExtractedDocument {
  paragraphs: ExtractedParagraph[];
  tables: ExtractedTable[];
  body: Array<{ type: 'paragraph'; data: ExtractedParagraph } | { type: 'table'; data: ExtractedTable }>; // Document body in order
  styles: Map<string, StyleInfo>;
  defaults?: RunFormatting; // Document-wide default run formatting
  paragraphDefaults?: ParagraphDefaults; // Document-wide default paragraph properties
  numbering?: Map<string, NumberingDefinition>; // Numbering definitions extracted from numbering.xml
  mediaFiles?: Map<string, Uint8Array>; // Media files (images, charts, etc.) from word/media/
  // Note: All content (document, styles, numbering) is BUILT from structured data
  // This allows full modification of all document aspects
}

export interface ExtractedParagraph {
  runs: ExtractedRun[];
  spacing?: SpacingInfo;
  alignment?: 'left' | 'center' | 'right' | 'justify';
  indentation?: IndentationInfo;
  styleName?: string;
  numbering?: NumberingInfo;
  isEmpty: boolean;
}

export interface ExtractedRun {
  text?: string;
  formatting?: RunFormatting;
  image?: ExtractedImage;
}

export interface ExtractedImage {
  src: string; // relationship ID (for reference)
  mediaPath?: string; // actual media file path (e.g., "word/media/image1.png")
  data?: string; // base64 encoded image data
  contentType?: string; // image/png, image/jpeg, etc.
  width?: number; // in points
  height?: number; // in points
}

export interface RunFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number; // in points
  fontFamily?: string;
  color?: string; // hex color
  colorTheme?: string; // theme color reference (e.g., "accent1")
  colorThemeShade?: string; // theme shade/tint (e.g., "BF")
  fontThemeAscii?: string; // ascii theme font reference (e.g., "majorHAnsi")
  fontThemeEastAsia?: string; // east asia theme font
  fontThemeHAnsi?: string; // high ansi theme font
  fontThemeCs?: string; // complex script theme font
  highlight?: string;
  verticalAlign?: 'superscript' | 'subscript'; // w:vertAlign
}

export interface SpacingInfo {
  before?: number; // in points
  after?: number; // in points
  line?: number; // line spacing value
  lineRule?: 'atLeast' | 'exactly' | 'auto';
}

export interface IndentationInfo {
  left?: number; // in points
  right?: number; // in points
  firstLine?: number; // in points
  hanging?: number; // in points
}

export interface NumberingInfo {
  id: string;
  level?: number; // Optional - may not be present in style definitions
}

export interface ExtractedTable {
  rows: ExtractedTableRow[];
  columnWidths?: number[];
}

export interface ExtractedTableRow {
  cells: ExtractedTableCell[];
}

export interface ExtractedTableCell {
  content: ExtractedParagraph[];
  colSpan?: number;
  rowSpan?: number;
  width?: number; // in points
  backgroundColor?: string; // hex color
  borders?: CellBorders;
  verticalAlign?: 'top' | 'center' | 'bottom';
  margins?: CellMargins;
}

export interface CellBorders {
  top?: BorderInfo;
  bottom?: BorderInfo;
  left?: BorderInfo;
  right?: BorderInfo;
}

export interface BorderInfo {
  style?: string; // 'single', 'double', 'dashed', etc.
  size?: number; // in eighths of a point
  color?: string; // hex color
}

export interface CellMargins {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}

export interface StyleInfo {
  name: string;
  type?: 'paragraph' | 'character' | 'table' | 'numbering'; // style type
  basedOn?: string;
  next?: string; // next style to apply
  link?: string; // linked character style
  runFormatting?: RunFormatting;
  spacing?: SpacingInfo;
  alignment?: 'left' | 'center' | 'right' | 'justify';
  indentation?: IndentationInfo;
  numbering?: NumberingInfo; // numbering defined in the style (for list styles)
  contextualSpacing?: boolean; // contextual spacing
  isDefault?: boolean;
  // Paragraph metadata
  keepNext?: boolean; // keep with next paragraph
  keepLines?: boolean; // keep lines together
  outlineLevel?: number; // outline level (0-8)
  // Style metadata
  uiPriority?: number;
  qFormat?: boolean; // primary style
  unhideWhenUsed?: boolean; // unhide when used
  // Table style properties (raw XML for now - full extraction would be complex)
  tablePropertiesXml?: string; // <w:tblPr> content
}

export interface NumberingDefinition {
  numId: string;
  abstractNumId: string;
  nsid?: string; // Namespace ID - unique identifier for the abstract numbering
  multiLevelType?: string; // 'singleLevel', 'hybridMultilevel', 'multilevel'
  tmpl?: string; // Template ID - identifies the template this numbering is based on
  levels: NumberingLevel[];
}

export interface NumberingLevel {
  level: number;
  start?: number; // Starting number (default 1)
  format: string; // 'bullet', 'decimal', 'lowerLetter', etc.
  text: string; // e.g., '●', '%1.', etc.
  paragraphStyleName?: string; // w:pStyle - links numbering to paragraph style
  alignment?: 'left' | 'center' | 'right';
  indentation?: IndentationInfo;
  tabs?: TabStop[]; // Tab stops in the numbering level
  fontFamily?: string;
  fontHint?: string; // Font hint attribute (e.g., "default")
}

export interface TabStop {
  val: 'num' | 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear'; // Tab alignment
  pos: number; // Position in points
}

export interface ParagraphDefaults {
  spacing?: SpacingInfo;
  alignment?: 'left' | 'center' | 'right' | 'justify';
  indentation?: IndentationInfo;
}

