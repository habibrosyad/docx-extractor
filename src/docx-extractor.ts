/**
 * Main DOCX Extractor
 * Built on docxjs's parser approach but focused on data extraction
 */

import JSZip from 'jszip';
import { JSDOM } from 'jsdom';
import {
  ExtractedDocument,
  ExtractedParagraph,
  ExtractedRun,
  ExtractedTable,
  ExtractedTableRow,
  ExtractedTableCell,
  ExtractedImage,
  RunFormatting,
  SpacingInfo,
  IndentationInfo,
  NumberingInfo,
  StyleInfo,
  NumberingDefinition,
  NumberingLevel,
  ParagraphDefaults
} from './types.js';
import { XmlParser } from './xml-parser.js';

export class DocxExtractor {
  private xml: XmlParser;
  private domParser: DOMParser;
  private xmlSerializer: XMLSerializer;
  private zip: JSZip | null = null;
  private relationships: Map<string, string> = new Map();
  private styles: Map<string, StyleInfo> = new Map();
  private documentDefaults: RunFormatting = {};
  private paragraphDefaults: ParagraphDefaults = {};
  private numbering: Map<string, NumberingDefinition> = new Map();

  constructor() {
    this.xml = new XmlParser();
    // Initialize DOM parser (works in Node.js with jsdom)
    const dom = new JSDOM();
    this.domParser = new dom.window.DOMParser();
    this.xmlSerializer = new dom.window.XMLSerializer();

    // Set up global references for XML parser if not already set
    if (typeof (global as any).window === 'undefined') {
      (global as any).window = dom.window;
      (global as any).DOMParser = dom.window.DOMParser;
      (global as any).XMLSerializer = dom.window.XMLSerializer;
      (global as any).Node = dom.window.Node;
      (global as any).Element = dom.window.Element;
      (global as any).Text = dom.window.Text;
      (global as any).Comment = dom.window.Comment;
      (global as any).DocumentFragment = dom.window.DocumentFragment;
    }
  }

  async extract(buffer: ArrayBuffer | Buffer): Promise<ExtractedDocument> {
    this.zip = await JSZip.loadAsync(buffer);

    // Load relationships
    await this.loadRelationships();

    // Load styles
    await this.loadStyles();

    // Load numbering definitions
    await this.loadNumbering();

    // Load document.xml for parsing (we'll BUILD it back from structured data)
    const documentXmlString = await this.zip.file('word/document.xml')?.async('string');
    if (!documentXmlString) {
      throw new Error('Invalid DOCX file: document.xml not found');
    }

    const doc = this.domParser.parseFromString(documentXmlString, 'text/xml');
    const body = this.xml.element(doc.documentElement, 'body');

    if (!body) {
      throw new Error('Invalid DOCX file: body not found');
    }

    // Load all media files (preserve for images/charts that are referenced in document)
    const mediaFiles = new Map<string, Uint8Array>();
    for (const filePath in this.zip.files) {
      if (filePath.startsWith('word/media/') && !this.zip.files[filePath].dir) {
        const content = await this.zip.file(filePath)?.async('uint8array');
        if (content) {
          mediaFiles.set(filePath, content);
        }
      }
    }

    const paragraphs: ExtractedParagraph[] = [];
    const tables: ExtractedTable[] = [];
    const bodyElements: Array<{ type: 'paragraph'; data: ExtractedParagraph } | { type: 'table'; data: ExtractedTable }> = [];

    for (const child of this.xml.elements(body)) {
      switch (child.localName) {
        case 'p':
          const para = await this.extractParagraph(child);
          paragraphs.push(para);
          bodyElements.push({ type: 'paragraph', data: para });
          break;
        case 'tbl':
          const table = await this.extractTable(child);
          tables.push(table);
          bodyElements.push({ type: 'table', data: table });
          break;
      }
    }

    return {
      paragraphs,
      tables,
      body: bodyElements,
      styles: this.styles,
      defaults: Object.keys(this.documentDefaults).length > 0 ? this.documentDefaults : undefined,
      paragraphDefaults: Object.keys(this.paragraphDefaults).length > 0 ? this.paragraphDefaults : undefined,
      numbering: this.numbering.size > 0 ? this.numbering : undefined,
      mediaFiles: mediaFiles.size > 0 ? mediaFiles : undefined
    };
  }

  private async resolveThemeFont(themeRef: string): Promise<string | null> {
    if (!this.zip) return null;

    const themeXml = await this.zip.file('word/theme/theme1.xml')?.async('string');
    if (!themeXml) return null;

    try {
      const doc = this.domParser.parseFromString(themeXml, 'text/xml');
      const fontScheme = this.xml.element(doc.documentElement, 'fontScheme');
      if (!fontScheme) return null;

      // Map theme references to actual fonts
      let fontElement: Element | null = null;

      if (themeRef === 'minorLatin' || themeRef === '+mn-lt' || themeRef.includes('minor')) {
        const minorFont = this.xml.element(fontScheme, 'minorFont');
        if (minorFont) {
          fontElement = this.xml.element(minorFont, 'latin');
        }
      } else if (themeRef === 'majorLatin' || themeRef === '+mj-lt' || themeRef.includes('major')) {
        const majorFont = this.xml.element(fontScheme, 'majorFont');
        if (majorFont) {
          fontElement = this.xml.element(majorFont, 'latin');
        }
      }

      if (fontElement) {
        const typeface = this.xml.attr(fontElement, 'typeface');
        return typeface;
      }
    } catch (error) {
      console.error('Error parsing theme:', error);
    }

    return null;
  }

  private async loadRelationships(): Promise<void> {
    if (!this.zip) return;

    const relsXml = await this.zip.file('word/_rels/document.xml.rels')?.async('string');
    if (!relsXml) return;

    const doc = this.domParser.parseFromString(relsXml, 'text/xml');
    const relationships = this.xml.elements(doc.documentElement, 'Relationship');

    for (const rel of relationships) {
      const id = this.xml.attr(rel, 'Id');
      const target = this.xml.attr(rel, 'Target');
      if (id && target) {
        this.relationships.set(id, target);
      }
    }
  }

  private async loadStyles(): Promise<void> {
    if (!this.zip) return;

    const stylesXml = await this.zip.file('word/styles.xml')?.async('string');
    if (!stylesXml) return;

    const doc = this.domParser.parseFromString(stylesXml, 'text/xml');

    // Extract document defaults
    const docDefaults = this.xml.element(doc.documentElement, 'docDefaults');
    if (docDefaults) {
      const rPrDefault = this.xml.element(docDefaults, 'rPrDefault');
      if (rPrDefault) {
        const rPr = this.xml.element(rPrDefault, 'rPr');
        if (rPr) {
          this.documentDefaults = this.extractRunFormattingFromElement(rPr);

          // Check for theme font references (like +mn-lt for minor Latin font)
          const rFonts = this.xml.element(rPr, 'rFonts');
          if (rFonts) {
            const ascii = this.xml.attr(rFonts, 'ascii');
            const asciiTheme = this.xml.attr(rFonts, 'asciiTheme');

            // If theme font is referenced, try to resolve it
            if (asciiTheme || (ascii && ascii.startsWith('+'))) {
              const themeFontName = await this.resolveThemeFont(asciiTheme || ascii || '');
              if (themeFontName) {
                this.documentDefaults.fontFamily = themeFontName;
              }
            }
          }

          console.log('📌 Document defaults extracted:', JSON.stringify(this.documentDefaults));
        }
      }

      // Extract paragraph defaults (pPrDefault)
      const pPrDefault = this.xml.element(docDefaults, 'pPrDefault');
      if (pPrDefault) {
        const pPr = this.xml.element(pPrDefault, 'pPr');
        if (pPr) {
          // Extract spacing
          const spacing = this.xml.element(pPr, 'spacing');
          if (spacing) {
            this.paragraphDefaults.spacing = {
              before: this.xml.lengthAttr(spacing, 'before') ?? undefined,
              after: this.xml.lengthAttr(spacing, 'after') ?? undefined,
              line: this.xml.lengthAttr(spacing, 'line') ?? undefined,
              lineRule: this.xml.attr(spacing, 'lineRule') as any ?? undefined
            };
          }

          // Extract alignment
          const jc = this.xml.element(pPr, 'jc');
          if (jc) {
            this.paragraphDefaults.alignment = this.xml.attr(jc, 'val') as any;
          }

          // Extract indentation
          const ind = this.xml.element(pPr, 'ind');
          if (ind) {
            this.paragraphDefaults.indentation = {
              left: this.xml.lengthAttr(ind, 'left') ?? undefined,
              right: this.xml.lengthAttr(ind, 'right') ?? undefined,
              firstLine: this.xml.lengthAttr(ind, 'firstLine') ?? undefined,
              hanging: this.xml.lengthAttr(ind, 'hanging') ?? undefined
            };
          }

          console.log('📌 Paragraph defaults extracted:', JSON.stringify(this.paragraphDefaults));
        }
      }
    }

    const styleElements = this.xml.elements(doc.documentElement, 'style');

    for (const styleElem of styleElements) {
      const styleId = this.xml.attr(styleElem, 'styleId');
      if (!styleId) continue;

      const styleInfo: StyleInfo = {
        name: styleId
      };

      // Get style type (paragraph, table, character, numbering)
      const styleType = this.xml.attr(styleElem, 'type');
      if (styleType) {
        styleInfo.type = styleType as any;
      }

      // Check if this is the default style
      const isDefault = this.xml.attr(styleElem, 'default') === '1';
      if (isDefault) {
        styleInfo.isDefault = true;
      }

      // Get style name
      const nameElem = this.xml.element(styleElem, 'name');
      if (nameElem) {
        const displayName = this.xml.attr(nameElem, 'val');
        if (displayName) {
          styleInfo.name = displayName;
        }
      }

      // Get basedOn
      const basedOn = this.xml.element(styleElem, 'basedOn');
      if (basedOn) {
        styleInfo.basedOn = this.xml.attr(basedOn, 'val') ?? undefined;
      }

      // Get next style
      const next = this.xml.element(styleElem, 'next');
      if (next) {
        styleInfo.next = this.xml.attr(next, 'val') ?? undefined;
      }

      // Get linked style
      const link = this.xml.element(styleElem, 'link');
      if (link) {
        styleInfo.link = this.xml.attr(link, 'val') ?? undefined;
      }

      // Get UI priority
      const uiPriority = this.xml.element(styleElem, 'uiPriority');
      if (uiPriority) {
        const priority = this.xml.intAttr(uiPriority, 'val');
        if (priority !== null) styleInfo.uiPriority = priority;
      }

      // Get qFormat
      if (this.xml.element(styleElem, 'qFormat')) {
        styleInfo.qFormat = true;
      }

      // Get unhideWhenUsed
      if (this.xml.element(styleElem, 'unhideWhenUsed')) {
        styleInfo.unhideWhenUsed = true;
      }

      // Extract run properties from style
      const rPr = this.xml.element(styleElem, 'rPr');
      if (rPr) {
        const formatting = this.extractRunFormattingFromElement(rPr);
        if (Object.keys(formatting).length > 0) {
          styleInfo.runFormatting = formatting;
        }
      }

      // Extract paragraph properties from style
      const pPr = this.xml.element(styleElem, 'pPr');
      if (pPr) {
        // Keep with next
        if (this.xml.element(pPr, 'keepNext')) {
          styleInfo.keepNext = true;
        }

        // Keep lines together
        if (this.xml.element(pPr, 'keepLines')) {
          styleInfo.keepLines = true;
        }

        // Outline level
        const outlineLvl = this.xml.element(pPr, 'outlineLvl');
        if (outlineLvl) {
          const level = this.xml.intAttr(outlineLvl, 'val');
          if (level !== null) styleInfo.outlineLevel = level;
        }

        // Numbering (for list styles)
        const numPr = this.xml.element(pPr, 'numPr');
        if (numPr) {
          const ilvl = this.xml.element(numPr, 'ilvl');
          const numId = this.xml.element(numPr, 'numId');

          if (numId) {
            styleInfo.numbering = {
              id: this.xml.attr(numId, 'val') || '',
              // Only include level if ilvl is explicitly present
              level: ilvl ? this.xml.intAttr(ilvl, 'val', 0) : undefined
            };
          }
        }

        // Contextual spacing
        if (this.xml.element(pPr, 'contextualSpacing')) {
          styleInfo.contextualSpacing = true;
        }

        // Spacing
        const spacingElem = this.xml.element(pPr, 'spacing');
        if (spacingElem) {
          styleInfo.spacing = {
            before: this.xml.lengthAttr(spacingElem, 'before') ?? undefined,
            after: this.xml.lengthAttr(spacingElem, 'after') ?? undefined,
            line: this.xml.intAttr(spacingElem, 'line') / 240 || undefined,
            lineRule: this.xml.attr(spacingElem, 'lineRule') as any ?? undefined
          };
        }

        // Alignment
        const jc = this.xml.element(pPr, 'jc');
        if (jc) {
          const jcVal = this.xml.attr(jc, 'val');
          styleInfo.alignment = (jcVal === 'start' ? 'left' : jcVal === 'end' ? 'right' : jcVal) as any;
        }

        // Indentation
        const ind = this.xml.element(pPr, 'ind');
        if (ind) {
          styleInfo.indentation = {
            left: this.xml.lengthAttr(ind, 'left') ?? undefined,
            right: this.xml.lengthAttr(ind, 'right') ?? undefined,
            firstLine: this.xml.lengthAttr(ind, 'firstLine') ?? undefined,
            hanging: this.xml.lengthAttr(ind, 'hanging') ?? undefined
          };
        }
      }

      // Extract table properties (for table styles)
      if (styleType === 'table') {
        const tblPr = this.xml.element(styleElem, 'tblPr');
        if (tblPr) {
          // Serialize the entire tblPr element to XML
          // This preserves borders, cell margins, etc.
          styleInfo.tablePropertiesXml = this.xmlSerializer.serializeToString(tblPr);
        }
      }

      // Store the style (always store, even if it has no explicit formatting)
      // This ensures we preserve all styles including minimal ones like "Normal"
      this.styles.set(styleId, styleInfo);

    }
  }

  private async loadNumbering(): Promise<void> {
    if (!this.zip) return;

    const numberingXmlString = await this.zip.file('word/numbering.xml')?.async('string');
    if (!numberingXmlString) return;

    const doc = this.domParser.parseFromString(numberingXmlString, 'text/xml');
    const numberingElem = doc.documentElement;

    // Parse abstract numbering definitions
    const abstractNums: Map<string, { levels: NumberingLevel[], nsid?: string, multiLevelType?: string, tmpl?: string }> = new Map();
    for (const abstractNum of this.xml.elements(numberingElem, 'abstractNum')) {
      const abstractNumId = this.xml.attr(abstractNum, 'abstractNumId');
      if (!abstractNumId) continue;

      // Extract nsid and tmpl (Word metadata)
      const nsidElem = this.xml.element(abstractNum, 'nsid');
      const nsid = nsidElem ? this.xml.attr(nsidElem, 'val') ?? undefined : undefined;

      // Extract multiLevelType
      const multiLevelTypeElem = this.xml.element(abstractNum, 'multiLevelType');
      const multiLevelType = multiLevelTypeElem ? this.xml.attr(multiLevelTypeElem, 'val') ?? undefined : undefined;

      const tmplElem = this.xml.element(abstractNum, 'tmpl');
      const tmpl = tmplElem ? this.xml.attr(tmplElem, 'val') ?? undefined : undefined;

      const levels: NumberingLevel[] = [];
      for (const lvl of this.xml.elements(abstractNum, 'lvl')) {
        const ilvl = this.xml.intAttr(lvl, 'ilvl', 0);

        // Start value
        const startElem = this.xml.element(lvl, 'start');
        const start = startElem ? this.xml.intAttr(startElem, 'val') : undefined;

        // Number format
        const numFmt = this.xml.element(lvl, 'numFmt');
        const format = numFmt ? (this.xml.attr(numFmt, 'val') || 'bullet') : 'bullet';

        // Level text
        const lvlText = this.xml.element(lvl, 'lvlText');
        const text = lvlText ? (this.xml.attr(lvlText, 'val') || '●') : '●';

        // Paragraph style reference
        const pStyleElem = this.xml.element(lvl, 'pStyle');
        const paragraphStyleName = pStyleElem ? this.xml.attr(pStyleElem, 'val') ?? undefined : undefined;

        // Alignment
        const lvlJc = this.xml.element(lvl, 'lvlJc');
        const alignment = lvlJc ? (this.xml.attr(lvlJc, 'val') as any) : undefined;

        // Paragraph properties
        const pPr = this.xml.element(lvl, 'pPr');
        let indentation: IndentationInfo | undefined;
        let tabs: any[] | undefined;

        if (pPr) {
          // Indentation
          const ind = this.xml.element(pPr, 'ind');
          if (ind) {
            indentation = {
              left: this.xml.lengthAttr(ind, 'left') ?? undefined,
              hanging: this.xml.lengthAttr(ind, 'hanging') ?? undefined
            };
          }

          // Tab stops
          const tabsElem = this.xml.element(pPr, 'tabs');
          if (tabsElem) {
            tabs = [];
            for (const tab of this.xml.elements(tabsElem, 'tab')) {
              const val = this.xml.attr(tab, 'val');
              const pos = this.xml.lengthAttr(tab, 'pos');
              if (val && pos !== null) {
                tabs.push({ val, pos });
              }
            }
            if (tabs.length === 0) tabs = undefined;
          }
        }

        // Run properties
        const rPr = this.xml.element(lvl, 'rPr');
        let fontFamily: string | undefined;
        let fontHint: string | undefined;

        if (rPr) {
          const rFonts = this.xml.element(rPr, 'rFonts');
          if (rFonts) {
            fontFamily = this.xml.attr(rFonts, 'ascii') ?? undefined;
            fontHint = this.xml.attr(rFonts, 'hint') ?? undefined;
          }
        }

        levels.push({
          level: ilvl,
          start,
          format,
          text,
          paragraphStyleName,
          alignment,
          indentation,
          tabs,
          fontFamily,
          fontHint
        });
      }

      abstractNums.set(abstractNumId, { levels, nsid, multiLevelType, tmpl });
    }

    // Parse numbering instances
    for (const num of this.xml.elements(numberingElem, 'num')) {
      const numId = this.xml.attr(num, 'numId');
      if (!numId) continue;

      const abstractNumId = this.xml.element(num, 'abstractNumId');
      const abstractNumIdVal = abstractNumId ? this.xml.attr(abstractNumId, 'val') : null;

      if (abstractNumIdVal && abstractNums.has(abstractNumIdVal)) {
        const abstractData = abstractNums.get(abstractNumIdVal)!;
        this.numbering.set(numId, {
          numId,
          abstractNumId: abstractNumIdVal,
          nsid: abstractData.nsid,
          multiLevelType: abstractData.multiLevelType,
          tmpl: abstractData.tmpl,
          levels: abstractData.levels
        });
      }
    }
  }

  private extractRunFormattingFromElement(rPr: Element): RunFormatting {
    const formatting: RunFormatting = {};

    // Bold
    if (this.xml.element(rPr, 'b')) {
      formatting.bold = this.xml.boolAttr(this.xml.element(rPr, 'b')!, 'val', true);
    }

    // Italic
    if (this.xml.element(rPr, 'i')) {
      formatting.italic = this.xml.boolAttr(this.xml.element(rPr, 'i')!, 'val', true);
    }

    // Underline
    const u = this.xml.element(rPr, 'u');
    if (u && this.xml.attr(u, 'val') !== 'none') {
      formatting.underline = true;
    }

    // Strike
    if (this.xml.element(rPr, 'strike')) {
      formatting.strike = this.xml.boolAttr(this.xml.element(rPr, 'strike')!, 'val', true);
    }

    // Font size (in half-points)
    const sz = this.xml.element(rPr, 'sz');
    if (sz) {
      formatting.fontSize = this.xml.intAttr(sz, 'val') / 2;
    }

    // Font family and theme references
    const rFonts = this.xml.element(rPr, 'rFonts');
    if (rFonts) {
      formatting.fontFamily = this.xml.attr(rFonts, 'ascii') ?? undefined;

      // Extract theme font references
      const asciiTheme = this.xml.attr(rFonts, 'asciiTheme');
      const eastAsiaTheme = this.xml.attr(rFonts, 'eastAsiaTheme');
      const hAnsiTheme = this.xml.attr(rFonts, 'hAnsiTheme');
      const csTheme = this.xml.attr(rFonts, 'cstheme');

      if (asciiTheme) formatting.fontThemeAscii = asciiTheme;
      if (eastAsiaTheme) formatting.fontThemeEastAsia = eastAsiaTheme;
      if (hAnsiTheme) formatting.fontThemeHAnsi = hAnsiTheme;
      if (csTheme) formatting.fontThemeCs = csTheme;
    }

    // Color and theme references
    const color = this.xml.element(rPr, 'color');
    if (color) {
      const colorVal = this.xml.attr(color, 'val');
      if (colorVal && colorVal !== 'auto') {
        formatting.color = `#${colorVal}`;
      }

      // Extract theme color references
      const themeColor = this.xml.attr(color, 'themeColor');
      const themeShade = this.xml.attr(color, 'themeShade');
      const themeTint = this.xml.attr(color, 'themeTint');

      if (themeColor) formatting.colorTheme = themeColor;
      if (themeShade) formatting.colorThemeShade = themeShade;
      else if (themeTint) formatting.colorThemeShade = themeTint; // tint can also go in the same field
    }

    // Highlight
    const highlight = this.xml.element(rPr, 'highlight');
    if (highlight) {
      formatting.highlight = this.xml.attr(highlight, 'val') ?? undefined;
    }

    return formatting;
  }

  private async extractParagraph(pElem: Element): Promise<ExtractedParagraph> {
    const runs: ExtractedRun[] = [];
    let spacing: SpacingInfo | undefined;
    let alignment: 'left' | 'center' | 'right' | 'justify' | undefined;
    let indentation: IndentationInfo | undefined;
    let styleName: string | undefined;
    let numbering: NumberingInfo | undefined;

    // Parse paragraph properties
    const pPr = this.xml.element(pElem, 'pPr');
    if (pPr) {
      // Spacing
      const spacingElem = this.xml.element(pPr, 'spacing');
      if (spacingElem) {
        spacing = {
          before: this.xml.lengthAttr(spacingElem, 'before') ?? undefined,
          after: this.xml.lengthAttr(spacingElem, 'after') ?? undefined,
          line: this.xml.intAttr(spacingElem, 'line') / 240 || undefined, // Convert to line height multiplier
          lineRule: this.xml.attr(spacingElem, 'lineRule') as any ?? undefined
        };
      }

      // Alignment
      const jc = this.xml.element(pPr, 'jc');
      if (jc) {
        const jcVal = this.xml.attr(jc, 'val');
        alignment = (jcVal === 'start' ? 'left' : jcVal === 'end' ? 'right' : jcVal) as any;
      }

      // Indentation
      const ind = this.xml.element(pPr, 'ind');
      if (ind) {
        indentation = {
          left: this.xml.lengthAttr(ind, 'left') ?? undefined,
          right: this.xml.lengthAttr(ind, 'right') ?? undefined,
          firstLine: this.xml.lengthAttr(ind, 'firstLine') ?? undefined,
          hanging: this.xml.lengthAttr(ind, 'hanging') ?? undefined
        };
      }

      // Style
      const pStyle = this.xml.element(pPr, 'pStyle');
      if (pStyle) {
        styleName = this.xml.attr(pStyle, 'val') ?? undefined;
      }

      // Numbering
      const numPr = this.xml.element(pPr, 'numPr');
      if (numPr) {
        const numId = this.xml.element(numPr, 'numId');
        const ilvl = this.xml.element(numPr, 'ilvl');
        if (numId) {
          numbering = {
            id: this.xml.attr(numId, 'val') || '',
            level: ilvl ? this.xml.intAttr(ilvl, 'val') : 0
          };
        }
      }
    }

    // Extract runs
    for (const child of this.xml.elements(pElem)) {
      if (child.localName === 'r') {
        const run = await this.extractRun(child);
        if (run) {
          runs.push(run);
        }
      }
    }

    const isEmpty = runs.length === 0 || runs.every(r => !r.text || r.text.trim().length === 0);

    // Note: We do NOT merge style formatting into runs or paragraph properties
    // This preserves the original document structure where formatting comes from styles
    // Word will apply style formatting automatically when rendering

    return {
      runs,
      spacing,
      alignment,
      indentation,
      styleName,
      numbering,
      isEmpty
    };
  }

  private async extractRun(rElem: Element): Promise<ExtractedRun | null> {
    let text = '';
    const formatting: RunFormatting = {};

    // Parse run properties
    const rPr = this.xml.element(rElem, 'rPr');
    if (rPr) {
      // Bold
      if (this.xml.element(rPr, 'b')) {
        formatting.bold = this.xml.boolAttr(this.xml.element(rPr, 'b')!, 'val', true);
      }

      // Italic
      if (this.xml.element(rPr, 'i')) {
        formatting.italic = this.xml.boolAttr(this.xml.element(rPr, 'i')!, 'val', true);
      }

      // Underline
      const u = this.xml.element(rPr, 'u');
      if (u && this.xml.attr(u, 'val') !== 'none') {
        formatting.underline = true;
      }

      // Strike
      if (this.xml.element(rPr, 'strike')) {
        formatting.strike = this.xml.boolAttr(this.xml.element(rPr, 'strike')!, 'val', true);
      }

      // Font size (in half-points)
      const sz = this.xml.element(rPr, 'sz');
      if (sz) {
        formatting.fontSize = this.xml.intAttr(sz, 'val') / 2; // Convert half-points to points
      }

      // Font family
      const rFonts = this.xml.element(rPr, 'rFonts');
      if (rFonts) {
        formatting.fontFamily = this.xml.attr(rFonts, 'ascii') ?? undefined;
      }

      // Color
      const color = this.xml.element(rPr, 'color');
      if (color) {
        const colorVal = this.xml.attr(color, 'val');
        if (colorVal && colorVal !== 'auto') {
          formatting.color = `#${colorVal}`;
        }
      }

      // Highlight
      const highlight = this.xml.element(rPr, 'highlight');
      if (highlight) {
        formatting.highlight = this.xml.attr(highlight, 'val') ?? undefined;
      }

      // Vertical alignment (superscript/subscript)
      const vertAlign = this.xml.element(rPr, 'vertAlign');
      if (vertAlign) {
        const val = this.xml.attr(vertAlign, 'val');
        if (val === 'superscript' || val === 'subscript') {
          formatting.verticalAlign = val;
        }
      }
    }

    // Extract text content and images
    let image: ExtractedImage | undefined;

    for (const child of this.xml.elements(rElem)) {
      const localName = child.localName;

      switch (localName) {
        case 't': // text
          text += child.textContent || '';
          break;
        case 'tab': // tab character
          text += '\t';
          break;
        case 'br': // line break
          // Check if it's a page break
          const type = this.xml.attr(child, 'type');
          if (type === 'page') {
            text += '\f'; // form feed for page break
          } else {
            text += '\n';
          }
          break;
        case 'drawing': // image/drawing
          image = await this.extractDrawing(child);
          break;
        case 'pict': // VML picture
          image = await this.extractVmlPicture(child);
          break;
      }
    }

    // Handle special line terminators (U+2028, U+2029)
    // These should be replaced with regular spaces for Univer
    text = text
      .replace(/\u2028/g, ' ') // Line Separator
      .replace(/\u2029/g, ' ') // Paragraph Separator
      .replace(/\u0085/g, ' ') // Next Line
      .replace(/\u000B/g, ' ') // Vertical Tab
      .replace(/\u000C/g, ' '); // Form Feed

    return {
      text: text || undefined,
      formatting: Object.keys(formatting).length > 0 ? formatting : undefined,
      image
    };
  }

  private async extractDrawing(drawingElem: Element): Promise<ExtractedImage | undefined> {
    // Look for inline or anchor
    const inline = this.xml.element(drawingElem, 'inline');
    const anchor = this.xml.element(drawingElem, 'anchor');
    const wrapper = inline || anchor;

    if (!wrapper) return undefined;

    let width: number | undefined;
    let height: number | undefined;

    // Get extent (size)
    const extent = this.xml.element(wrapper, 'extent');
    if (extent) {
      const cx = this.xml.lengthAttr(extent, 'cx', 'emu');
      const cy = this.xml.lengthAttr(extent, 'cy', 'emu');
      if (cx !== null) width = cx;
      if (cy !== null) height = cy;
    }

    // Get graphic data
    const graphic = this.xml.element(wrapper, 'graphic');
    if (!graphic) return undefined;

    const graphicData = this.xml.element(graphic, 'graphicData');
    if (!graphicData) return undefined;

    const pic = this.xml.element(graphicData, 'pic');
    if (!pic) return undefined;

    // Get blipFill for image reference
    const blipFill = this.xml.element(pic, 'blipFill');
    if (!blipFill) return undefined;

    const blip = this.xml.element(blipFill, 'blip');
    if (!blip) return undefined;

    // Get the relationship ID (r:embed)
    const embedId = this.xml.attr(blip, 'embed');
    if (!embedId) return undefined;

    // Get image path from relationships
    const imagePath = this.relationships.get(embedId);
    if (!imagePath) return undefined;

    // Full media path for preservation
    const fullMediaPath = 'word/' + imagePath;

    // Load image data
    const imageData = await this.loadImageData(fullMediaPath);

    return {
      src: embedId,
      mediaPath: fullMediaPath,
      width,
      height,
      ...imageData
    };
  }

  private async extractVmlPicture(pictElem: Element): Promise<ExtractedImage | undefined> {
    // VML pictures (older format) - simplified extraction
    // TODO: Implement full VML picture parsing if needed
    return undefined;
  }

  private async loadImageData(path: string): Promise<{ data?: string; contentType?: string }> {
    if (!this.zip) return {};

    const imageFile = this.zip.file(path);
    if (!imageFile) return {};

    try {
      const arrayBuffer = await imageFile.async('arraybuffer');
      const uint8Array = new Uint8Array(arrayBuffer);

      // Convert to base64
      const base64 = Buffer.from(uint8Array).toString('base64');

      // Determine content type from extension
      const ext = path.split('.').pop()?.toLowerCase();
      const contentTypeMap: Record<string, string> = {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'bmp': 'image/bmp',
        'svg': 'image/svg+xml',
        'webp': 'image/webp'
      };

      return {
        data: base64,
        contentType: ext ? contentTypeMap[ext] : undefined
      };
    } catch (error) {
      console.error('Error loading image:', path, error);
      return {};
    }
  }

  private async extractTable(tblElem: Element): Promise<ExtractedTable> {
    const rows: ExtractedTableRow[] = [];

    for (const tr of this.xml.elements(tblElem, 'tr')) {
      const cells: ExtractedTableCell[] = [];

      for (const tc of this.xml.elements(tr, 'tc')) {
        const content: ExtractedParagraph[] = [];

        // Extract paragraphs in cell
        for (const p of this.xml.elements(tc, 'p')) {
          content.push(await this.extractParagraph(p));
        }

        // Extract cell properties
        const cellProps = this.extractCellProperties(tc);

        cells.push({
          content,
          ...cellProps
        });
      }

      rows.push({ cells });
    }

    return { rows };
  }

  private extractCellProperties(tcElem: Element): Partial<ExtractedTableCell> {
    const props: Partial<ExtractedTableCell> = {};

    const tcPr = this.xml.element(tcElem, 'tcPr');
    if (!tcPr) return props;

    // Grid span (colspan)
    const gridSpan = this.xml.element(tcPr, 'gridSpan');
    if (gridSpan) {
      const val = this.xml.intAttr(gridSpan, 'val');
      if (val > 1) props.colSpan = val;
    }

    // Vertical merge (rowspan) - this is complex, handled by vMerge attribute
    const vMerge = this.xml.element(tcPr, 'vMerge');
    if (vMerge) {
      const val = this.xml.attr(vMerge, 'val');
      // 'restart' = start of merged region, no val = continuation
      // Note: rowSpan calculation requires tracking state across rows
      // For now, just mark it as a merged cell
      if (val === 'restart' || val === null) {
        props.rowSpan = val === 'restart' ? 1 : 0; // 0 = continuation cell
      }
    }

    // Cell width
    const tcW = this.xml.element(tcPr, 'tcW');
    if (tcW) {
      const width = this.xml.lengthAttr(tcW, 'w');
      if (width !== null) props.width = width;
    }

    // Cell shading (background color)
    const shd = this.xml.element(tcPr, 'shd');
    if (shd) {
      const fill = this.xml.attr(shd, 'fill');
      if (fill && fill !== 'auto') {
        props.backgroundColor = '#' + fill;
      }
    }

    // Vertical alignment
    const vAlign = this.xml.element(tcPr, 'vAlign');
    if (vAlign) {
      const val = this.xml.attr(vAlign, 'val');
      if (val === 'top' || val === 'center' || val === 'bottom') {
        props.verticalAlign = val;
      }
    }

    // Cell borders
    const tcBorders = this.xml.element(tcPr, 'tcBorders');
    if (tcBorders) {
      const borders: any = {};

      for (const side of ['top', 'bottom', 'left', 'right']) {
        const borderElem = this.xml.element(tcBorders, side);
        if (borderElem) {
          const borderInfo: any = {};

          const style = this.xml.attr(borderElem, 'val');
          if (style && style !== 'none') {
            borderInfo.style = style;
          }

          const size = this.xml.intAttr(borderElem, 'sz');
          if (size > 0) {
            borderInfo.size = size;
          }

          const color = this.xml.attr(borderElem, 'color');
          if (color && color !== 'auto') {
            borderInfo.color = '#' + color;
          }

          if (Object.keys(borderInfo).length > 0) {
            borders[side] = borderInfo;
          }
        }
      }

      if (Object.keys(borders).length > 0) {
        props.borders = borders;
      }
    }

    // Cell margins
    const tcMar = this.xml.element(tcPr, 'tcMar');
    if (tcMar) {
      const margins: any = {};

      for (const side of ['top', 'bottom', 'left', 'right']) {
        const marginElem = this.xml.element(tcMar, side);
        if (marginElem) {
          const margin = this.xml.lengthAttr(marginElem, 'w');
          if (margin !== null) {
            margins[side] = margin;
          }
        }
      }

      if (Object.keys(margins).length > 0) {
        props.margins = margins;
      }
    }

    return props;
  }
}

