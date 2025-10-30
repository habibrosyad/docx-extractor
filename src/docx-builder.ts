/**
 * DOCX Builder - Converts extracted document data back to DOCX format
 */

import JSZip from 'jszip';
import {
  ExtractedDocument,
  ExtractedParagraph,
  ExtractedRun,
  ExtractedTable,
  ExtractedTableCell,
  RunFormatting,
  SpacingInfo,
  IndentationInfo,
} from './types.js';

export class DocxBuilder {
  private relationships: Map<string, { id: string; type: string; target: string }> = new Map();
  private nextRelId = 1;

  async build(document: ExtractedDocument): Promise<Uint8Array> {
    const zip = new JSZip();

    // Create required DOCX structure
    this.addContentTypes(zip);
    this.addRels(zip);
    this.addDocumentProps(zip);
    this.addDocumentRels(zip, document);
    this.addSettings(zip);
    this.addWebSettings(zip);
    this.addStyles(zip, document);
    this.addNumbering(zip, document);
    this.addDocument(zip, document);

    // Add any embedded images
    await this.addImages(zip, document);

    // Generate the DOCX file (use uint8array for platform compatibility)
    const buffer = await zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 9 }
    });

    return buffer;
  }

  private addContentTypes(zip: JSZip): void {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
    zip.file('[Content_Types].xml', xml);
  }

  private addRels(zip: JSZip): void {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
    zip.file('_rels/.rels', xml);
  }

  private addDocumentProps(zip: JSZip): void {
    // Get current date in ISO 8601 format (W3CDTF)
    const now = new Date();
    const isoDate = now.toISOString();

    // Add core.xml (required metadata)
    const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title/>
  <dc:subject/>
  <dc:creator/>
  <cp:keywords/>
  <dc:description/>
  <cp:lastModifiedBy/>
  <cp:revision>1</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">${isoDate}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${isoDate}</dcterms:modified>
</cp:coreProperties>`;
    zip.file('docProps/core.xml', coreXml);

    // Add app.xml (application metadata)
    const appXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <TotalTime>0</TotalTime>
  <Words>0</Words>
  <Characters>0</Characters>
  <Paragraphs>0</Paragraphs>
  <Lines>0</Lines>
  <Pages>1</Pages>
</Properties>`;
    zip.file('docProps/app.xml', appXml);
  }

  private addSettings(zip: JSZip): void {
    // Add minimal settings.xml (required by Word)
    const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:defaultTabStop w:val="720"/>
</w:settings>`;
    zip.file('word/settings.xml', settingsXml);
  }

  private addWebSettings(zip: JSZip): void {
    // Add minimal webSettings.xml (required by Word)
    const webSettingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:optimizeForBrowser/>
</w:webSettings>`;
    zip.file('word/webSettings.xml', webSettingsXml);
  }

  private addDocumentRels(zip: JSZip, document: ExtractedDocument): void {
    // Build relationships from extracted data AND preserved media files
    let relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>`;

    // Reset nextRelId to start from 5 (after styles, numbering, settings, webSettings)
    this.nextRelId = 5;

    // Add relationships for preserved media files (images, charts, etc.)
    if (document.mediaFiles && document.mediaFiles.size > 0) {
      for (const filePath of document.mediaFiles.keys()) {
        // Extract the relative path (e.g., word/media/image1.png -> media/image1.png)
        const target = filePath.replace('word/', '');
        const relId = `rId${this.nextRelId++}`;

        // Determine relationship type based on file extension
        const ext = filePath.split('.').pop()?.toLowerCase();
        let relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
        if (ext === 'xml' && filePath.includes('/charts/')) {
          relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
        }

        relsXml += `\n  <Relationship Id="${relId}" Type="${relType}" Target="${target}"/>`;
        this.relationships.set(filePath, { id: relId, type: 'image', target });
      }
    }

    // Also add relationships for any NEW images (not from preserved mediaFiles)
    // Only add if image has data but no mediaPath (meaning it's a new image, not from original doc)
    let imageCount = 0;
    for (const para of document.paragraphs) {
      for (const run of para.runs) {
        if (run.image?.data && !run.image.mediaPath) {
          // This is a NEW image (not preserved from original)
          imageCount++;
          const relId = `rId${this.nextRelId++}`;
          const ext = run.image.contentType?.split('/')[1] || 'png';
          const target = `media/image${imageCount}.${ext}`;
          relsXml += `\n  <Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/>`;
          this.relationships.set(run.image.src, { id: relId, type: 'image', target });
        }
      }
    }

    relsXml += '\n</Relationships>';
    zip.file('word/_rels/document.xml.rels', relsXml);
  }

  private addStyles(zip: JSZip, document: ExtractedDocument): void {
    // Build styles.xml from structured data
    let stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`;

    // Add document defaults
    const hasRunDefaults = document.defaults && Object.keys(document.defaults).length > 0;
    const hasParagraphDefaults = document.paragraphDefaults && Object.keys(document.paragraphDefaults).length > 0;

    if (hasRunDefaults || hasParagraphDefaults) {
      stylesXml += '\n  <w:docDefaults>';

      // Add rPrDefault (run properties default)
      if (hasRunDefaults) {
        stylesXml += '\n    <w:rPrDefault>\n      <w:rPr>';

        // Font theme references (must come before fontSize)
        if (document.defaults!.fontThemeAscii || document.defaults!.fontThemeEastAsia || document.defaults!.fontThemeHAnsi || document.defaults!.fontThemeCs) {
          stylesXml += '\n        <w:rFonts';
          if (document.defaults!.fontThemeAscii) stylesXml += ` w:asciiTheme="${document.defaults!.fontThemeAscii}"`;
          if (document.defaults!.fontThemeEastAsia) stylesXml += ` w:eastAsiaTheme="${document.defaults!.fontThemeEastAsia}"`;
          if (document.defaults!.fontThemeHAnsi) stylesXml += ` w:hAnsiTheme="${document.defaults!.fontThemeHAnsi}"`;
          if (document.defaults!.fontThemeCs) stylesXml += ` w:cstheme="${document.defaults!.fontThemeCs}"`;
          stylesXml += '/>';
        }

        if (document.defaults!.fontSize) {
          stylesXml += `\n        <w:sz w:val="${document.defaults!.fontSize * 2}"/>`;
          stylesXml += `\n        <w:szCs w:val="${document.defaults!.fontSize * 2}"/>`;
        }
        if (document.defaults!.fontFamily) {
          stylesXml += `\n        <w:rFonts w:ascii="${document.defaults!.fontFamily}" w:hAnsi="${document.defaults!.fontFamily}"/>`;
        }
        stylesXml += '\n      </w:rPr>\n    </w:rPrDefault>';
      }

      // Add pPrDefault (paragraph properties default) - CRITICAL FOR SPACING!
      if (hasParagraphDefaults) {
        stylesXml += '\n    <w:pPrDefault>\n      <w:pPr>';

        if (document.paragraphDefaults!.spacing) {
          stylesXml += this.buildSpacingXml(document.paragraphDefaults!.spacing);
        }

        if (document.paragraphDefaults!.alignment) {
          stylesXml += `\n        <w:jc w:val="${document.paragraphDefaults!.alignment}"/>`;
        }

        if (document.paragraphDefaults!.indentation) {
          stylesXml += this.buildIndentationXml(document.paragraphDefaults!.indentation);
        }

        stylesXml += '\n      </w:pPr>\n    </w:pPrDefault>';
      }

      stylesXml += '\n  </w:docDefaults>';
    }

    // Collect table styles used in the document
    const usedTableStyles = new Set<string>();
    for (const table of document.tables) {
      if (table.tableStyle) {
        usedTableStyles.add(table.tableStyle);
      }
    }

    // Add styles
    if (document.styles && document.styles.size > 0) {
      for (const [styleId, styleInfo] of document.styles) {
        // Add default attribute if this is the default style
        const defaultAttr = styleInfo.isDefault ? ' w:default="1"' : '';
        // Use the style type from extracted data, default to 'paragraph'
        const styleType = styleInfo.type || 'paragraph';
        stylesXml += `\n  <w:style w:type="${styleType}" w:styleId="${styleId}"${defaultAttr}>`;
        stylesXml += `\n    <w:name w:val="${styleInfo.name}"/>`;

        // Mark this style as found if it's a used table style
        if (styleType === 'table' && usedTableStyles.has(styleId)) {
          usedTableStyles.delete(styleId);
        }

        if (styleInfo.basedOn) {
          stylesXml += `\n    <w:basedOn w:val="${styleInfo.basedOn}"/>`;
        }

        if (styleInfo.next) {
          stylesXml += `\n    <w:next w:val="${styleInfo.next}"/>`;
        }

        if (styleInfo.link) {
          stylesXml += `\n    <w:link w:val="${styleInfo.link}"/>`;
        }

        if (styleInfo.uiPriority !== undefined) {
          stylesXml += `\n    <w:uiPriority w:val="${styleInfo.uiPriority}"/>`;
        }

        if (styleInfo.qFormat) {
          stylesXml += '\n    <w:qFormat/>';
        }

        if (styleInfo.unhideWhenUsed) {
          stylesXml += '\n    <w:unhideWhenUsed/>';
        }

        // Paragraph properties
        if (styleInfo.spacing || styleInfo.alignment || styleInfo.indentation || styleInfo.keepNext || styleInfo.keepLines || styleInfo.outlineLevel !== undefined || styleInfo.numbering || styleInfo.contextualSpacing) {
          stylesXml += '\n    <w:pPr>';

          // Numbering (for list styles)
          if (styleInfo.numbering) {
            stylesXml += '\n      <w:numPr>';
            // Only output ilvl if it was explicitly present in the original
            if (styleInfo.numbering.level !== undefined) {
              stylesXml += `\n        <w:ilvl w:val="${styleInfo.numbering.level}"/>`;
            }
            stylesXml += `\n        <w:numId w:val="${styleInfo.numbering.id}"/>`;
            stylesXml += '\n      </w:numPr>';
          }

          // Contextual spacing
          if (styleInfo.contextualSpacing) {
            stylesXml += '\n      <w:contextualSpacing/>';
          }

          if (styleInfo.keepNext) {
            stylesXml += '\n      <w:keepNext/>';
          }

          if (styleInfo.keepLines) {
            stylesXml += '\n      <w:keepLines/>';
          }

          if (styleInfo.spacing) {
            stylesXml += this.buildSpacingXml(styleInfo.spacing);
          }

          if (styleInfo.outlineLevel !== undefined) {
            stylesXml += `\n      <w:outlineLvl w:val="${styleInfo.outlineLevel}"/>`;
          }

          if (styleInfo.alignment) {
            stylesXml += `\n      <w:jc w:val="${styleInfo.alignment}"/>`;
          }

          if (styleInfo.indentation) {
            stylesXml += this.buildIndentationXml(styleInfo.indentation);
          }

          stylesXml += '\n    </w:pPr>';
        }

        // Run properties
        if (styleInfo.runFormatting) {
          stylesXml += '\n    <w:rPr>';
          stylesXml += this.buildRunFormattingXml(styleInfo.runFormatting);
          stylesXml += '\n    </w:rPr>';
        }

        // Table properties (for table styles)
        if (styleType === 'table' && styleInfo.tablePropertiesXml) {
          // Insert the preserved table properties XML
          // The XML should already have correct namespace prefixes from serializeToString
          // Just ensure any stray ns0:, ns1: prefixes are replaced with w:, but preserve existing w: prefixes
          let tblPrXml = styleInfo.tablePropertiesXml;
          // Only replace ns\d+: prefixes, not w: prefixes
          tblPrXml = tblPrXml.replace(/<(ns\d+):/g, '<w:');
          tblPrXml = tblPrXml.replace(/<\/(ns\d+):/g, '</w:');
          // Also ensure elements without any prefix get w: prefix (except they should already have it from serializeToString)
          // But handle case where localName elements might not have prefix
          // Replace unqualified element tags that should be w: namespace
          const tableElements = ['tblPr', 'tblBorders', 'tblInd', 'top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'tblLook'];
          for (const elem of tableElements) {
            // Only add prefix if element doesn't already have w: or other prefix
            tblPrXml = tblPrXml.replace(new RegExp(`<(?!/w:)(?!ns\\d+:)(?![a-z-]+:)${elem}(\\s|>)`, 'g'), `<w:${elem}$1`);
            tblPrXml = tblPrXml.replace(new RegExp(`</(?!/w:)(?!ns\\d+:)(?![a-z-]+:)${elem}>`, 'g'), `</w:${elem}>`);
          }
          tblPrXml = tblPrXml.replace(/xmlns:ns\d+="[^"]*"/g, '');
          // Indent properly
          tblPrXml = '\n    ' + tblPrXml.trim().replace(/\n/g, '\n    ');
          stylesXml += tblPrXml;
        }

        stylesXml += '\n  </w:style>';
      }
    }

    // Add any missing table styles that are referenced but not found
    // This ensures tables can render even if their style wasn't fully extracted
    for (const tableStyle of usedTableStyles) {
      // Add minimal table style definition so Word can render the table
      stylesXml += `\n  <w:style w:type="table" w:styleId="${tableStyle}">`;
      stylesXml += `\n    <w:name w:val="${tableStyle}"/>`;
      stylesXml += `\n    <w:tblPr>
      <w:tblStyleRowBandSize w:val="1"/>
      <w:tblStyleColBandSize w:val="1"/>
    </w:tblPr>`;
      stylesXml += '\n  </w:style>';
    }

    stylesXml += '\n</w:styles>';

    zip.file('word/styles.xml', stylesXml);
  }

  private addNumbering(zip: JSZip, document: ExtractedDocument): void {
    // Only generate numbering.xml if there are actual numbering definitions or paragraphs with numbering
    const hasNumbering = document.numbering && document.numbering.size > 0;
    const hasNumberedParagraphs = document.paragraphs.some(p => p.numbering);

    if (!hasNumbering && !hasNumberedParagraphs) {
      // Generate minimal empty numbering.xml
      const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;
      zip.file('word/numbering.xml', numberingXml);
      return;
    }

    // Build numbering.xml from structured data or use default
    let numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`;

    if (hasNumbering && document.numbering) {
      // Build from extracted numbering definitions
      const abstractNums = new Set<string>();

      // First, add abstract numbering definitions
      for (const [numId, numDef] of document.numbering) {
        if (!abstractNums.has(numDef.abstractNumId)) {
          abstractNums.add(numDef.abstractNumId);

          numberingXml += `\n  <w:abstractNum w:abstractNumId="${numDef.abstractNumId}">`;

          // Add nsid if present (Word metadata)
          if (numDef.nsid) {
            numberingXml += `\n    <w:nsid w:val="${numDef.nsid}"/>`;
          }

          // Add multiLevelType (use extracted value or default to hybridMultilevel)
          const multiLevelType = numDef.multiLevelType || 'hybridMultilevel';
          numberingXml += `\n    <w:multiLevelType w:val="${multiLevelType}"/>`;

          // Add tmpl if present (Word metadata)
          if (numDef.tmpl) {
            numberingXml += `\n    <w:tmpl w:val="${numDef.tmpl}"/>`;
          }

          // Iterate over the Map entries
          for (const [levelNum, level] of numDef.levels.entries()) {
            numberingXml += `\n    <w:lvl w:ilvl="${levelNum}">`;

            // Start value
            numberingXml += `\n      <w:start w:val="${level.start ?? 1}"/>`;

            // Number format
            numberingXml += `\n      <w:numFmt w:val="${level.format ?? 'bullet'}"/>`;

            // Paragraph style reference (important for spacing!)
            if (level.paragraphStyleName) {
              numberingXml += `\n      <w:pStyle w:val="${level.paragraphStyleName}"/>`;
            }

            // Level text
            numberingXml += `\n      <w:lvlText w:val="${level.text ?? '●'}"/>`;

            // Alignment
            if (level.alignment) {
              numberingXml += `\n      <w:lvlJc w:val="${level.alignment}"/>`;
            }

            // Paragraph properties (indentation and tabs)
            if (level.indentation || level.tabs) {
              numberingXml += '\n      <w:pPr>';

              // Tab stops
              if (level.tabs && level.tabs.length > 0) {
                numberingXml += '\n        <w:tabs>';
                for (const tab of level.tabs) {
                  numberingXml += `\n          <w:tab w:val="${tab.val}" w:pos="${Math.round(tab.pos * 20)}"/>`;
                }
                numberingXml += '\n        </w:tabs>';
              }

              // Indentation
              if (level.indentation) {
                numberingXml += '\n        <w:ind';
                if (level.indentation.left !== undefined) {
                  numberingXml += ` w:left="${Math.round(level.indentation.left * 20)}"`;
                }
                if (level.indentation.hanging !== undefined) {
                  numberingXml += ` w:hanging="${Math.round(level.indentation.hanging * 20)}"`;
                }
                numberingXml += '/>';
              }

              numberingXml += '\n      </w:pPr>';
            }

            // Run properties (font)
            if (level.fontFamily || level.fontHint) {
              numberingXml += '\n      <w:rPr>';
              numberingXml += '\n        <w:rFonts';
              if (level.fontFamily) {
                numberingXml += ` w:ascii="${level.fontFamily}" w:hAnsi="${level.fontFamily}"`;
              }
              if (level.fontHint) {
                numberingXml += ` w:hint="${level.fontHint}"`;
              }
              numberingXml += '/>';
              numberingXml += '\n      </w:rPr>';
            }

            numberingXml += '\n    </w:lvl>';
          }

          numberingXml += '\n  </w:abstractNum>';
        }
      }

      // Then, add numbering instances
      if (document.numbering) {
        for (const [numId, numDef] of document.numbering) {
          numberingXml += `\n  <w:num w:numId="${numId}">`;
          numberingXml += `\n    <w:abstractNumId w:val="${numDef.abstractNumId}"/>`;
          numberingXml += '\n  </w:num>';
        }
      }
    } else {
      // Fallback: Create generic numbering for documents without extracted definitions
      numberingXml += `
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="●"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>`;
    }

    numberingXml += '\n</w:numbering>';

    zip.file('word/numbering.xml', numberingXml);
  }

  private addDocument(zip: JSZip, document: ExtractedDocument): void {
    // Build document.xml from structured data (enables modifications)
    let docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>`;

    // Add body elements in document order
    for (const element of document.body) {
      if (element.type === 'paragraph') {
        docXml += this.buildParagraphXml(element.data);
      } else if (element.type === 'table') {
        docXml += this.buildTableXml(element.data);
      }
    }

    // Add section properties (page layout)
    docXml += `
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>`;

    zip.file('word/document.xml', docXml);
  }

  private async addImages(zip: JSZip, document: ExtractedDocument): Promise<void> {
    // Use preserved media files if available, or extract from document data
    if (document.mediaFiles && document.mediaFiles.size > 0) {
      // Copy all preserved media files
      for (const [filePath, content] of document.mediaFiles.entries()) {
        zip.file(filePath, content);
      }
    } else {
      // Fallback: Extract images from paragraph data
      let imageCount = 0;
      for (const para of document.paragraphs) {
        for (const run of para.runs) {
          if (run.image?.data) {
            imageCount++;
            const ext = run.image.contentType?.split('/')[1] || 'png';
            // Decode base64 to binary in a platform-agnostic way
            const binaryString = atob(run.image.data);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
              bytes[i] = binaryString.charCodeAt(i);
            }
            zip.file(`word/media/image${imageCount}.${ext}`, bytes);
          }
        }
      }
    }
  }

  private buildParagraphXml(para: ExtractedParagraph): string {
    let xml = '\n    <w:p>';

    // Paragraph properties
    if (para.spacing || para.alignment || para.indentation || para.styleName || para.numbering) {
      xml += '\n      <w:pPr>';

      if (para.styleName) {
        xml += `\n        <w:pStyle w:val="${para.styleName}"/>`;
      }

      // Add numbering properties ONLY if explicitly set on the paragraph
      // (not if it comes from the style - Word will apply that automatically)
      if (para.numbering) {
        xml += '\n        <w:numPr>';
        // Only output ilvl if it's explicitly defined
        if (para.numbering.level !== undefined) {
          xml += `\n          <w:ilvl w:val="${para.numbering.level}"/>`;
        }
        xml += `\n          <w:numId w:val="${para.numbering.id}"/>`;
        xml += '\n        </w:numPr>';
      }

      if (para.spacing) {
        xml += this.buildSpacingXml(para.spacing);
      }

      if (para.alignment) {
        xml += `\n        <w:jc w:val="${para.alignment}"/>`;
      }

      if (para.indentation) {
        xml += this.buildIndentationXml(para.indentation);
      }

      xml += '\n      </w:pPr>';
    }

    // Runs
    // CRITICAL: Every paragraph MUST have at least one run, even if empty
    // Word will not render paragraphs (especially in tables) without at least one <w:r> element
    if (para.runs && para.runs.length > 0) {
      for (const run of para.runs) {
        xml += this.buildRunXml(run);
      }
    } else {
      // Add empty run if paragraph has no runs
      xml += '\n      <w:r>';
      xml += '\n        <w:t></w:t>';
      xml += '\n      </w:r>';
    }

    xml += '\n    </w:p>';
    return xml;
  }

  private buildRunXml(run: ExtractedRun): string {
    let xml = '\n      <w:r>';

    // Run properties
    if (run.formatting && Object.keys(run.formatting).length > 0) {
      xml += '\n        <w:rPr>';
      xml += this.buildRunFormattingXml(run.formatting);
      xml += '\n        </w:rPr>';
    }

    // Text or image
    // CRITICAL: Every run MUST have either text or image content
    // Word will not render runs without at least one content element
    if (run.image) {
      xml += this.buildImageXml(run.image);
    } else if (run.text !== undefined && run.text !== null && run.text !== '') {
      // Handle line breaks and tabs
      const parts = run.text.split('\n');

      for (let i = 0; i < parts.length; i++) {
        if (i > 0) {
          // Add line break element between parts
          xml += '\n        <w:br/>';
        }

        if (parts[i]) {
          // Handle tabs
          const tabParts = parts[i].split('\t');

          for (let j = 0; j < tabParts.length; j++) {
            if (j > 0) {
              // Add tab element
              xml += '\n        <w:tab/>';
            }

            if (tabParts[j]) {
              // Escape XML special characters
              const escapedText = tabParts[j]
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;');
              xml += `\n        <w:t xml:space="preserve">${escapedText}</w:t>`;
            }
          }
        }
      }
    } else {
      // Empty run - must have at least an empty text element for Word to render it
      xml += '\n        <w:t></w:t>';
    }

    xml += '\n      </w:r>';
    return xml;
  }

  private buildImageXml(image: any): string {
    // Look up relationship by mediaPath (preferred) or fallback to src
    const lookupKey = image.mediaPath || image.src;
    const rel = this.relationships.get(lookupKey);

    if (!rel) {
      console.warn(`Relationship for image ${lookupKey} not found. Available keys:`, Array.from(this.relationships.keys()).slice(0, 5));
      return '';
    }

    // If we have preserved drawing XML, use it with updated relationship ID
    if (image.drawingXml) {
      // Replace the old relationship ID with the new one
      let drawingXml = image.drawingXml;
      // Replace r:embed="oldId" with r:embed="newId"
      drawingXml = drawingXml.replace(/r:embed="[^"]*"/g, `r:embed="${rel.id}"`);
      // The drawingXml already includes the <drawing> wrapper from serialization
      return `\n        ${drawingXml}`;
    }

    // Fallback: generate inline image (for compatibility with old extracted data)
    const width = image.width || 100;
    const height = image.height || 100;
    const cx = Math.round(width * 12700); // Convert points to EMUs (1 point = 12700 EMUs)
    const cy = Math.round(height * 12700);

    // Generate unique IDs for this image
    const picId = this.nextRelId++;
    const docPrId = this.nextRelId++;

    // Extract image name from mediaPath or use default
    const imageName = image.mediaPath ? image.mediaPath.split('/').pop() : 'image.png';

    return `\n        <w:drawing>
          <wp:inline xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <wp:extent cx="${cx}" cy="${cy}"/>
            <wp:docPr id="${docPrId}" name="Picture ${picId}"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks noChangeAspect="1"/>
            </wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="${picId}" name="${imageName}"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="${rel.id}"/>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="${cx}" cy="${cy}"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect"/>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>`;
  }

  private buildTableXml(table: ExtractedTable): string {
    let xml = '\n    <w:tbl>';

    // Table properties
    xml += '\n      <w:tblPr>';

    // Use extracted table style if available, otherwise default to TableGrid
    const tableStyle = table.tableStyle || 'TableGrid';
    xml += `\n        <w:tblStyle w:val="${tableStyle}"/>`;
    xml += '\n        <w:tblW w:w="0" w:type="auto"/>';
    xml += '\n        <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';

    xml += '\n      </w:tblPr>';

    // Table grid (column definitions) - CRITICAL for proper rendering
    xml += '\n      <w:tblGrid>';

    // Use extracted column widths if available
    if (table.columnWidths && table.columnWidths.length > 0) {
      for (const width of table.columnWidths) {
        // Width is already in points from extraction, convert to twips
        xml += `\n        <w:gridCol w:w="${Math.round(width * 20)}"/>`;
      }
    } else if (table.rows.length > 0 && table.rows[0].cells.length > 0) {
      // Fallback: use cell widths from first row
      for (const cell of table.rows[0].cells) {
        const width = cell.width || 2160; // Default width in twips (1.5 inches)
        xml += `\n        <w:gridCol w:w="${Math.round(width * 20)}"/>`;
      }
    } else {
      // Last resort: use a default width for each column (if we can determine column count from first row)
      const columnCount = table.rows.length > 0 ? table.rows[0].cells.length : 1;
      for (let i = 0; i < columnCount; i++) {
        xml += '\n        <w:gridCol w:w="2160"/>';
      }
    }

    xml += '\n      </w:tblGrid>';

    // Rows
    for (const row of table.rows) {
      xml += '\n      <w:tr>';

      for (const cell of row.cells) {
        xml += this.buildTableCellXml(cell);
      }

      xml += '\n      </w:tr>';
    }

    xml += '\n    </w:tbl>';
    return xml;
  }

  private buildTableCellXml(cell: ExtractedTableCell): string {
    let xml = '\n        <w:tc>';

    // Cell properties - ALWAYS include tcPr (Word requires it)
    xml += '\n          <w:tcPr>';

    // Cell width (add default if not specified)
    if (cell.width) {
      xml += `\n            <w:tcW w:w="${Math.round(cell.width * 20)}" w:type="dxa"/>`;
    } else {
      // Default: auto width
      xml += '\n            <w:tcW w:w="0" w:type="auto"/>';
    }

    if (cell.colSpan && cell.colSpan > 1) {
      xml += `\n            <w:gridSpan w:val="${cell.colSpan}"/>`;
    }

    if (cell.rowSpan !== undefined) {
      if (cell.rowSpan === 1 || cell.rowSpan > 0) {
        xml += '\n            <w:vMerge w:val="restart"/>';
      } else {
        xml += '\n            <w:vMerge/>';
      }
    }

    if (cell.backgroundColor) {
      const fill = cell.backgroundColor.replace('#', '');
      xml += `\n            <w:shd w:val="clear" w:fill="${fill}"/>`;
    }

    if (cell.borders) {
      xml += '\n            <w:tcBorders>';
      for (const [side, border] of Object.entries(cell.borders)) {
        if (border) {
          const style = border.style || 'single';
          const size = border.size || 4;
          const color = border.color?.replace('#', '') || '000000';
          xml += `\n              <w:${side} w:val="${style}" w:sz="${size}" w:color="${color}"/>`;
        }
      }
      xml += '\n            </w:tcBorders>';
    }

    if (cell.verticalAlign) {
      xml += `\n            <w:vAlign w:val="${cell.verticalAlign}"/>`;
    }

    if (cell.margins) {
      xml += '\n            <w:tcMar>';
      for (const [side, margin] of Object.entries(cell.margins)) {
        if (margin !== undefined) {
          xml += `\n              <w:${side} w:w="${Math.round(margin * 20)}" w:type="dxa"/>`;
        }
      }
      xml += '\n            </w:tcMar>';
    }

    xml += '\n          </w:tcPr>';

    // Cell content (paragraphs)
    // CRITICAL: Every cell MUST have at least one paragraph, even if empty
    // Word will not render cells without at least one <w:p> element
    if (cell.content && cell.content.length > 0) {
      for (const para of cell.content) {
        xml += this.buildParagraphXml(para);
      }
    } else {
      // Add empty paragraph if cell has no content
      xml += '\n          <w:p>\n          </w:p>';
    }

    xml += '\n        </w:tc>';
    return xml;
  }

  private buildRunFormattingXml(formatting: RunFormatting): string {
    let xml = '';

    if (formatting.bold) {
      xml += '\n          <w:b/>';
      xml += '\n          <w:bCs/>';
    }

    if (formatting.italic) {
      xml += '\n          <w:i/>';
      xml += '\n          <w:iCs/>';
    }

    if (formatting.underline) {
      xml += '\n          <w:u w:val="single"/>';
    }

    if (formatting.strike) {
      xml += '\n          <w:strike/>';
    }

    if (formatting.fontSize) {
      const sz = formatting.fontSize * 2;
      xml += `\n          <w:sz w:val="${sz}"/>`;
      xml += `\n          <w:szCs w:val="${sz}"/>`;
    }

    // Font family with theme references
    if (formatting.fontFamily || formatting.fontThemeAscii || formatting.fontThemeEastAsia || formatting.fontThemeHAnsi || formatting.fontThemeCs) {
      xml += '\n          <w:rFonts';
      if (formatting.fontFamily) xml += ` w:ascii="${formatting.fontFamily}" w:hAnsi="${formatting.fontFamily}"`;
      if (formatting.fontThemeAscii) xml += ` w:asciiTheme="${formatting.fontThemeAscii}"`;
      if (formatting.fontThemeEastAsia) xml += ` w:eastAsiaTheme="${formatting.fontThemeEastAsia}"`;
      if (formatting.fontThemeHAnsi) xml += ` w:hAnsiTheme="${formatting.fontThemeHAnsi}"`;
      if (formatting.fontThemeCs) xml += ` w:cstheme="${formatting.fontThemeCs}"`;
      xml += '/>';
    }

    // Color with theme references
    if (formatting.color || formatting.colorTheme) {
      xml += '\n          <w:color';
      if (formatting.color) {
        const color = formatting.color.replace('#', '');
        xml += ` w:val="${color}"`;
      }
      if (formatting.colorTheme) xml += ` w:themeColor="${formatting.colorTheme}"`;
      if (formatting.colorThemeShade) xml += ` w:themeShade="${formatting.colorThemeShade}"`;
      xml += '/>';
    }

    if (formatting.highlight) {
      xml += `\n          <w:highlight w:val="${formatting.highlight}"/>`;
    }

    if (formatting.verticalAlign === 'superscript' || formatting.verticalAlign === 'subscript') {
      xml += `\n          <w:vertAlign w:val="${formatting.verticalAlign}"/>`;
    }

    return xml;
  }

  private buildSpacingXml(spacing: SpacingInfo): string {
    let xml = '\n        <w:spacing';

    if (spacing.before !== undefined) {
      xml += ` w:before="${Math.round(spacing.before * 20)}"`;
    }

    if (spacing.after !== undefined) {
      xml += ` w:after="${Math.round(spacing.after * 20)}"`;
    }

    if (spacing.line !== undefined) {
      xml += ` w:line="${Math.round(spacing.line * 20)}"`;
    }

    if (spacing.lineRule) {
      xml += ` w:lineRule="${spacing.lineRule}"`;
    }

    xml += '/>';
    return xml;
  }

  private buildIndentationXml(indentation: IndentationInfo): string {
    let xml = '\n        <w:ind';

    if (indentation.left !== undefined) {
      xml += ` w:left="${Math.round(indentation.left * 20)}"`;
    }

    if (indentation.right !== undefined) {
      xml += ` w:right="${Math.round(indentation.right * 20)}"`;
    }

    if (indentation.firstLine !== undefined) {
      xml += ` w:firstLine="${Math.round(indentation.firstLine * 20)}"`;
    }

    if (indentation.hanging !== undefined) {
      xml += ` w:hanging="${Math.round(indentation.hanging * 20)}"`;
    }

    xml += '/>';
    return xml;
  }
}

