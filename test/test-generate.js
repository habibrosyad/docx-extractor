/**
 * Generate comprehensive test.docx with manually composed payload
 * Demonstrates all features: headings, formatting, tables, lists, styles
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { DocxBuilder } from '../dist/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function generateTestDocx() {
  try {
    console.log('📝 Generating test.docx from payload...\n');

    // Manually compose document with all features
    const document = {
      // Body elements (in document order)
      body: [
        // Title
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'Document Test', formatting: { bold: true, fontSize: 18 } }],
            styleName: 'Title',
            alignment: 'center',
            spacing: { after: 0 },
            isEmpty: false
          }
        },

        // Empty line
        {
          type: 'paragraph',
          data: { runs: [], isEmpty: true, spacing: { after: 20 } }
        },

        // Heading 1
        {
          type: 'paragraph',
          data: {
            runs: [{ text: '1. Text Formatting' }],
            styleName: 'Heading1',
            spacing: { before: 0, after: 0 },
            isEmpty: false
          }
        },

        // Paragraph with mixed formatting
        {
          type: 'paragraph',
          data: {
            runs: [
              { text: 'This document demonstrates ' },
              { text: 'bold', formatting: { bold: true } },
              { text: ', ' },
              { text: 'italic', formatting: { italic: true } },
              { text: ', ' },
              { text: 'underline', formatting: { underline: true } },
              { text: ', ' },
              { text: 'strikethrough', formatting: { strike: true } },
              { text: ', ' },
              { text: 'red text', formatting: { color: 'FF0000' } },
              { text: ', and ' },
              { text: 'highlighted', formatting: { highlight: 'yellow' } },
              { text: ' text.' }
            ],
            spacing: { after: 0 },
            isEmpty: false
          }
        },

        // Superscript and subscript
        {
          type: 'paragraph',
          data: {
            runs: [
              { text: 'Scientific notation: E=mc' },
              { text: '2', formatting: { verticalAlign: 'superscript' } },
              { text: ' and chemical formula: H' },
              { text: '2', formatting: { verticalAlign: 'subscript' } },
              { text: 'O' }
            ],
            spacing: { after: 0 },
            isEmpty: false
          }
        },

        // Empty line
        {
          type: 'paragraph',
          data: { runs: [], isEmpty: true, spacing: { after: 20 } }
        },

        // Heading 2
        {
          type: 'paragraph',
          data: {
            runs: [{ text: '2. Font Styles' }],
            styleName: 'Heading2',
            spacing: { before: 0, after: 0 },
            isEmpty: false
          }
        },

        // Different fonts
        {
          type: 'paragraph',
          data: {
            runs: [
              { text: 'Default font (Calibri), ' },
              { text: 'Arial', formatting: { fontFamily: 'Arial' } },
              { text: ', ' },
              { text: 'Times New Roman', formatting: { fontFamily: 'Times New Roman' } },
              { text: ', and ' },
              { text: 'Courier New', formatting: { fontFamily: 'Courier New' } }
            ],
            spacing: { after: 0 },
            isEmpty: false
          }
        },

        // Different sizes
        {
          type: 'paragraph',
          data: {
            runs: [
              { text: 'Font sizes: ' },
              { text: 'small (8pt)', formatting: { fontSize: 8 } },
              { text: ', ' },
              { text: 'normal (11pt)', formatting: { fontSize: 11 } },
              { text: ', ' },
              { text: 'large (14pt)', formatting: { fontSize: 14 } },
              { text: ', ' },
              { text: 'huge (18pt)', formatting: { fontSize: 18 } }
            ],
            spacing: { after: 0 },
            isEmpty: false
          }
        },

        // Empty line
        {
          type: 'paragraph',
          data: { runs: [], isEmpty: true, spacing: { after: 20 } }
        },

        // Heading 2
        {
          type: 'paragraph',
          data: {
            runs: [{ text: '3. Paragraph Alignment' }],
            styleName: 'Heading2',
            spacing: { before: 0, after: 0 },
            isEmpty: false
          }
        },

        // Left aligned
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'This paragraph is left aligned (default).' }],
            alignment: 'left',
            spacing: { after: 0 },
            isEmpty: false
          }
        },

        // Center aligned
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'This paragraph is center aligned.' }],
            alignment: 'center',
            spacing: { after: 0 },
            isEmpty: false
          }
        },

        // Right aligned
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'This paragraph is right aligned.' }],
            alignment: 'right',
            spacing: { after: 20 },
            isEmpty: false
          }
        },

        // Justified
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'This paragraph is justified. It should stretch across the entire width when there is enough text to demonstrate the justification effect properly.' }],
            alignment: 'justify',
            spacing: { after: 20 },
            isEmpty: false
          }
        },

        // Empty line
        {
          type: 'paragraph',
          data: { runs: [], isEmpty: true, spacing: { after: 20 } }
        },

        // Heading 2 - Tables
        {
          type: 'paragraph',
          data: {
            runs: [{ text: '4. Tables' }],
            styleName: 'Heading2',
            spacing: { before: 0, after: 0 },
            isEmpty: false
          }
        },

        // Simple table with proper cell widths (CRITICAL for Word compatibility)
        {
          type: 'table',
          data: {
            rows: [
              {
                cells: [
                  {
                    content: [{ runs: [{ text: 'Header 1', formatting: { bold: true } }], isEmpty: false }],
                    width: 144  // 2 inches in points
                  },
                  {
                    content: [{ runs: [{ text: 'Header 2', formatting: { bold: true } }], isEmpty: false }],
                    width: 144
                  },
                  {
                    content: [{ runs: [{ text: 'Header 3', formatting: { bold: true } }], isEmpty: false }],
                    width: 144
                  }
                ]
              },
              {
                cells: [
                  {
                    content: [{ runs: [{ text: 'Row 1, Cell 1' }], isEmpty: false }],
                    width: 144
                  },
                  {
                    content: [{ runs: [{ text: 'Row 1, Cell 2' }], isEmpty: false }],
                    width: 144
                  },
                  {
                    content: [{ runs: [{ text: 'Row 1, Cell 3' }], isEmpty: false }],
                    width: 144
                  }
                ]
              },
              {
                cells: [
                  {
                    content: [{ runs: [{ text: 'Row 2, Cell 1' }], isEmpty: false }],
                    width: 144
                  },
                  {
                    content: [{ runs: [{ text: 'Row 2, Cell 2' }], isEmpty: false }],
                    width: 144
                  },
                  {
                    content: [{ runs: [{ text: 'Row 2, Cell 3' }], isEmpty: false }],
                    width: 144
                  }
                ]
              }
            ]
          }
        },

        // Empty line
        {
          type: 'paragraph',
          data: { runs: [], isEmpty: true, spacing: { after: 20 } }
        },

        // Heading 3
        {
          type: 'paragraph',
          data: {
            runs: [{ text: '4.1 Formatted Table with Colors' }],
            styleName: 'Heading3',
            spacing: { before: 0, after: 0 },
            isEmpty: false
          }
        },

        // Formatted table with colors and proper widths
        {
          type: 'table',
          data: {
            rows: [
              {
                cells: [
                  {
                    content: [{ runs: [{ text: 'Product', formatting: { bold: true, color: 'FFFFFF' } }], isEmpty: false }],
                    backgroundColor: '4472C4',
                    width: 180
                  },
                  {
                    content: [{ runs: [{ text: 'Price', formatting: { bold: true, color: 'FFFFFF' } }], isEmpty: false }],
                    backgroundColor: '4472C4',
                    width: 120
                  },
                  {
                    content: [{ runs: [{ text: 'Stock', formatting: { bold: true, color: 'FFFFFF' } }], isEmpty: false }],
                    backgroundColor: '4472C4',
                    width: 120
                  }
                ]
              },
              {
                cells: [
                  {
                    content: [{ runs: [{ text: 'Widget A' }], isEmpty: false }],
                    width: 180
                  },
                  {
                    content: [{ runs: [{ text: '$19.99', formatting: { color: '00B050', bold: true } }], isEmpty: false }],
                    width: 120
                  },
                  {
                    content: [{ runs: [{ text: '50' }], isEmpty: false }],
                    width: 120
                  }
                ]
              },
              {
                cells: [
                  {
                    content: [{ runs: [{ text: 'Widget B' }], isEmpty: false }],
                    width: 180
                  },
                  {
                    content: [{ runs: [{ text: '$29.99', formatting: { color: '00B050', bold: true } }], isEmpty: false }],
                    width: 120
                  },
                  {
                    content: [{ runs: [{ text: '30' }], isEmpty: false }],
                    width: 120
                  }
                ]
              }
            ]
          }
        },

        // Empty line
        {
          type: 'paragraph',
          data: { runs: [], isEmpty: true, spacing: { after: 20 } }
        },

        // Heading 2
        {
          type: 'paragraph',
          data: {
            runs: [{ text: '5. Lists' }],
            styleName: 'Heading2',
            spacing: { before: 0, after: 0 },
            isEmpty: false
          }
        },

        // Bullet list
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'First bullet item' }],
            numbering: { id: '1', level: 0 },
            spacing: { after: 0 },
            isEmpty: false
          }
        },
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'Second bullet item with ' }, { text: 'bold text', formatting: { bold: true } }],
            numbering: { id: '1', level: 0 },
            spacing: { after: 0 },
            isEmpty: false
          }
        },
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'Third bullet item' }],
            numbering: { id: '1', level: 0 },
            spacing: { after: 20 },
            isEmpty: false
          }
        },

        // Empty line
        {
          type: 'paragraph',
          data: { runs: [], isEmpty: true, spacing: { after: 20 } }
        },

        // Heading 3
        {
          type: 'paragraph',
          data: {
            runs: [{ text: '5.1 Numbered List' }],
            styleName: 'Heading3',
            spacing: { before: 0, after: 0 },
            isEmpty: false
          }
        },

        // Numbered list
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'First step in the process' }],
            numbering: { id: '2', level: 0 },
            spacing: { after: 0 },
            isEmpty: false
          }
        },
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'Second step in the process' }],
            numbering: { id: '2', level: 0 },
            spacing: { after: 0 },
            isEmpty: false
          }
        },
        {
          type: 'paragraph',
          data: {
            runs: [{ text: 'Third step in the process' }],
            numbering: { id: '2', level: 0 },
            spacing: { after: 20 },
            isEmpty: false
          }
        },

        // Empty line
        {
          type: 'paragraph',
          data: { runs: [], isEmpty: true, spacing: { after: 20 } }
        },

        // Conclusion
        {
          type: 'paragraph',
          data: {
            runs: [{ text: '6. Conclusion' }],
            styleName: 'Heading2',
            spacing: { before: 0, after: 0 },
            isEmpty: false
          }
        },

        {
          type: 'paragraph',
          data: {
            runs: [
              { text: 'This document successfully demonstrates all major features: ' },
              { text: 'headings', formatting: { bold: true } },
              { text: ', ' },
              { text: 'text formatting', formatting: { italic: true } },
              { text: ', ' },
              { text: 'tables', formatting: { underline: true } },
              { text: ', and ' },
              { text: 'lists', formatting: { bold: true, italic: true } },
              { text: '. All elements are compatible with Microsoft Word.' }
            ],
            spacing: { after: 20 },
            isEmpty: false
          }
        }
      ],

      // Styles (comprehensive style definitions)
      styles: new Map([
        ['Normal', {
          name: 'Normal',
          type: 'paragraph',
          isDefault: true,
          runFormatting: { fontSize: 11, fontFamily: 'Calibri' }
        }],
        ['Title', {
          name: 'Title',
          type: 'paragraph',
          runFormatting: { fontSize: 18, bold: true, color: '000000' },
          spacing: { after: 0 },
          alignment: 'center'
        }],
        ['Heading1', {
          name: 'Heading 1',
          type: 'paragraph',
          runFormatting: { fontSize: 16, bold: true, color: '2E74B5' },
          spacing: { before: 0, after: 0 }
        }],
        ['Heading2', {
          name: 'Heading 2',
          type: 'paragraph',
          runFormatting: { fontSize: 13, bold: true, color: '2E74B5' },
          spacing: { before: 0, after: 0 }
        }],
        ['Heading3', {
          name: 'Heading 3',
          type: 'paragraph',
          runFormatting: { fontSize: 11, bold: true, color: '1F4D78' },
          spacing: { before: 0, after: 0 }
        }]
      ]),

      // Document defaults
      defaults: {
        fontSize: 11,
        fontFamily: 'Calibri'
      },

      // Paragraph defaults
      paragraphDefaults: {
        spacing: {
          after: 0,
          line: 13.8,
          lineRule: 'auto'
        }
      },

      // Numbering definitions
      numbering: new Map([
        ['1', {
          abstractNumId: '0',
          multiLevelType: 'hybridMultilevel',
          levels: new Map([
            [0, {
              format: 'bullet',
              text: '•',
              alignment: 'left',
              indentation: { left: 360, hanging: 360 },
              fontFamily: 'Symbol'
            }]
          ])
        }],
        ['2', {
          abstractNumId: '1',
          multiLevelType: 'hybridMultilevel',
          levels: new Map([
            [0, {
              format: 'decimal',
              text: '%1.',
              alignment: 'left',
              indentation: { left: 360, hanging: 360 }
            }]
          ])
        }]
      ]),

      // No media files
      mediaFiles: new Map()
    };

    // Build paragraphs and tables arrays from body (for compatibility)
    document.paragraphs = document.body.filter(e => e.type === 'paragraph').map(e => e.data);
    document.tables = document.body.filter(e => e.type === 'table').map(e => e.data);

    console.log(`🔨 Building test.docx with manual payload...`);

    // Build the DOCX
    const builder = new DocxBuilder();
    const outputBuffer = await builder.build(document);

    // Save to test directory
    const outputPath = path.join(__dirname, 'test.docx');
    fs.writeFileSync(outputPath, outputBuffer);

    console.log(`\n✅ Created: ${outputPath}`);
    console.log(`   Size: ${outputBuffer.length} bytes (${(outputBuffer.length / 1024).toFixed(2)} KB)`);
    console.log(`\n📋 Document contains:`);
    console.log(`   - ${document.paragraphs.length} paragraphs`);
    console.log(`   - ${document.tables.length} tables`);
    console.log(`   - ${document.body.filter(e => e.type === 'paragraph' && e.data.numbering).length} list items`);
    console.log(`   - ${document.styles.size} custom styles`);
    console.log(`   - ${document.numbering.size} numbering definitions`);
    console.log(`\n✨ Features demonstrated:`);
    console.log(`   ✅ Multiple heading levels (Title, H1, H2, H3)`);
    console.log(`   ✅ Text formatting (bold, italic, underline, strike)`);
    console.log(`   ✅ Superscript and subscript`);
    console.log(`   ✅ Colors and highlighting`);
    console.log(`   ✅ Multiple font families and sizes`);
    console.log(`   ✅ Paragraph alignment (left, center, right, justify)`);
    console.log(`   ✅ Bullet and numbered lists`);
    console.log(`   ✅ Tables with formatting and colors`);
    console.log(`\n✅ Ready to test! Run: npm test`);

  } catch (error) {
    console.error('❌ Error:', error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

generateTestDocx();
