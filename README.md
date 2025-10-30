# @habibrosyad/docx-extractor

Extract DOCX documents to structured data and rebuild them programmatically. This library provides full round-trip conversion with preservation of formatting, styles, tables, images, and numbering.

## Installation

```bash
npm install @habibrosyad/docx-extractor
```

## Quick Start

```javascript
import { DocxExtractor, DocxBuilder } from '@habibrosyad/docx-extractor';
import fs from 'fs';

// Extract DOCX to structured data
const extractor = new DocxExtractor();
const buffer = fs.readFileSync('document.docx');
const document = await extractor.extract(buffer);

// Modify the document structure
document.paragraphs.forEach(para => {
  if (para.runs) {
    para.runs.forEach(run => {
      // Modify text, formatting, etc.
    });
  }
});

// Rebuild DOCX
const builder = new DocxBuilder();
const newBuffer = await builder.build(document);
fs.writeFileSync('output.docx', newBuffer);
```

## Features

- ✅ **Full Structure Extraction**: Paragraphs, tables, images, styles, numbering
- ✅ **Formatting Preservation**: Fonts, colors, spacing, alignment, borders
- ✅ **Round-Trip Conversion**: Extract → Modify → Rebuild without data loss
- ✅ **Style Support**: All paragraph and character styles with inheritance
- ✅ **Table Support**: Complete table structure with cell properties
- ✅ **Image Support**: Extract and embed images in documents
- ✅ **Numbering Support**: Bullet lists and numbered lists

## API

### DocxExtractor

Extract a DOCX file to structured data.

```javascript
const extractor = new DocxExtractor();
const document = await extractor.extract(buffer);
```

**Returns**: `ExtractedDocument` with:
- `paragraphs`: Array of paragraphs
- `tables`: Array of tables  
- `body`: Ordered sequence of paragraphs and tables
- `styles`: Map of style definitions
- `defaults`: Document-wide run formatting defaults
- `paragraphDefaults`: Document-wide paragraph defaults (spacing, etc.)
- `numbering`: Map of numbering definitions
- `mediaFiles`: Map of embedded images/media

### DocxBuilder

Build a DOCX file from structured data.

```javascript
const builder = new DocxBuilder();
const buffer = await builder.build(document);
```

**Parameters**: `ExtractedDocument` (from extractor or construct it yourself)

**Returns**: `Promise<Uint8Array>` - DOCX file buffer (compatible with Node.js, browsers, and edge runtimes)

## Platform Compatibility

This library works across all JavaScript runtimes:
- ✅ **Node.js**: Full support (v18+)
- ✅ **Cloudflare Workers**: Full support
- ✅ **Browsers**: Full support
- ✅ **Deno**: Full support
- ✅ **Bun**: Full support

### Using in Cloudflare Workers

```javascript
export default {
  async fetch(request) {
    // Get DOCX file from request
    const arrayBuffer = await request.arrayBuffer();
    
    // Extract and process
    const extractor = new DocxExtractor();
    const document = await extractor.extract(arrayBuffer);
    
    // Modify document...
    
    // Rebuild
    const builder = new DocxBuilder();
    const outputBuffer = await builder.build(document);
    
    // Return as response
    return new Response(outputBuffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': 'attachment; filename="output.docx"'
      }
    });
  }
};
```

## Document Structure

### Paragraph
- `runs`: Array of text runs with formatting
- `spacing`: Paragraph spacing (before, after, line)
- `alignment`: Left, center, right, justify
- `indentation`: Left, right, firstLine, hanging
- `styleName`: Applied paragraph style

### Table
- `rows`: Array of table rows
- `rows[i].cells`: Array of table cells
- `cells[i].runs`: Text content with formatting
- `cells[i].width`, `backgroundColor`, `borders`: Cell properties

### Run Formatting
- `bold`, `italic`, `underline`, `strike`
- `fontSize`, `fontFamily`, `color`
- `highlight` color

## Requirements

- Modern JavaScript runtime with ES modules support (Node.js >= 18, Deno, Bun, Cloudflare Workers, or modern browsers)

## License

MIT
