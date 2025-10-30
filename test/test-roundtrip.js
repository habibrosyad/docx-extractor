import { DocxExtractor, DocxBuilder } from '../dist/index.js';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

(async () => {
  console.log('=== Round-trip Test ===\n');

  // Read from ~/Downloads/test4.docx
  const inputPath = path.join(__dirname, 'test.docx');

  if (!fs.existsSync(inputPath)) {
    console.error(`❌ File not found: ${inputPath}`);
    console.error('Please ensure test4.docx exists in your Downloads folder.');
    process.exit(1);
  }

  console.log(`1️⃣ Extracting ${inputPath}...`);
  const originalBuffer = fs.readFileSync(inputPath);

  const extractor1 = new DocxExtractor();
  const originalDoc = await extractor1.extract(originalBuffer);

  console.log(`   - ${originalDoc.paragraphs.length} paragraphs`);
  console.log(`   - ${originalDoc.tables.length} tables`);

  // Save extracted JSON
  const jsonPath = path.join(process.cwd(), 'test', 'output.json');
  fs.writeFileSync(jsonPath, JSON.stringify(originalDoc, (key, value) => {
    // Handle Map serialization
    if (value instanceof Map) {
      return Object.fromEntries(value);
    }
    // Handle Uint8Array serialization
    if (value instanceof Uint8Array) {
      return Array.from(value);
    }
    return value;
  }, 2));
  console.log(`   - Saved JSON to ${jsonPath}`);

  console.log('\n2️⃣ Building output.docx...');
  const builder = new DocxBuilder();
  const rebuiltBuffer = await builder.build(originalDoc);

  const outputPath = path.join(process.cwd(), 'test', 'output.docx');
  fs.writeFileSync(outputPath, rebuiltBuffer);
  console.log(`   - Saved DOCX to ${outputPath}`);

  console.log('\n3️⃣ Extracting output.docx...');
  const extractor2 = new DocxExtractor();
  const rebuiltDoc = await extractor2.extract(rebuiltBuffer);

  console.log(`   - ${rebuiltDoc.paragraphs.length} paragraphs`);
  console.log(`   - ${rebuiltDoc.tables.length} tables`);

  console.log('\n4️⃣ Comparing documents...');

  // Compare paragraph count
  const paraMatch = originalDoc.paragraphs.length === rebuiltDoc.paragraphs.length;
  console.log(`   Paragraph count: ${paraMatch ? '✅' : '❌'} (${originalDoc.paragraphs.length} vs ${rebuiltDoc.paragraphs.length})`);

  // Compare table count
  const tableMatch = originalDoc.tables.length === rebuiltDoc.tables.length;
  console.log(`   Table count: ${tableMatch ? '✅' : '❌'} (${originalDoc.tables.length} vs ${rebuiltDoc.tables.length})`);

  // Compare first paragraph text (if any)
  if (originalDoc.paragraphs.length > 0 && rebuiltDoc.paragraphs.length > 0) {
    const origText = originalDoc.paragraphs[0].runs.map(r => r.text || '').join('');
    const rebuiltText = rebuiltDoc.paragraphs[0].runs.map(r => r.text || '').join('');
    const textMatch = origText === rebuiltText;
    console.log(`   First paragraph text: ${textMatch ? '✅' : '❌'}`);
  }

  // Compare first table (if any)
  if (originalDoc.tables.length > 0 && rebuiltDoc.tables.length > 0) {
    const origTable = originalDoc.tables[0];
    const rebuiltTable = rebuiltDoc.tables[0];

    if (origTable.rows.length > 0 && rebuiltTable.rows.length > 0) {
      const origCell = origTable.rows[0].cells[0];
      const rebuiltCell = rebuiltTable.rows[0].cells[0];

      const origCellText = origCell.content.map(p => p.runs.map(r => r.text || '').join('')).join('');
      const rebuiltCellText = rebuiltCell.content.map(p => p.runs.map(r => r.text || '').join('')).join('');

      const cellMatch = origCellText === rebuiltCellText;
      console.log(`   First table cell [1,1]: ${cellMatch ? '✅' : '❌'}`);
    }
  }

  console.log('\n✅ Round-trip test complete');
})();
