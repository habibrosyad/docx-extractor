/**
 * Test round-trip DOCX conversion: Extract → Build → Extract
 * Uses test.docx file from the test directory
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { DocxExtractor, DocxBuilder } from '../dist/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function testRoundTrip() {
  try {
    console.log('🔄 Testing DOCX Round-trip Conversion...\n');

    // Input file
    const inputPath = path.join(__dirname, 'test.docx');

    if (!fs.existsSync(inputPath)) {
      console.error(`❌ Input file not found: ${inputPath}`);
      console.error('   Please create a test.docx file with various headings, formatting, and tables.');
      process.exit(1);
    }

    console.log('📖 Step 1: Extracting from test.docx...');
    const inputBuffer = fs.readFileSync(inputPath);
    const extractor = new DocxExtractor();
    const document = await extractor.extract(inputBuffer);

    console.log(`✅ Extracted: ${document.paragraphs.length} paragraphs, ${document.tables.length} tables`);
    console.log(`📝 Styles: ${document.styles.size}, Numbering: ${document.numbering?.size || 0}\n`);

    // Step 2: Save extracted data as JSON
    console.log('💾 Step 2: Saving extracted data to extracted_data.json...');
    const jsonData = {
      paragraphs: document.paragraphs,
      tables: document.tables,
      body: document.body,
      styles: Object.fromEntries(document.styles),
      defaults: document.defaults,
      paragraphDefaults: document.paragraphDefaults,
      numbering: document.numbering ? Object.fromEntries(document.numbering) : {},
      mediaFiles: document.mediaFiles ? Object.fromEntries(
        Array.from(document.mediaFiles.entries()).map(([k, v]) => [k, '[Binary data: ' + v.length + ' bytes]'])
      ) : {}
    };

    const jsonPath = path.join(__dirname, 'extracted_data.json');
    fs.writeFileSync(jsonPath, JSON.stringify(jsonData, null, 2));
    const jsonStats = fs.statSync(jsonPath);
    console.log(`✅ Saved: ${jsonStats.size} bytes (${(jsonStats.size / 1024).toFixed(2)} KB)`);
    console.log(`   Location: ${jsonPath}\n`);

    // Step 3: Build new DOCX from extracted data
    console.log('🔨 Step 3: Building new DOCX from extracted data...');
    const builder = new DocxBuilder();
    const outputBuffer = await builder.build(document);

    // Save the output file
    const outputPath = path.join(__dirname, 'output.docx');
    fs.writeFileSync(outputPath, outputBuffer);
    console.log(`✅ Built: ${outputBuffer.length} bytes`);
    console.log(`   Location: ${outputPath}\n`);

    // Step 4: Verify by re-extracting
    console.log('🔍 Step 4: Verifying output by re-extracting...');
    const document2 = await extractor.extract(outputBuffer);
    console.log(`✅ Re-extracted: ${document2.paragraphs.length} paragraphs, ${document2.tables.length} tables\n`);

    // Step 5: Compare
    console.log('📊 Step 5: Comparison Results:');
    console.log(`  📝 Paragraphs: ${document.paragraphs.length} → ${document2.paragraphs.length} ${document.paragraphs.length === document2.paragraphs.length ? '✅' : '❌'}`);
    console.log(`  📋 Tables: ${document.tables.length} → ${document2.tables.length} ${document.tables.length === document2.tables.length ? '✅' : '❌'}`);
    console.log(`  🎨 Styles: ${document.styles.size} → ${document2.styles.size} ${document.styles.size === document2.styles.size ? '✅' : '❌'}\n`);

    // Check first few paragraphs for content match
    console.log('📝 Checking first 5 paragraphs content:');
    let contentMatch = 0;
    for (let i = 0; i < Math.min(5, document.paragraphs.length); i++) {
      const text1 = document.paragraphs[i].runs.map(r => r.text || '').join('');
      const text2 = document2.paragraphs[i].runs.map(r => r.text || '').join('');
      const match = text1 === text2;
      if (match) contentMatch++;

      const preview = text1.substring(0, 50) + (text1.length > 50 ? '...' : '');
      console.log(`  [${i}] ${match ? '✅' : '❌'} "${preview}"`);
    }

    console.log(`\n🎯 Content Match: ${contentMatch}/5 paragraphs`);

    const allMatch = document.paragraphs.length === document2.paragraphs.length &&
      document.tables.length === document2.tables.length &&
      contentMatch === 5;

    if (allMatch) {
      console.log('\n🎉 SUCCESS! Round-trip conversion worked perfectly!');
      console.log('   The rebuilt DOCX matches the original structure and content.\n');
    } else {
      console.log('\n⚠️  PARTIAL SUCCESS! The DOCX was rebuilt but some differences exist.');
    }

    console.log('📁 Output files:');
    console.log(`   - ${jsonPath}`);
    console.log(`   - ${outputPath}\n`);

  } catch (error) {
    console.error('❌ Error:', error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

testRoundTrip();
