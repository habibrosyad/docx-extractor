# Test Files

## test.docx

A sample DOCX file (`test.docx`) is included with:
- Various headings (Heading 1, Heading 2, Heading 3)
- Formatted text (bold, italic, underline)
- Bullet lists and numbered lists
- A table with multiple rows and columns
- Different paragraph styles

## Running the Test

```bash
npm test
```

This will:
1. Extract `test.docx` to structured data
2. Save extracted data to `extracted_data.json`
3. Rebuild DOCX from extracted data to `output.docx`
4. Verify the round-trip conversion

## Output Files

- `extracted_data.json` - Full structured representation of the document
- `output.docx` - Rebuilt DOCX file from extracted data

