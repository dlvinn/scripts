# ODF (OpenDocument Format) Support

## Overview

Your `docx_format_fixer.py` now supports **both** Microsoft Word (.docx) and OpenDocument (.odt) formats!

## Supported Formats

| Format | Extension | Library Used | Status |
|--------|-----------|--------------|--------|
| Microsoft Word | `.docx` | python-docx | ✅ Fully Supported |
| OpenDocument Text | `.odt` | odfpy | ✅ Fully Supported |

## What Gets Fixed

### For Both DOCX and ODT:
1. ✅ **Encoding Issues** (Mojibake: `ÂÈ` → `آب`)
2. ✅ **RTL Direction** (Right-to-Left text flow)
3. ✅ **Right Alignment** (Visual alignment)
4. ✅ **Table Content** (Paragraphs inside tables)

### DOCX-Specific:
- Column reversal in tables (mirroring for RTL)
- Numbered header fixes (e.g., "المجال .2")
- XML-level final alignment pass

### ODT-Specific:
- Automatic style creation with RTL properties
- ODF-compliant writing-mode attributes
- fo:text-align="end" for RTL right alignment

## Usage

### Process Both DOCX and ODT Files:
```bash
python docx_format_fixer.py "C:\path\to\documents"
```

The script will automatically:
- Find all `.docx` files
- Find all `.odt` files
- Process each with the appropriate handler
- Show statistics for both formats

### Example Output:
```
Arabic Document Format Fixer (DOCX & ODF)
✓ Mojibake encoding fix enabled (ÂÈ → آب)

Searching recursively in: C:\Users\Documents\Arabic Docs
Found 8 document(s) to process:
  - 5 .docx file(s)
  - 3 .odt file(s)
============================================================

Processing: report1.docx
  Original: 45 paragraphs, 2345 chars, 2 tables
✓ Fixed and saved to: report1_fixed.docx
  - RTL paragraphs: 45
  - Alignments: 45
  - Encoding fixes: 12
  - Table cells: 8
  ✓ Content validation: PASSED

Processing ODF: report2.odt
  Original: 32 paragraphs, 1876 chars, 1 tables
✓ Fixed and saved to: report2_fixed.odt
  - RTL paragraphs: 32
  - Alignments: 32
  - Encoding fixes: 8
  - Table cells: 4
  ✓ Content validation: PASSED

...

============================================================
SUMMARY REPORT
============================================================
Documents processed: 8
Successfully fixed: 8
Errors: 0
```

## Installation

The ODF support requires the `odfpy` library:

```bash
pip install odfpy
```

Or install all dependencies:

```bash
pip install -r requirements.txt
```

## Technical Details

### ODF Handler (`odf_handler.py`)

The ODF handler uses the `odfpy` library to:

1. **Load ODF documents**: `odf.opendocument.load()`
2. **Parse XML structure**: Find all `text.P` (paragraph) elements
3. **Fix encoding**: Apply the same `str.maketrans` Mojibake fix
4. **Create RTL styles**: Generate automatic styles with:
   - `writing-mode='rl-tb'` (Right-to-Left, Top-to-Bottom)
   - `text-align='end'` (Right align for RTL)
5. **Apply to all content**: Paragraphs, tables, lists
6. **Save modified ODF**: Preserve all original formatting

### Architecture

```
docx_format_fixer.py
├── Detects file type (.docx or .odt)
├── Routes to appropriate handler
│
├── DOCX Handler (ArabicDocxFixer)
│   └── Uses python-docx library
│
└── ODF Handler (ODFHandler)
    └── Uses odfpy library
```

## File Structure

```
document fixer/
├── docx_format_fixer.py      # Main script (supports both formats)
├── odf_handler.py             # ODF-specific handler
├── encoding_fixer.py          # Shared encoding fix (Mojibake)
├── requirements.txt           # Updated with odfpy
└── ODF_SUPPORT.md            # This file
```

## Limitations

### ODF-Specific Limitations:
1. **Column reversal**: Not yet implemented for ODF tables (DOCX only)
2. **Numbered headers**: Pattern fixes are DOCX-specific
3. **Style preservation**: Creates new styles rather than modifying existing ones

These limitations don't affect the core functionality:
- ✅ Encoding fixes work perfectly
- ✅ RTL direction is correct
- ✅ Right alignment is applied
- ✅ All text content is preserved

## Command-Line Options

All existing options work for both formats:

```bash
# Dry run (preview without changes)
python docx_format_fixer.py "C:\path" --dry-run

# Non-recursive (current folder only)
python docx_format_fixer.py "C:\path" --no-recursive

# Disable encoding fix
python docx_format_fixer.py "C:\path" --no-encoding-fix
```

## Compatibility

### LibreOffice / OpenOffice
- ✅ Full compatibility
- ODT files can be opened and edited after processing

### Microsoft Word
- ✅ Can open `.odt` files (with limited formatting)
- Recommend using LibreOffice for best results

### Google Docs
- ✅ Can import `.odt` files
- RTL and alignment preserved

## Testing

Test the ODF handler directly:

```bash
python odf_handler.py
```

Expected output:
```
ODF Handler Test
============================================================
Encoding test:
  Original: ÂÈ ÇáÚÑÇÞ
  Fixed:    آب العراق
```

## Troubleshooting

### "No module named 'odf'"
```bash
pip install odfpy
```

### ODT file won't open after processing
- Check that the original file opens correctly
- Ensure it's a valid ODF document (not a renamed DOCX)
- Try with `--dry-run` first to preview

### Encoding fixes not working in ODT
- Verify encoding fix is enabled (default)
- Check that the text actually contains Mojibake
- Run `discover_mojibake.py` to diagnose

## Future Enhancements

Potential future additions for ODF:
- [ ] Table column reversal (like DOCX)
- [ ] Numbered header pattern fixes
- [ ] Better style preservation
- [ ] Support for .ods (spreadsheets)
- [ ] Support for .odp (presentations)

## Need Help?

Run the test suite:
```bash
python odf_handler.py
python encoding_fixer.py
```

Check for ODF files:
```bash
python docx_format_fixer.py "C:\path" --dry-run
```

All tools work with both formats! 🎉
