# ODF to DOCX Converter

## Overview

Convert OpenDocument (.odt) files to Microsoft Word (.docx) format while preserving:
- ✅ Text content
- ✅ Paragraph structure
- ✅ Tables
- ✅ RTL/Arabic text direction
- ✅ Right alignment for Arabic text

## Quick Start

### Convert Single File
```bash
python odf_to_docx_converter.py --file document.odt
```

Output: `document.docx` (same folder)

### Convert Entire Folder
```bash
python odf_to_docx_converter.py --folder "C:\Documents"
```

Converts all `.odt` files to `.docx` (searches subfolders)

### Custom Output Name
```bash
python odf_to_docx_converter.py --file input.odt --output result.docx
```

## Usage Examples

### Example 1: Convert Single File
```bash
python odf_to_docx_converter.py --file "C:\Reports\report.odt"
```

**Output:**
```
============================================================
Converting: report.odt
============================================================
Loading ODF: report.odt
Converting content...
  Converted 45 paragraphs
  Converted 2 tables
Saving to: report.docx
✓ Conversion complete!
============================================================
```

### Example 2: Batch Convert Folder
```bash
python odf_to_docx_converter.py --folder "C:\Arabic Documents"
```

**Output:**
```
Found 5 .odt file(s) to convert
============================================================

Converting: report1.odt
✓ Conversion complete!

Converting: report2.odt
✓ Conversion complete!

...

============================================================
CONVERSION SUMMARY
============================================================
Successfully converted: 5
Errors: 0
Total files: 5
============================================================
```

### Example 3: Current Folder Only (No Recursion)
```bash
python odf_to_docx_converter.py --folder "C:\Documents" --no-recursive
```

Only converts `.odt` files in the specified folder, not subfolders.

### Example 4: Interactive Mode
```bash
python odf_to_docx_converter.py
```

Then follow the prompts:
```
ODF to DOCX Converter
============================================================
1. Convert single file
2. Convert folder

Select option (1 or 2): 1
Enter .odt file path: C:\report.odt
Enter output path (optional, press Enter to skip):
```

## What Gets Converted

| Element | Status | Notes |
|---------|--------|-------|
| Paragraphs | ✅ | Full text content preserved |
| Tables | ✅ | Structure and content preserved |
| Arabic Text | ✅ | Auto-detected, RTL applied |
| Text Direction | ✅ | Right-to-left for Arabic |
| Alignment | ✅ | Right-aligned for Arabic |
| Basic Formatting | ⚠️ | Limited (bold/italic may be lost) |
| Images | ❌ | Not yet supported |
| Complex Styles | ⚠️ | May be simplified |

## Features

### 1. Automatic RTL Detection
The converter automatically detects Arabic text (>30% Arabic characters) and applies:
- RTL (Right-to-Left) text direction
- Right alignment
- Proper BIDI settings

### 2. Table Conversion
Tables are fully converted with:
- All rows and columns preserved
- Cell content maintained
- Arabic text properly formatted

### 3. Batch Processing
Convert entire folders of `.odt` files in one command:
- Recursive subfolder search (default)
- Non-recursive option available
- Progress tracking
- Error handling per file

## Command-Line Options

```bash
python odf_to_docx_converter.py [OPTIONS]
```

| Option | Description |
|--------|-------------|
| `--file PATH` | Convert single .odt file |
| `--folder PATH` | Convert all .odt files in folder |
| `--output PATH` | Custom output path (file mode only) |
| `--no-recursive` | Don't search subfolders (folder mode) |

## Use Cases

### Use Case 1: LibreOffice → Word Migration
You have documents in LibreOffice format and need them in Word:
```bash
python odf_to_docx_converter.py --folder "C:\LibreOffice Docs"
```

### Use Case 2: Arabic Document Sharing
Convert Arabic `.odt` files for colleagues who use Microsoft Word:
```bash
python odf_to_docx_converter.py --file "arabic_report.odt"
```

### Use Case 3: Batch Document Processing
Process an entire archive of documents:
```bash
python odf_to_docx_converter.py --folder "C:\Archive\2024"
```

### Use Case 4: Specific Output Location
Convert and save to a different location:
```bash
python odf_to_docx_converter.py --file "draft.odt" --output "C:\Final\report.docx"
```

## Integration with Format Fixer

**Workflow: Convert → Fix**

```bash
# Step 1: Convert ODF to DOCX
python odf_to_docx_converter.py --folder "C:\Documents"

# Step 2: Fix formatting in converted DOCX files
python docx_format_fixer.py "C:\Documents"
```

This two-step process:
1. Converts `.odt` → `.docx`
2. Fixes encoding, RTL, and alignment issues

**Alternative: Fix First, Then Convert**

```bash
# Step 1: Fix ODF files directly
python docx_format_fixer.py "C:\Documents"  # Processes .odt files

# Step 2: Convert fixed ODF to DOCX
python odf_to_docx_converter.py --folder "C:\Documents"
```

## Technical Details

### Libraries Used
- `odfpy` - ODF document parsing
- `python-docx` - DOCX document creation

### Conversion Process
1. **Load ODF**: Parse `.odt` XML structure
2. **Extract Content**: Get paragraphs and tables
3. **Detect Language**: Check for Arabic characters
4. **Create DOCX**: Build equivalent Word document
5. **Apply Formatting**: Set RTL/alignment for Arabic
6. **Save**: Write `.docx` file

### Limitations

1. **Image Handling**: Images are not yet converted
2. **Complex Formatting**: Some styles may be simplified
3. **Custom Fonts**: Font information may be lost
4. **Document Properties**: Metadata not preserved
5. **Headers/Footers**: Not yet supported

These limitations don't affect the core text content and structure.

## Troubleshooting

### "Module not found: odf"
```bash
pip install odfpy
```

### "File not found" Error
- Check file path is correct
- Use absolute paths
- Ensure file has `.odt` extension

### Converted File Missing Content
- Check original `.odt` file opens correctly
- Verify it's a valid ODF document
- Try opening in LibreOffice first

### Arabic Text Wrong Direction
This shouldn't happen (auto-detected), but if it does:
1. Use the format fixer after conversion:
   ```bash
   python docx_format_fixer.py --file converted.docx
   ```

### Permission Denied
- Close any open Word/LibreOffice instances
- Check write permissions on output folder
- Try saving to a different location

## Performance

| Files | Time (approx) |
|-------|---------------|
| 1 file | < 1 second |
| 10 files | ~5 seconds |
| 100 files | ~30 seconds |

*Times vary based on document size and complexity*

## Comparison with Other Methods

| Method | Pros | Cons |
|--------|------|------|
| **This Script** | Free, batch processing, preserves Arabic | Limited formatting |
| **LibreOffice Export** | Best quality | Manual, one at a time |
| **Online Converters** | No install needed | Privacy concerns, limited batch |
| **MS Word Import** | Built-in | Poor format preservation |

## Future Enhancements

Potential additions:
- [ ] Image conversion
- [ ] Header/footer support
- [ ] Better style preservation
- [ ] Font mapping
- [ ] Progress bar for large batches
- [ ] Reverse converter (DOCX → ODF)

## Examples Gallery

### Before (ODF)
```
document.odt
- 50 paragraphs
- 2 tables
- Arabic + English text
```

### After (DOCX)
```
document.docx
✓ 50 paragraphs converted
✓ 2 tables converted
✓ Arabic text: RTL + Right-aligned
✓ English text: LTR + Left-aligned
```

## FAQ

**Q: Will the original .odt files be deleted?**
A: No, originals are never modified or deleted.

**Q: Can I convert .docx to .odt?**
A: Not yet, this script only does ODF → DOCX. (Feature request noted!)

**Q: Does it work with .ods (spreadsheets)?**
A: No, only .odt (text documents) are supported.

**Q: What about .odp (presentations)?**
A: Not supported yet.

**Q: Can I use this in my own scripts?**
A: Yes! Import the `ODFToDocxConverter` class:
```python
from odf_to_docx_converter import ODFToDocxConverter

converter = ODFToDocxConverter()
converter.convert_file('input.odt', 'output.docx')
```

## Need Help?

Run in interactive mode:
```bash
python odf_to_docx_converter.py
```

Or test with a simple file first:
```bash
python odf_to_docx_converter.py --file test.odt
```

Check that dependencies are installed:
```bash
pip install odfpy python-docx
```

---

**Happy Converting! 🎉**
