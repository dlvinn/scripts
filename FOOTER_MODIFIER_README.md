# Add Date to ODF Footers

## Overview

Automatically adds a date (or custom text) to the footers of all `.odt` (OpenDocument) files in a folder.

**Perfect for:**
- Adding "Last Updated" dates to documents
- Batch updating footer information
- Document version tracking
- Compliance requirements

## Quick Start

### Preview Changes (SAFE)
```bash
python add_date_to_footer.py "C:\Documents" --dry-run
```

### Add Current Date to Footers
```bash
python add_date_to_footer.py "C:\Documents"
```

**Result:** Adds `تاريخ التحديث: 2025-01-15` to all footers

## Features

✅ **Preserves Existing Content** - Adds date without removing current footer
✅ **Smart Detection** - Skips files that already have the date
✅ **Batch Processing** - Process entire folders at once
✅ **Custom Formats** - Use any date format you want
✅ **Custom Text** - Add any text, not just dates
✅ **Dry Run Mode** - Preview before making changes

## Usage Examples

### Example 1: Add Current Date (Default)
```bash
python add_date_to_footer.py "C:\Reports"
```

**Adds:**
```
تاريخ التحديث: 2025-01-15
```
(Arabic: "Update Date: 2025-01-15")

### Example 2: Custom Date Format
```bash
python add_date_to_footer.py "C:\Reports" --format "%d/%m/%Y"
```

**Adds:**
```
تاريخ التحديث: 15/01/2025
```

### Example 3: Custom Text
```bash
python add_date_to_footer.py "C:\Reports" --text "آخر مراجعة: 2025-01-15"
```

**Adds:**
```
آخر مراجعة: 2025-01-15
```
(Arabic: "Last Review: 2025-01-15")

### Example 4: English Text
```bash
python add_date_to_footer.py "C:\Reports" --text "Last Updated: January 15, 2025"
```

**Adds:**
```
Last Updated: January 15, 2025
```

### Example 5: Current Folder Only
```bash
python add_date_to_footer.py "C:\Reports" --no-recursive
```

Only processes `.odt` files in `C:\Reports`, not subfolders.

### Example 6: Preview First (RECOMMENDED)
```bash
python add_date_to_footer.py "C:\Reports" --dry-run
```

Shows what would be changed without actually modifying files.

## Date Format Options

Use `--format` to customize the date:

| Format Code | Output | Example |
|-------------|--------|---------|
| `%Y-%m-%d` | YYYY-MM-DD | 2025-01-15 |
| `%d/%m/%Y` | DD/MM/YYYY | 15/01/2025 |
| `%m/%d/%Y` | MM/DD/YYYY | 01/15/2025 |
| `%Y/%m/%d` | YYYY/MM/DD | 2025/01/15 |
| `%d-%m-%Y` | DD-MM-YYYY | 15-01-2025 |
| `%B %d, %Y` | Month DD, YYYY | January 15, 2025 |
| `%d %B %Y` | DD Month YYYY | 15 January 2025 |
| `%A, %B %d, %Y` | Day, Month DD, YYYY | Wednesday, January 15, 2025 |

### Examples with Formats:

```bash
# American format
python add_date_to_footer.py "C:\Docs" --format "%m/%d/%Y"
# Output: تاريخ التحديث: 01/15/2025

# European format
python add_date_to_footer.py "C:\Docs" --format "%d.%m.%Y"
# Output: تاريخ التحديث: 15.01.2025

# Full date
python add_date_to_footer.py "C:\Docs" --format "%B %d, %Y"
# Output: تاريخ التحديث: January 15, 2025
```

## How It Works

### What the Script Does:

1. **Finds all .odt files** in the specified folder
2. **Loads each document** and locates footer elements
3. **Checks existing content** - skips if date already present
4. **Adds new line** with date/text to footer
5. **Saves the document** (unless --dry-run)

### Footer Structure:

**Before:**
```
Footer content here
Page 1 of 3
```

**After:**
```
Footer content here
Page 1 of 3
تاريخ التحديث: 2025-01-15
```

The date is **added on a new line** at the bottom of the footer.

## Example Output

```bash
$ python add_date_to_footer.py "C:\Reports" --dry-run

============================================================
Add Date to ODF Footers
============================================================
*** DRY RUN MODE - Files will not be modified ***

Searching in: C:\Reports
Mode: Recursive (including subfolders)

Text to add: تاريخ التحديث: 2025-01-15

Found 5 .odt file(s)
============================================================

============================================================
Processing: report1.odt
============================================================
  ✓ Document loaded
  ✓ Found 1 footer(s)
  → Adding: تاريخ التحديث: 2025-01-15
  ✓ Modified footer #1
  [DRY RUN] Would save changes

============================================================
Processing: report2.odt
============================================================
  ✓ Document loaded
  ✓ Found 1 footer(s)
  → Adding: تاريخ التحديث: 2025-01-15
  ○ Footer #1 already has date, skipped
  → No changes needed

============================================================
Processing: report3.odt
============================================================
  ✓ Document loaded
  ⚠ No footers found in document

============================================================
SUMMARY
============================================================
Files modified: 1
Files skipped: 4
Errors: 0

*** This was a DRY RUN - no files were modified ***
============================================================
```

## Command-Line Options

| Option | Description | Example |
|--------|-------------|---------|
| `folder_path` | Path to folder with .odt files | `"C:\Documents"` |
| `--dry-run` | Preview without saving | Always use first! |
| `--no-recursive` | Current folder only | Skip subfolders |
| `--format FORMAT` | Date format | `"%d/%m/%Y"` |
| `--text TEXT` | Custom text (overrides date) | `"Last Update: 2025"` |

## Use Cases

### Use Case 1: Quarterly Report Updates
```bash
python add_date_to_footer.py "C:\Reports\Q1-2025" --text "Q1 2025 - Updated: 2025-01-15"
```

### Use Case 2: Policy Documents
```bash
python add_date_to_footer.py "C:\Policies" --text "تاريخ النشر: 2025-01-15"
```
(Arabic: "Publication Date: 2025-01-15")

### Use Case 3: Monthly Archive
```bash
python add_date_to_footer.py "C:\Archive\January" --format "%B %Y"
```
Output: `تاريخ التحديث: January 2025`

### Use Case 4: Version Tracking
```bash
python add_date_to_footer.py "C:\Docs" --text "النسخة 2.0 - 2025-01-15"
```
(Arabic: "Version 2.0 - 2025-01-15")

## Safety Features

### 1. Duplicate Prevention
If a footer already contains the date text, it's **skipped automatically**:
```
  ○ Footer #1 already has date, skipped
```

### 2. Dry Run Mode
Always preview first:
```bash
python add_date_to_footer.py "C:\Docs" --dry-run
```

### 3. Original Content Preserved
The script **adds** to existing footers, never replaces them.

### 4. Error Handling
If a file can't be processed, it's reported but other files continue.

## Troubleshooting

### Issue: "No footers found"
**Cause:** Document doesn't have a footer defined

**Solution:**
- Open document in LibreOffice
- Go to Insert → Header and Footer → Footer
- Add footer manually, then run script

### Issue: "Permission Denied"
**Cause:** Files are open

**Solution:**
- Close all LibreOffice instances
- Ensure files aren't in use

### Issue: Date appears twice
**Cause:** Running script multiple times

**Solution:**
- Script should skip duplicates automatically
- Check if using different `--text` each time

### Issue: Date format wrong
**Cause:** Wrong format code

**Solution:**
- Check format codes above
- Test with `--dry-run` first

## Integration with Other Scripts

### Workflow 1: Fix → Add Footer
```bash
# 1. Fix formatting
python docx_format_fixer.py "C:\Documents"

# 2. Add date to footers
python add_date_to_footer.py "C:\Documents"
```

### Workflow 2: Convert → Add Footer
```bash
# 1. Convert DOCX to ODT
python odf_to_docx_converter.py "C:\Documents"

# 2. Add date to ODT footers
python add_date_to_footer.py "C:\Documents"
```

### Workflow 3: Complete Processing
```bash
# 1. Fix encoding and formatting
python docx_format_fixer.py "C:\Docs"

# 2. Replace with fixed versions
python replace_with_fixed.py "C:\Docs" --dry-run
python replace_with_fixed.py "C:\Docs"

# 3. Add date to footers
python add_date_to_footer.py "C:\Docs"
```

## Advanced Examples

### Example 1: Arabic Date with Custom Format
```bash
python add_date_to_footer.py "C:\Docs" --text "التاريخ: ١٥ يناير ٢٠٢٥"
```

### Example 2: Multiple Pieces of Information
```bash
python add_date_to_footer.py "C:\Docs" --text "النسخة: 2.0 | التاريخ: 2025-01-15 | القسم: الجودة"
```

### Example 3: Approval Date
```bash
python add_date_to_footer.py "C:\Docs" --text "تاريخ الاعتماد: 2025-01-15"
```
(Arabic: "Approval Date: 2025-01-15")

## Technical Details

### What Gets Modified:
- **Footer elements** in master page styles
- **All footers** (default, left, right)
- **Text content** only (formatting preserved)

### What's NOT Modified:
- Headers (only footers)
- Body content
- Page numbers
- Document structure

### File Format:
- **Input:** `.odt` (OpenDocument Text)
- **Output:** Same file, modified in place
- **Backup:** Make your own before running!

## Best Practices

1. ✅ **Always use --dry-run first**
   ```bash
   python add_date_to_footer.py "PATH" --dry-run
   ```

2. ✅ **Make backups** before processing important documents

3. ✅ **Test on small folder** first

4. ✅ **Use meaningful text** that includes context
   ```bash
   --text "Last Updated: 2025-01-15 by Quality Department"
   ```

5. ✅ **Consistent format** across all documents
   ```bash
   --format "%Y-%m-%d"  # Use same format for all runs
   ```

## FAQ

**Q: Will this work on DOCX files?**
A: No, only `.odt` (OpenDocument) files. Use LibreOffice to convert DOCX → ODT first.

**Q: Can I remove dates after adding them?**
A: Not automatically. You'd need to manually edit or create a removal script.

**Q: What if document has multiple footers?**
A: Script adds date to **all footers** in the document.

**Q: Does it modify headers too?**
A: No, only footers. Headers are left unchanged.

**Q: Can I add today's date automatically?**
A: Yes! Just don't use `--text`, and it uses current date.

**Q: What if I run it twice?**
A: It checks for duplicates and skips files that already have the date.

**Q: Can I use images in footer?**
A: No, only text is supported.

## Quick Reference

```bash
# Basic usage
python add_date_to_footer.py "FOLDER"

# With options
python add_date_to_footer.py "FOLDER" --dry-run --format "%d/%m/%Y"

# Custom text
python add_date_to_footer.py "FOLDER" --text "Your text here"

# Current folder only
python add_date_to_footer.py "FOLDER" --no-recursive
```

---

**Add dates to footers easily! 📅**
