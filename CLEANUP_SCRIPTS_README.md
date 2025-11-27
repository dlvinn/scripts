# Cleanup Scripts Documentation

## Overview

Two powerful cleanup scripts for managing your document folders:

1. **`replace_with_fixed.py`** - Replace original files with "_fixed" versions
2. **`keep_only_odf.py`** - Delete all files except ODF files

## ⚠️ IMPORTANT SAFETY WARNINGS

**THESE SCRIPTS DELETE FILES PERMANENTLY!**

- Always use `--dry-run` first to preview changes
- Make backups before running
- Cannot be undone
- Be very careful with folder paths

---

# Script 1: Replace with Fixed

## Purpose

Replaces original files with their "_fixed" versions and removes the "_fixed" suffix.

### What It Does

1. Finds all files with `_fixed` in the name (e.g., `report_fixed.docx`)
2. Checks if the original exists (e.g., `report.docx`)
3. **Deletes the original** file
4. **Renames the fixed file** to the original name

### Example

**Before:**
```
documents/
├── report.docx          (old, needs fixing)
├── report_fixed.docx    (fixed version)
├── invoice.odt          (old)
└── invoice_fixed.odt    (fixed)
```

**After:**
```
documents/
├── report.docx          (was report_fixed.docx, now renamed)
└── invoice.odt          (was invoice_fixed.odt, now renamed)
```

## Usage

### Preview Changes (RECOMMENDED FIRST STEP)
```bash
python replace_with_fixed.py "C:\Documents" --dry-run
```

This shows what **would** happen without actually changing anything.

### Actually Replace Files
```bash
python replace_with_fixed.py "C:\Documents"
```

### Current Folder Only (No Subfolders)
```bash
python replace_with_fixed.py "C:\Documents" --no-recursive
```

### Interactive Mode
```bash
python replace_with_fixed.py
```

## Example Output

```bash
$ python replace_with_fixed.py "C:\Reports" --dry-run

============================================================
Replace Original Files with Fixed Versions
============================================================
*** DRY RUN MODE - No files will be modified ***

Searching in: C:\Reports
Mode: Recursive (including subfolders)

Found 3 fixed file(s)
============================================================

  Checking: report_fixed.docx
    ✓ Found pair:
      Original: report.docx
      Fixed:    report_fixed.docx
    [DRY RUN] Would replace:
      Delete: report.docx
      Rename: report_fixed.docx → report.docx

  Checking: invoice_fixed.odt
    ✓ Found pair:
      Original: invoice.odt
      Fixed:    invoice_fixed.odt
    [DRY RUN] Would replace:
      Delete: invoice.odt
      Rename: invoice_fixed.odt → invoice.odt

  Checking: backup_fixed.docx
    ⚠ Original not found: backup.docx
    → Skipping (no original to replace)

============================================================
SUMMARY
============================================================
Files replaced: 2
Files skipped: 1
Errors: 0

*** This was a DRY RUN - no files were actually modified ***
============================================================
```

## Command-Line Options

| Option | Description |
|--------|-------------|
| `folder_path` | Path to folder containing files |
| `--dry-run` | Preview changes without modifying files |
| `--no-recursive` | Don't search subfolders |

---

# Script 2: Keep Only ODF

## Purpose

Deletes ALL files except ODF (OpenDocument) files.

### What It Does

1. Scans folder for all files
2. **Keeps** files with ODF extensions: `.odt`, `.ods`, `.odp`, `.odg`, `.odf`
3. **Deletes** everything else

### Example

**Before:**
```
documents/
├── report.odt           (ODF - KEEP)
├── report.docx          (Word - DELETE)
├── invoice.ods          (ODF - KEEP)
├── backup.pdf           (PDF - DELETE)
├── notes.txt            (Text - DELETE)
└── image.png            (Image - DELETE)
```

**After:**
```
documents/
├── report.odt           (kept)
└── invoice.ods          (kept)
```

## Usage

### Preview Changes (RECOMMENDED FIRST STEP)
```bash
python keep_only_odf.py "C:\Documents" --dry-run
```

### Actually Delete Non-ODF Files
```bash
python keep_only_odf.py "C:\Documents"
```

### Keep Additional Extensions
```bash
python keep_only_odf.py "C:\Documents" --keep-also .pdf .txt
```

This keeps `.odt`, `.ods`, `.odp`, `.odg`, `.odf`, **AND** `.pdf`, `.txt`

### Current Folder Only
```bash
python keep_only_odf.py "C:\Documents" --no-recursive
```

### Interactive Mode
```bash
python keep_only_odf.py
```

## Example Output

```bash
$ python keep_only_odf.py "C:\Documents" --dry-run

============================================================
Keep Only ODF Files
============================================================
*** DRY RUN MODE - No files will be modified ***

Searching in: C:\Documents
Mode: Recursive (including subfolders)

Keeping extensions: .odf, .odg, .odp, .ods, .odt

Found 6 file(s)
============================================================

    ✓ Keeping: report.odt
    ✗ Deleting: report.docx
      [DRY RUN] Would delete
    ✓ Keeping: invoice.ods
    ✗ Deleting: backup.pdf
      [DRY RUN] Would delete
    ✗ Deleting: notes.txt
      [DRY RUN] Would delete
    ✗ Deleting: image.png
      [DRY RUN] Would delete

============================================================
SUMMARY
============================================================
Files kept: 2
Files deleted: 4
Errors: 0

*** This was a DRY RUN - no files were actually deleted ***
============================================================
```

## Command-Line Options

| Option | Description |
|--------|-------------|
| `folder_path` | Path to folder |
| `--dry-run` | Preview changes without deleting files |
| `--no-recursive` | Don't search subfolders |
| `--keep-also EXT [EXT ...]` | Additional extensions to keep |

## ODF Extensions Kept by Default

| Extension | Type |
|-----------|------|
| `.odt` | OpenDocument Text (LibreOffice Writer) |
| `.ods` | OpenDocument Spreadsheet (LibreOffice Calc) |
| `.odp` | OpenDocument Presentation (LibreOffice Impress) |
| `.odg` | OpenDocument Graphics (LibreOffice Draw) |
| `.odf` | OpenDocument Formula |

---

# Common Workflows

## Workflow 1: Fix Documents, Then Clean Up

```bash
# Step 1: Fix all documents (creates _fixed versions)
python docx_format_fixer.py "C:\Documents"

# Step 2: Preview replacement (SAFE)
python replace_with_fixed.py "C:\Documents" --dry-run

# Step 3: Replace originals with fixed versions
python replace_with_fixed.py "C:\Documents"
```

## Workflow 2: Convert to ODF, Keep Only ODF

```bash
# Step 1: Fix and convert documents
python docx_format_fixer.py "C:\Documents"

# Step 2: Preview what will be deleted (SAFE)
python keep_only_odf.py "C:\Documents" --dry-run

# Step 3: Delete all non-ODF files
python keep_only_odf.py "C:\Documents"
```

## Workflow 3: Full Cleanup Pipeline

```bash
# Step 1: Fix encoding and formatting
python docx_format_fixer.py "C:\Documents"

# Step 2: Replace originals with fixed versions
python replace_with_fixed.py "C:\Documents" --dry-run
python replace_with_fixed.py "C:\Documents"

# Step 3: Keep only ODF files
python keep_only_odf.py "C:\Documents" --dry-run
python keep_only_odf.py "C:\Documents"
```

---

# Safety Best Practices

## 1. Always Use --dry-run First

```bash
# SAFE - Preview only
python replace_with_fixed.py "C:\Documents" --dry-run
python keep_only_odf.py "C:\Documents" --dry-run

# DANGEROUS - Actually modifies files
python replace_with_fixed.py "C:\Documents"
python keep_only_odf.py "C:\Documents"
```

## 2. Make Backups

Before running destructive operations:

```bash
# Windows
xcopy "C:\Documents" "C:\Documents_Backup" /E /I /H

# Or use File Explorer to copy the folder
```

## 3. Test on Small Folder First

```bash
# Create test folder with sample files
# Run scripts on test folder first
# Verify results
# Then run on real data
```

## 4. Check Output Carefully

Read the dry-run output line by line to ensure it's doing what you expect.

---

# Troubleshooting

## Error: "Permission Denied"

**Cause:** Files are open or locked

**Solution:**
- Close all Word/LibreOffice/Excel instances
- Check if files are in use
- Run as administrator

## Error: "File Not Found"

**Cause:** Path is incorrect

**Solution:**
- Use absolute paths
- Check spelling
- Use quotes around paths with spaces

## Files Not Being Replaced

**Cause:** No matching pairs found

**Solution:**
- Check that both `file.ext` and `file_fixed.ext` exist
- Verify file names match exactly (case-sensitive on some systems)
- Use `--dry-run` to see what the script detects

## Too Many Files Being Deleted

**Cause:** Wrong folder or forgot --dry-run

**Solution:**
- ALWAYS use `--dry-run` first
- Double-check folder path
- Use `--keep-also` to keep additional extensions

---

# FAQ

**Q: Can I undo these operations?**
A: No, deletions are permanent. Always use backups and `--dry-run`.

**Q: What happens to files in subfolders?**
A: By default, scripts search recursively. Use `--no-recursive` to disable.

**Q: Are temporary files (~$) deleted?**
A: No, temporary files are automatically skipped.

**Q: Can I keep .docx and .odt files together?**
A: Yes! Use: `python keep_only_odf.py "C:\Documents" --keep-also .docx`

**Q: What if I want to keep the "_fixed" suffix?**
A: Don't run `replace_with_fixed.py`. The fixed files will keep their names.

**Q: Can I test on a single file?**
A: No, these scripts work on folders. Create a test folder with sample files.

---

# Integration with Other Scripts

## Compatible Scripts

| Script | Purpose | Use With |
|--------|---------|----------|
| `docx_format_fixer.py` | Fix documents | Creates `_fixed` files → Use `replace_with_fixed.py` |
| `odf_to_docx_converter.py` | Convert ODF → DOCX | Run before or after cleanup |
| `discover_mojibake.py` | Find encoding issues | Run before fixing |

---

# Quick Reference

## Replace with Fixed

```bash
# Preview
python replace_with_fixed.py "PATH" --dry-run

# Execute
python replace_with_fixed.py "PATH"
```

## Keep Only ODF

```bash
# Preview
python keep_only_odf.py "PATH" --dry-run

# Execute
python keep_only_odf.py "PATH"

# Keep additional types
python keep_only_odf.py "PATH" --keep-also .pdf .txt
```

---

**Use these scripts carefully! Always preview with --dry-run first! 🛡️**
