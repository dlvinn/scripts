# Arabic Word Document Tools

Two powerful Python tools for handling Arabic Word documents (.docx) with proper RTL (right-to-left) formattin

## 📋 Tools Included

1. **docx_format_fixer.py** - Fixes formatting issues in existing Arabic documents
2. **docx_generator.py** - Generates professional Arabic documents from scratch

---

## 🚀 Quick Start

### Installation

1. **Install Python** (if not already installed)

   - Download from [python.org](https://www.python.org/downloads/)
   - Make sure Python 3.7+ is installed

2. **Install required libraries**

   ```bash
   pip install -r requirements.txt
   ```

   Or install manually:

   ```bash
   pip install python-docx Pillow arabic-reshaper python-bidi
   ```

---

## 🔧 Tool 1: Format Fixer

### What It Fixes

- ✅ RTL (right-to-left) text direction for Arabic paragraphs
- ✅ Bullet point formatting and spacing
- ✅ Table cell alignment for Arabic content
- ✅ Mixed English/Arabic text direction problems
- ✅ Paragraph alignment (right-aligned for Arabic)
- ✅ Font consistency (uses Arial for Arabic text)

### Usage

#### Method 1: Command Line with Argument

```bash
python docx_format_fixer.py "C:\Users\Documents\Arabic Docs"
```

#### Method 2: Command Line Interactive

```bash
python docx_format_fixer.py
```

Then enter the folder path when prompted.

#### Method 3: VS Code Terminal

1. Open VS Code
2. Open the terminal (`` Ctrl+` `` or View → Terminal)
3. Navigate to the folder containing the script
4. Run the command:
   ```bash
   python docx_format_fixer.py "path/to/your/documents"
   ```

### Example Output

```
Found 3 document(s) to process
============================================================

Processing: report_arabic.docx
✓ Fixed and saved to: report_arabic_fixed.docx
  - RTL paragraphs: 15
  - Alignments: 12
  - Fonts: 15
  - Table cells: 8
  - Bullets: 5

============================================================
SUMMARY REPORT
============================================================
Documents processed: 3
Successfully fixed: 3
Errors: 0
```

### Notes

- Original files are **never modified**
- Fixed versions are saved as `filename_fixed.docx`
- Files starting with `~$` (temp files) are ignored
- Already fixed files (`_fixed.docx`) are skipped

---

## 📄 Tool 2: Document Generator

### Features

- ✅ Professional header with company logo
- ✅ Styled company name
- ✅ Properly formatted Arabic content
- ✅ Automatic RTL formatting
- ✅ Professional signature/approval section with 3 columns:
  - الإعداد (Prepared by)
  - المراجعة (Reviewed by)
  - الاعتماد (Approved by)
- ✅ Consistent Arabic fonts and spacing

### Usage

#### Method 1: Interactive Mode (Easiest)

```bash
python docx_generator.py
```

Then follow the prompts:

1. Enter logo path (or skip)
2. Enter company name in Arabic
3. Enter document content (type 'END' when done)
4. Enter output file name

#### Method 2: Command Line Arguments

```bash
python docx_generator.py --logo logo.png --company "شركة المثال" --content "المحتوى هنا" --output document.docx
```

#### Method 3: Content from File

```bash
python docx_generator.py --logo logo.png --company "شركة المثال" --content-file content.txt --output document.docx
```

### Example Usage Scenarios

#### Scenario 1: Simple Document with Logo

```bash
python docx_generator.py --logo "C:\logos\company_logo.png" --company "شركة التقنية المتقدمة" --content "هذا تقرير شهري عن أداء الشركة" --output monthly_report.docx
```

#### Scenario 2: Document with Bullet Points

Create a file `content.txt`:

```
تقرير الجودة الشهري

النقاط الرئيسية:
• تم فحص جميع المنتجات
• المواد الخام مطابقة للمواصفات
• لا توجد مشاكل في الإنتاج
• التوصيات: الاستمرار في نفس المعايير
```

Then run:

```bash
python docx_generator.py --logo logo.png --company "شركة الجودة" --content-file content.txt --output quality_report.docx
```

#### Scenario 3: Interactive Mode Example

```bash
python docx_generator.py

# Then enter:
Enter logo image path (or press Enter to skip): C:\logos\logo.png
Enter company name (in Arabic): شركة المثال للتجارة
Enter document content (in Arabic).
You can paste multiple lines. Type 'END' on a new line when done:
بسم الله الرحمن الرحيم

تقرير الأداء السنوي

• المبيعات: نمو بنسبة 25%
• الأرباح: زيادة ملحوظة
• التوسع: فتح 3 فروع جديدة
END
Enter output file path (default: arabic_document.docx): annual_report.docx
```

---

## 🎯 Common Issues and Solutions

### Issue 1: "Module not found" Error

**Solution:**

```bash
pip install python-docx Pillow arabic-reshaper python-bidi
```

### Issue 2: Logo not appearing

**Possible causes:**

- File path is incorrect
- Image format not supported (use PNG, JPG, or JPEG)

**Solution:**

- Use absolute paths: `C:\Users\Documents\logo.png`
- Check if file exists
- Try a different image format

### Issue 3: Arabic text still showing incorrectly

**Solution:**

- Make sure you're using the **fixed** version of the document
- Try opening in Microsoft Word (not LibreOffice or Google Docs)
- Ensure the document was saved properly

### Issue 4: "Permission denied" when saving

**Possible causes:**

- File is open in Word
- Don't have write permissions

**Solution:**

- Close the file in Word
- Run as administrator (right-click → Run as administrator)
- Save to a different location

---

## 📝 Examples

### Example 1: Fix All Documents in a Folder

```bash
# Fix all documents in the current directory
python docx_format_fixer.py .

# Fix documents in a specific folder
python docx_format_fixer.py "C:\Users\YourName\Documents\Arabic Reports"
```

### Example 2: Generate Multiple Documents

Create a batch script `generate_reports.bat`:

```batch
@echo off
python docx_generator.py --logo logo.png --company "شركة المثال" --content-file report1.txt --output report1.docx
python docx_generator.py --logo logo.png --company "شركة المثال" --content-file report2.txt --output report2.docx
python docx_generator.py --logo logo.png --company "شركة المثال" --content-file report3.txt --output report3.docx
echo All reports generated!
pause
```

### Example 3: Quality Control Report Template

```bash
python docx_generator.py --logo qc_logo.png --company "قسم مراقبة الجودة" --content "
تقرير الفحص اليومي

معلومات عامة:
• التاريخ: 2025-01-15
• المفتش: أحمد محمد
• القسم: الإنتاج

نتائج الفحص:
• المنتجات المفحوصة: 100 وحدة
• المطابق للمواصفات: 98 وحدة
• المرفوض: 2 وحدة
• نسبة النجاح: 98%

التوصيات:
• الاستمرار في نفس معايير الجودة
• مراجعة عملية الإنتاج للوحدات المرفوضة
• تدريب إضافي للعاملين
" --output daily_qc_report.docx
```

---

## 🔍 Technical Details

### Supported Features

- **RTL Text Direction**: Proper right-to-left text flow
- **Arabic Fonts**: Arial (primary), Traditional Arabic (fallback)
- **Paragraph Alignment**: Right-aligned for Arabic text
- **Bullet Points**: Proper RTL bullet formatting with spacing
- **Tables**: RTL cell alignment and text direction
- **Mixed Content**: Handles documents with both Arabic and English
- **Image Handling**: Logo insertion with proper sizing
- **Professional Layout**: Margins, spacing, and formatting

### File Formats

- **Input**: .docx files (Microsoft Word 2007+)
- **Output**: .docx files (Microsoft Word 2007+)
- **Images**: PNG, JPG, JPEG (for logos)

### Limitations

- Only works with .docx format (not .doc)
- Best results in Microsoft Word (may vary in other editors)
- Large images may need resizing before use
- Complex document structures might need manual review

---

## 🆘 Getting Help

If you encounter issues:

1. **Check Python version**: `python --version` (needs 3.7+)
2. **Verify installations**: `pip list | findstr docx`
3. **Check file paths**: Use absolute paths with forward slashes or escaped backslashes
4. **Test with sample**: Try with a simple test document first

### Example Test

Create a test file:

```bash
# Create a simple test
python docx_generator.py --company "شركة الاختبار" --content "مرحبا بك" --output test.docx
```

---

## 📚 Additional Resources

- [python-docx Documentation](https://python-docx.readthedocs.io/)
- [Arabic Text in Python](https://pypi.org/project/arabic-reshaper/)
- [Pillow Documentation](https://pillow.readthedocs.io/)

---

## 💡 Tips and Tricks

1. **Batch Processing**: Process multiple documents at once by putting them in one folder
2. **Backup First**: Keep original files before fixing
3. **Logo Size**: Use logos around 200-300px wide for best results
4. **Content Files**: Save content in .txt files for reusability
5. **Templates**: Create template content files for common document types
6. **Quality Check**: Always review generated documents in Word before distribution

---

## ⚙️ Advanced Usage

### Custom Font

Edit the scripts to change the default font:

```python
# In docx_generator.py or docx_format_fixer.py
ARABIC_FONT = 'Traditional Arabic'  # Change to your preferred font
```

### Custom Margins

Modify the margin settings in `docx_generator.py`:

```python
section.top_margin = Inches(1.0)    # Increase top margin
section.bottom_margin = Inches(1.0) # Increase bottom margin
```

### Custom Logo Size

Change logo size in `docx_generator.py`:

```python
logo_run.add_picture(str(self.logo_path), width=Inches(2.0))  # Larger logo
```

---

## 📞 Support

For issues or questions:

1. Check this README thoroughly
2. Review error messages carefully
3. Test with simple examples first
4. Ensure all dependencies are installed

---

**Version**: 1.0
**Last Updated**: January 2025
**Python Version**: 3.7+
