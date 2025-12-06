# Arabic Document Tools GUI

A user-friendly graphical interface for the Arabic Document Tools suite.

## 🚀 Getting Started

### Prerequisites
Ensure you have the required dependencies installed:
```bash
pip install -r requirements.txt
```
You also need `tkinter` installed (usually included with Python).

### Running the App
To launch the GUI, run:
```bash
python gui_app.py
```

## 🛠 Features

The application is divided into several tabs:

### 1. Format Fixer
Fixes formatting issues in existing `.docx` and `.odt` files.
- **Select Folder**: Choose the folder containing your documents.
- **Recursive**: Check to include subfolders.
- **Dry Run**: Preview changes without modifying files.
- **Fix Encoding**: Fix Mojibake (garbled text) issues.

### 2. Document Generator
Creates new professional Arabic documents.
- **Logo**: (Optional) Path to a company logo image.
- **Company Name**: Name of the company/organization.
- **Output Filename**: Name of the generated file.
- **Content**: The main text of the document. You can type directly or load from a `.txt` file.

### 3. ODF Converter
Converts OpenDocument (`.odt`) files to Microsoft Word (`.docx`).
- **Convert Folder**: Batch convert all `.odt` files in a folder.
- **Convert Single File**: Convert a specific `.odt` file.

### 4. Mojibake Scanner
Scans `.docx` files for encoding issues (Mojibake, e.g., "ÃÚãÇá" instead of "أعمال") without modifying them.
- **Recursive**: Scan subfolders.
- Provides a detailed report of corrupted texts.

### 5. Footer Modifier
Adds or updates dates/text in the footers of `.odt` files.
- **Date Format**: Custom format (e.g., `%Y-%m-%d`).
- **Custom Text**: Optional text to replace the date.
- **Dry Run**: Preview without saving changes.

### 6. Tools
Maintenance utilities.
- **Replace Original with Fixed**: Replaces the original files with their `_fixed` versions and removes the `_fixed` suffix. **Warning**: This permanently deletes the original files. Always use Dry Run first.

## 📝 Notes
- The "Log Output" section at the bottom displays progress and messages from the tools.
- Long-running operations run in the background to keep the interface responsive.
