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

The application is divided into three main tabs:

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
- **Content**: The main text of the document.

### 3. ODF Converter
Converts OpenDocument (`.odt`) files to Microsoft Word (`.docx`).
- **Convert Folder**: Batch convert all `.odt` files in a folder.
- **Convert Single File**: Convert a specific `.odt` file.

## 📝 Notes
- The "Log Output" section at the bottom displays progress and messages from the tools.
- Long-running operations run in the background to keep the interface responsive.
