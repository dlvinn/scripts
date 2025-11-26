#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Complete Align Right Tool (with RTL/Bidi Support)
"""

import sys
import argparse
from pathlib import Path
from docx import Document
# Import specific enums for alignment
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def align_document_right(doc_path, output_path=None):
    """Align all content in a document to the right (Text + Direction + Tables + Lists)"""
    print(f"\nProcessing: {doc_path.name}")

    try:
        # Load document
        doc = Document(doc_path)
    except Exception as e:
        print(f"Error opening file: {e}")
        return None

    count = 0
    list_count = 0

    # 1. Align all regular paragraphs
    for paragraph in doc.paragraphs:
        # Visual Alignment (Right side of page)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Text Direction (Right-to-Left logic) - Crucial for Arabic/Hebrew
        # If this fails, the paragraph might be empty or read-only, so we try/except
        try:
            paragraph.paragraph_format.bidi = True
        except:
            pass

        # Fix numbered/bulleted lists - set list numbering to RTL
        if paragraph.style.name.startswith('List') or paragraph._element.pPr is not None:
            try:
                from docx.oxml.ns import qn
                from docx.oxml import OxmlElement

                pPr = paragraph._element.get_or_add_pPr()

                # Check if this paragraph has list numbering
                numPr = pPr.find(qn('w:numPr'))
                if numPr is not None:
                    # Add bidi to make list numbers appear on right
                    existing_bidi = pPr.find(qn('w:bidi'))
                    if existing_bidi is None:
                        bidi = OxmlElement('w:bidi')
                        bidi.set(qn('w:val'), '1')
                        pPr.append(bidi)
                    list_count += 1
            except:
                pass

        count += 1

    # 2. Align all tables and their contents
    for table in doc.tables:
        # Align the table container itself to the right
        try:
            table.alignment = WD_TABLE_ALIGNMENT.RIGHT
        except:
            pass

        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    try:
                        paragraph.paragraph_format.bidi = True
                    except:
                        pass

                    # Fix numbered lists in tables too
                    if paragraph.style.name.startswith('List') or paragraph._element.pPr is not None:
                        try:
                            from docx.oxml.ns import qn
                            from docx.oxml import OxmlElement

                            pPr = paragraph._element.get_or_add_pPr()
                            numPr = pPr.find(qn('w:numPr'))
                            if numPr is not None:
                                existing_bidi = pPr.find(qn('w:bidi'))
                                if existing_bidi is None:
                                    bidi = OxmlElement('w:bidi')
                                    bidi.set(qn('w:val'), '1')
                                    pPr.append(bidi)
                                list_count += 1
                        except:
                            pass

                    count += 1

    # Save
    if output_path is None:
        output_path = doc_path.parent / f"{doc_path.stem}_aligned{doc_path.suffix}"

    try:
        doc.save(output_path)
        print(f"✓ Aligned {count} paragraphs to right (with RTL support)")
        if list_count > 0:
            print(f"✓ Fixed {list_count} numbered/bulleted lists")
        print(f"✓ Saved to: {output_path.name}\n")
        return output_path
    except Exception as e:
        print(f"Error saving file: {e}")
        return None


def process_folder(folder_path, recursive=True):
    """Process all .docx files in a folder and subfolders"""
    folder = Path(folder_path)

    if not folder.exists():
        print(f"Error: Folder not found: {folder}")
        return

    if not folder.is_dir():
        print(f"Error: Not a directory: {folder}")
        return

    # Find all .docx files (recursively if enabled)
    if recursive:
        # Search in current folder and all subfolders
        docx_files = [
            f for f in folder.rglob("*.docx")
            if not f.name.startswith("~$") and "_aligned" not in f.name
        ]
        print(f"\nSearching recursively in: {folder}")
    else:
        # Search only in current folder
        docx_files = [
            f for f in folder.glob("*.docx")
            if not f.name.startswith("~$") and "_aligned" not in f.name
        ]
        print(f"\nSearching in: {folder}")

    if not docx_files:
        print(f"No .docx files found")
        return

    print(f"Found {len(docx_files)} document(s) to process")
    print("=" * 60)

    success_count = 0
    error_count = 0

    for doc_path in docx_files:
        # Show which subfolder the file is in
        try:
            rel_path = doc_path.relative_to(folder)
            if rel_path.parent != Path('.'):
                print(f"\n[{rel_path.parent}]")
        except:
            pass

        result = align_document_right(doc_path)
        if result:
            success_count += 1
        else:
            error_count += 1

    # Summary
    print("=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"Successfully aligned: {success_count}")
    print(f"Errors: {error_count}")


def process_single_file(file_path, output_path=None):
    """Process a single file"""
    file_path = Path(file_path)

    if not file_path.exists():
        print(f"Error: File not found: {file_path}")
        return

    if not file_path.suffix.lower() == '.docx':
        print(f"Error: File must be a .docx file")
        return

    if output_path:
        output_path = Path(output_path)
        
    align_document_right(file_path, output_path)


def main():
    parser = argparse.ArgumentParser(
        description='Align all content in Word documents to the right (RTL Support)',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    parser.add_argument(
        'path',
        nargs='?',
        help='Path to folder or file'
    )

    parser.add_argument(
        '--file',
        help='Process a single file instead of a folder'
    )

    parser.add_argument(
        '--output',
        help='Output file path (only for single file mode)'
    )

    args = parser.parse_args()

    # Determine mode based on arguments
    if args.file:
        # Single file mode via flag
        file_path = args.file.strip('"').strip("'")
        output_path = args.output.strip('"').strip("'") if args.output else None
        process_single_file(file_path, output_path)
        
    elif args.path:
        # Path argument provided
        path = Path(args.path.strip('"').strip("'"))
        if path.is_file():
            process_single_file(path)
        else:
            process_folder(path)
            
    else:
        # No arguments provided -> Interactive Mode
        print("Align Right Tool (RTL/Arabic Support)")
        print("=" * 60)
        print("1. Process folder (all .docx files)")
        print("2. Process single file")
        choice = input("\nSelect option (1 or 2): ").strip()

        if choice == "1":
            folder_input = input("Enter folder path: ").strip().strip('"').strip("'")
            if folder_input:
                process_folder(folder_input)
        elif choice == "2":
            file_input = input("Enter file path: ").strip().strip('"').strip("'")
            output_input = input("Enter output path (optional): ").strip().strip('"').strip("'")
            if file_input:
                process_single_file(file_input, output_input if output_input else None)
        else:
            print("Invalid choice")

if __name__ == "__main__":
    main()