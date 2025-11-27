#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Add Date to Single ODF File Footer
Test version for processing one file at a time
"""

import sys
import io
import argparse
from pathlib import Path
from datetime import datetime
from odf import opendocument, text, style
from odf.opendocument import load

# Ensure UTF-8 output on Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


def get_all_footers(doc):
    """Get all footer elements from document"""
    footers = []
    try:
        master_styles = doc.masterstyles
        if not master_styles:
            return footers

        for elem in master_styles.childNodes:
            if elem.qname == (style.STYLENS, 'master-page'):
                for child in elem.childNodes:
                    if child.qname == (style.STYLENS, 'footer'):
                        footers.append(child)
                    elif child.qname == (style.STYLENS, 'footer-left'):
                        footers.append(child)
    except:
        pass

    return footers


def get_paragraph_text(para):
    """Extract all text from a paragraph"""
    text_content = []
    for node in para.childNodes:
        if hasattr(node, 'data'):
            text_content.append(node.data)
    return ''.join(text_content)


def add_text_to_footer(footer, text_to_add):
    """Add text to a footer element"""
    # Get existing paragraphs
    paragraphs = []
    for child in footer.childNodes:
        if child.qname == (text.TEXTNS, 'p'):
            paragraphs.append(child)

    if not paragraphs:
        # No paragraphs, create one
        new_para = text.P(text=text_to_add)
        footer.addElement(new_para)
        return True

    # Check if already exists
    last_para = paragraphs[-1]
    existing_text = get_paragraph_text(last_para)
    if text_to_add in existing_text:
        return False

    # Add to last paragraph
    last_para.addElement(text.LineBreak())
    last_para.addText(text_to_add)
    return True


def process_single_file(file_path, text_to_add, dry_run=False):
    """Process a single ODF file"""
    file_path = Path(file_path)

    if not file_path.exists():
        print(f"✗ Error: File not found: {file_path}")
        return False

    if file_path.suffix.lower() != '.odt':
        print(f"✗ Error: File must be .odt format")
        return False

    print(f"\n{'='*60}")
    print(f"Processing: {file_path.name}")
    print(f"{'='*60}")

    try:
        # Load document
        doc = load(str(file_path))
        print(f"✓ Document loaded")

        # Get footers
        footers = get_all_footers(doc)

        if not footers:
            print(f"✗ No footers found in document")
            print(f"\nTo add a footer:")
            print(f"  1. Open in LibreOffice")
            print(f"  2. Insert → Header and Footer → Footer")
            print(f"  3. Save, then run this script again")
            return False

        print(f"✓ Found {len(footers)} footer(s)")
        print(f"\n→ Text to add: {text_to_add}")

        # Show existing footer content
        for i, footer in enumerate(footers, 1):
            paragraphs = [child for child in footer.childNodes if child.qname == (text.TEXTNS, 'p')]
            if paragraphs:
                existing = get_paragraph_text(paragraphs[-1])
                print(f"\nFooter #{i} current content:")
                print(f"  {existing[:100]}...")  # First 100 chars

        # Modify footers
        modified_any = False
        for i, footer in enumerate(footers, 1):
            was_modified = add_text_to_footer(footer, text_to_add)
            if was_modified:
                print(f"\n✓ Modified footer #{i}")
                modified_any = True
            else:
                print(f"\n○ Footer #{i} already has this text")

        if not modified_any:
            print(f"\n→ No changes needed")
            return False

        # Save
        if dry_run:
            print(f"\n[DRY RUN] Would save changes to:")
            print(f"  {file_path}")
            print(f"\n✓ Run without --dry-run to actually save")
        else:
            doc.save(str(file_path))
            print(f"\n✓ Changes saved!")
            print(f"\nTo view footer:")
            print(f"  1. Open file in LibreOffice")
            print(f"  2. Scroll to bottom of any page")
            print(f"  3. Check the footer area")

        return True

    except Exception as e:
        print(f"✗ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Add date/text to footer of a single ODF file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Preview changes (safe)
  python add_date_to_footer_single.py document.odt --text "Date: 30 May 2025" --dry-run

  # Actually add text
  python add_date_to_footer_single.py document.odt --text "Date: 30 May 2025"

  # With Arabic text
  python add_date_to_footer_single.py document.odt --text "تاريخ التحديث: 30 May 2025"
        """
    )

    parser.add_argument(
        'file',
        help='Path to .odt file'
    )

    parser.add_argument(
        '--text',
        required=True,
        help='Text to add to footer'
    )

    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview without saving'
    )

    args = parser.parse_args()

    # Process file
    file_path = args.file.strip('"').strip("'")
    success = process_single_file(file_path, args.text, dry_run=args.dry_run)

    if success:
        print(f"\n{'='*60}")
        print("SUCCESS!")
        print(f"{'='*60}")
    else:
        print(f"\n{'='*60}")
        print("NO CHANGES MADE")
        print(f"{'='*60}")


if __name__ == "__main__":
    main()
