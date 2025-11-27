#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Add Date to ODF Footers
Adds or updates date in footers of .odt files
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


class ODFFooterModifier:
    """Modifies footers in ODF documents"""

    def __init__(self, date_format="%Y-%m-%d", date_text=None, dry_run=False):
        """
        Initialize footer modifier

        Args:
            date_format: Format for the date (default: YYYY-MM-DD)
            date_text: Custom text to add (if None, uses current date)
            dry_run: Preview without saving
        """
        self.date_format = date_format
        self.date_text = date_text
        self.dry_run = dry_run
        self.modified_count = 0
        self.skipped_count = 0
        self.errors = []

    def get_footer_text(self):
        """Get the text to add to footer"""
        if self.date_text:
            return self.date_text

        # Use current date
        current_date = datetime.now().strftime(self.date_format)
        return f"تاريخ التحديث: {current_date}"  # "Update Date: YYYY-MM-DD"

    def get_all_footers(self, doc):
        """
        Get all footer elements from document

        Args:
            doc: ODF document

        Returns:
            List of footer elements
        """
        footers = []

        # Get master styles (where footers are defined)
        try:
            master_styles = doc.masterstyles
            if not master_styles:
                return footers

            # Get all master pages
            for elem in master_styles.childNodes:
                if elem.qname == (style.STYLENS, 'master-page'):
                    # Look for footer in this master page
                    for child in elem.childNodes:
                        if child.qname == (style.STYLENS, 'footer'):
                            footers.append(child)
                        elif child.qname == (style.STYLENS, 'footer-left'):
                            footers.append(child)
        except:
            pass

        return footers

    def add_text_to_footer(self, footer, text_to_add):
        """
        Add text to a footer element

        Args:
            footer: Footer element
            text_to_add: Text to append

        Returns:
            bool: True if modified
        """
        # Get existing paragraphs in footer
        paragraphs = []
        for child in footer.childNodes:
            if child.qname == (text.TEXTNS, 'p'):
                paragraphs.append(child)

        if not paragraphs:
            # No paragraphs, create one
            new_para = text.P(text=text_to_add)
            footer.addElement(new_para)
            return True

        # Add to last paragraph (or create new one)
        last_para = paragraphs[-1]

        # Check if date already exists
        existing_text = self.get_paragraph_text(last_para)
        if "تاريخ التحديث" in existing_text or text_to_add in existing_text:
            # Already has date, skip
            return False

        # Add line break and new text
        last_para.addElement(text.LineBreak())
        last_para.addText(text_to_add)

        return True

    def get_paragraph_text(self, para):
        """Extract all text from a paragraph"""
        text_content = []

        for node in para.childNodes:
            if hasattr(node, 'data'):
                text_content.append(node.data)

        return ''.join(text_content)

    def modify_file(self, odt_path):
        """
        Modify footers in an ODF file

        Args:
            odt_path: Path to .odt file

        Returns:
            bool: True if successful
        """
        odt_path = Path(odt_path)

        print(f"\n{'='*60}")
        print(f"Processing: {odt_path.name}")
        print(f"{'='*60}")

        try:
            # Load document
            doc = load(str(odt_path))
            print(f"  ✓ Document loaded")

            # Get all footers
            footers = self.get_all_footers(doc)

            if not footers:
                print(f"  ⚠ No footers found in document")
                self.skipped_count += 1
                return False

            print(f"  ✓ Found {len(footers)} footer(s)")

            # Get text to add
            text_to_add = self.get_footer_text()
            print(f"  → Adding: {text_to_add}")

            # Modify each footer
            modified_any = False
            for i, footer in enumerate(footers, 1):
                was_modified = self.add_text_to_footer(footer, text_to_add)
                if was_modified:
                    print(f"  ✓ Modified footer #{i}")
                    modified_any = True
                else:
                    print(f"  ○ Footer #{i} already has date, skipped")

            if not modified_any:
                print(f"  → No changes needed")
                self.skipped_count += 1
                return False

            # Save if not dry run
            if self.dry_run:
                print(f"  [DRY RUN] Would save changes")
            else:
                doc.save(str(odt_path))
                print(f"  ✓ Saved changes")

            self.modified_count += 1
            return True

        except Exception as e:
            print(f"  ✗ Error: {str(e)}")
            self.errors.append(f"{odt_path.name}: {str(e)}")
            return False

    def process_folder(self, folder_path, recursive=True):
        """
        Process all .odt files in folder

        Args:
            folder_path: Path to folder
            recursive: Search subfolders
        """
        folder_path = Path(folder_path)

        if not folder_path.exists():
            print(f"Error: Folder not found: {folder_path}")
            return

        if not folder_path.is_dir():
            print(f"Error: Not a directory: {folder_path}")
            return

        print(f"\n{'='*60}")
        print(f"Add Date to ODF Footers")
        print(f"{'='*60}")

        if self.dry_run:
            print("*** DRY RUN MODE - Files will not be modified ***")

        print(f"\nSearching in: {folder_path}")
        if recursive:
            print("Mode: Recursive (including subfolders)")
        else:
            print("Mode: Current folder only")

        print(f"\nText to add: {self.get_footer_text()}")

        # Find all .odt files
        if recursive:
            odt_files = [
                f for f in folder_path.rglob("*.odt")
                if not f.name.startswith(".~") and not f.name.startswith("~$")
            ]
        else:
            odt_files = [
                f for f in folder_path.glob("*.odt")
                if not f.name.startswith(".~") and not f.name.startswith("~$")
            ]

        if not odt_files:
            print(f"\n✓ No .odt files found")
            return

        print(f"\nFound {len(odt_files)} .odt file(s)")
        print(f"{'='*60}")

        # Process each file
        for odt_file in odt_files:
            # Show subfolder if recursive
            if recursive:
                try:
                    rel_path = odt_file.relative_to(folder_path)
                    if rel_path.parent != Path('.'):
                        print(f"\n[{rel_path.parent}]")
                except:
                    pass

            self.modify_file(odt_file)

        # Print summary
        self.print_summary()

    def print_summary(self):
        """Print summary of operations"""
        print(f"\n{'='*60}")
        print("SUMMARY")
        print(f"{'='*60}")
        print(f"Files modified: {self.modified_count}")
        print(f"Files skipped: {self.skipped_count}")
        print(f"Errors: {len(self.errors)}")

        if self.errors:
            print(f"\nErrors encountered:")
            for error in self.errors:
                print(f"  - {error}")

        if self.dry_run:
            print(f"\n*** This was a DRY RUN - no files were modified ***")

        print(f"{'='*60}")


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Add date to footers in ODF (.odt) documents',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Preview changes (safe)
  python add_date_to_footer.py "C:\\Documents" --dry-run

  # Add current date to all footers
  python add_date_to_footer.py "C:\\Documents"

  # Custom date format
  python add_date_to_footer.py "C:\\Documents" --format "%d/%m/%Y"

  # Custom text
  python add_date_to_footer.py "C:\\Documents" --text "آخر تحديث: 2025-01-15"

  # Current folder only
  python add_date_to_footer.py "C:\\Documents" --no-recursive

Date Format Codes:
  %Y - Year (2025)
  %m - Month (01-12)
  %d - Day (01-31)
  %B - Month name (January)
  %A - Day name (Monday)

  Default: %Y-%m-%d (e.g., 2025-01-15)
        """
    )

    parser.add_argument(
        'folder_path',
        nargs='?',
        help='Path to folder containing .odt files'
    )

    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview changes without modifying files (RECOMMENDED)'
    )

    parser.add_argument(
        '--no-recursive',
        action='store_true',
        help='Do not search in subfolders'
    )

    parser.add_argument(
        '--format',
        default='%Y-%m-%d',
        help='Date format (default: YYYY-MM-DD)'
    )

    parser.add_argument(
        '--text',
        help='Custom text to add (overrides date)'
    )

    args = parser.parse_args()

    # Get folder path
    if args.folder_path:
        folder_path = args.folder_path.strip('"').strip("'")
    else:
        print("Add Date to ODF Footers")
        print("=" * 60)
        folder_path = input("Enter folder path: ").strip('"').strip("'")

        if not folder_path:
            print("Error: No folder path provided")
            sys.exit(1)

    # Create modifier
    modifier = ODFFooterModifier(
        date_format=args.format,
        date_text=args.text,
        dry_run=args.dry_run
    )

    # Process folder
    recursive = not args.no_recursive
    modifier.process_folder(folder_path, recursive=recursive)


if __name__ == "__main__":
    main()
