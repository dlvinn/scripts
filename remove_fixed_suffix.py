#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Remove "_fixed" Suffix from Filenames
Simply renames files to remove the "_fixed" suffix
Does NOT delete anything - just renames
"""

import sys
import io
import argparse
from pathlib import Path

# Ensure UTF-8 output
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


def remove_fixed_suffix(file_path, dry_run=False):
    """
    Remove "_fixed" suffix from a single file

    Args:
        file_path: Path to file with "_fixed" suffix
        dry_run: Preview without actually renaming

    Returns:
        bool: True if successful
    """
    file_path = Path(file_path)

    if not file_path.exists():
        print(f"✗ Error: File not found: {file_path}")
        return False

    # Check if filename has "_fixed"
    if "_fixed" not in file_path.stem:
        print(f"\n{'='*60}")
        print(f"File: {file_path.name}")
        print(f"{'='*60}")
        print(f"✓ This file doesn't have '_fixed' in the name")
        print(f"  No changes needed")
        return False

    # Calculate new name
    new_stem = file_path.stem.replace('_fixed', '')
    new_name = f"{new_stem}{file_path.suffix}"
    new_path = file_path.parent / new_name

    print(f"\n{'='*60}")
    print(f"File: {file_path.name}")
    print(f"{'='*60}")

    # Check if target already exists
    if new_path.exists():
        print(f"⚠ WARNING: Target file already exists!")
        print(f"  Current:  {file_path.name}")
        print(f"  Target:   {new_name}")
        print(f"  Location: {new_path}")
        print(f"\n✗ Cannot rename - target file exists")
        print(f"  You need to:")
        print(f"    1. Delete or rename the existing file: {new_name}")
        print(f"    2. Then run this script again")
        return False

    print(f"Old name: {file_path.name}")
    print(f"New name: {new_name}")

    if dry_run:
        print(f"\n[DRY RUN] Would rename file")
        print(f"  From: {file_path}")
        print(f"  To:   {new_path}")
        print(f"\n✓ Run without --dry-run to actually rename")
    else:
        try:
            file_path.rename(new_path)
            print(f"\n✓ File renamed successfully!")
            print(f"  New location: {new_path}")
        except Exception as e:
            print(f"\n✗ Error renaming file: {str(e)}")
            return False

    return True


def remove_fixed_suffix_batch(folder_path, recursive=True, dry_run=False):
    """Remove "_fixed" suffix from all files in a folder"""
    folder_path = Path(folder_path)

    if not folder_path.exists():
        print(f"✗ Error: Folder not found: {folder_path}")
        return

    if not folder_path.is_dir():
        print(f"✗ Error: Not a directory: {folder_path}")
        return

    print(f"\n{'='*60}")
    print(f"Remove '_fixed' Suffix from Filenames")
    print(f"{'='*60}")

    if dry_run:
        print("*** DRY RUN MODE - Files will not be renamed ***")

    print(f"\nSearching in: {folder_path}")
    if recursive:
        print("Mode: Recursive (including subfolders)")
        fixed_files = [f for f in folder_path.rglob("*_fixed.*") if f.is_file()]
    else:
        print("Mode: Current folder only")
        fixed_files = [f for f in folder_path.glob("*_fixed.*") if f.is_file()]

    if not fixed_files:
        print(f"\n✓ No files with '_fixed' suffix found")
        return

    print(f"\nFound {len(fixed_files)} file(s) with '_fixed' suffix")
    print(f"{'='*60}")

    renamed_count = 0
    skipped_count = 0
    error_count = 0

    for file_path in fixed_files:
        # Show subfolder if recursive
        if recursive:
            try:
                rel_path = file_path.relative_to(folder_path)
                if rel_path.parent != Path('.'):
                    print(f"\n[{rel_path.parent}]")
            except:
                pass

        result = remove_fixed_suffix(file_path, dry_run=dry_run)
        if result:
            renamed_count += 1
        else:
            skipped_count += 1

    # Summary
    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    print(f"Files renamed: {renamed_count}")
    print(f"Files skipped: {skipped_count}")

    if dry_run:
        print(f"\n*** This was a DRY RUN - no files were renamed ***")

    print(f"{'='*60}")


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Remove "_fixed" suffix from filenames (just renames, no deletion)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process a folder directly (recommended)
  python remove_fixed_suffix.py "C:\\Documents"

  # Process a folder with --dry-run
  python remove_fixed_suffix.py "C:\\Documents" --dry-run

  # Legacy support for --folder
  python remove_fixed_suffix.py --folder "C:\\Documents"

  # Process a single file
  python remove_fixed_suffix.py --file "document_fixed.odt"

What it does:
  report_fixed.odt  →  report.odt
  invoice_fixed.docx  →  invoice.docx
        """
    )

    parser.add_argument(
        'path',
        nargs='?',
        default=None,
        help='Path to a folder to process.'
    )

    parser.add_argument(
        '--file',
        help='Path to single file to rename (optional)'
    )

    parser.add_argument(
        '--folder',
        help='Path to folder containing files to rename (optional, superseded by positional path)'
    )

    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview changes without renaming'
    )

    parser.add_argument(
        '--no-recursive',
        action='store_true',
        help='Do not search subfolders'
    )

    args = parser.parse_args()

    # Determine the folder path to use
    folder_to_process = args.path if args.path else args.folder

    if folder_to_process:
        # Folder mode
        folder_path = folder_to_process.strip('"').strip("'")
        recursive = not args.no_recursive
        remove_fixed_suffix_batch(folder_path, recursive=recursive, dry_run=args.dry_run)

    elif args.file:
        # Single file mode
        file_path = args.file.strip('"').strip("'")
        remove_fixed_suffix(file_path, dry_run=args.dry_run)

    else:
        # Interactive mode if no paths are provided
        print("Remove '_fixed' Suffix from Filenames")
        print("=" * 60)
        print("1. Remove from single file")
        print("2. Remove from all files in folder")
        choice = input("\nSelect option (1 or 2): ").strip()

        if choice == "1":
            file_input = input("Enter file path: ").strip().strip('"').strip("'")
            if file_input:
                remove_fixed_suffix(file_input, dry_run=False)

        elif choice == "2":
            folder_input = input("Enter folder path: ").strip().strip('"').strip("'")
            recursive_input = input("Search in subfolders? (Y/n): ").strip().lower()
            recursive = recursive_input != 'n'

            if folder_input:
                remove_fixed_suffix_batch(folder_input, recursive=recursive, dry_run=False)

        else:
            print("Invalid choice")


if __name__ == "__main__":
    main()
