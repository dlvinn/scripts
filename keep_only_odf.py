#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Keep Only ODF Files
Deletes all files except ODF (.odt) files in a folder
"""

import os
import sys
import argparse
from pathlib import Path
import shutil


class ODFOnlyKeeper:
    """Keeps only ODF files, deletes everything else"""

    # ODF extensions to keep
    ODF_EXTENSIONS = {'.odt', '.ods', '.odp', '.odg', '.odf'}

    def __init__(self, dry_run=False, keep_extensions=None):
        """
        Initialize keeper

        Args:
            dry_run: If True, only preview changes
            keep_extensions: Additional extensions to keep (optional)
        """
        self.dry_run = dry_run
        self.deleted_count = 0
        self.kept_count = 0
        self.errors = []

        # Extensions to keep
        self.keep_extensions = set(self.ODF_EXTENSIONS)
        if keep_extensions:
            self.keep_extensions.update(keep_extensions)

    def should_keep_file(self, file_path):
        """
        Check if file should be kept

        Args:
            file_path: Path to file

        Returns:
            bool: True if should keep, False if should delete
        """
        file_path = Path(file_path)

        # Keep if extension matches
        if file_path.suffix.lower() in self.keep_extensions:
            return True

        return False

    def get_all_files(self, folder_path, recursive=True):
        """
        Get all files in folder

        Args:
            folder_path: Path to folder
            recursive: Whether to search subfolders

        Returns:
            List of file paths
        """
        folder_path = Path(folder_path)

        if recursive:
            files = [f for f in folder_path.rglob("*") if f.is_file()]
        else:
            files = [f for f in folder_path.glob("*") if f.is_file()]

        return files

    def process_file(self, file_path):
        """
        Process a single file (keep or delete)

        Args:
            file_path: Path to file

        Returns:
            str: 'kept', 'deleted', or 'error'
        """
        file_path = Path(file_path)

        # Skip temporary files
        if file_path.name.startswith('~$') or file_path.name.startswith('.~'):
            return 'kept'

        if self.should_keep_file(file_path):
            # Keep this file
            print(f"    ✓ Keeping: {file_path.name}")
            self.kept_count += 1
            return 'kept'
        else:
            # Delete this file
            print(f"    ✗ Deleting: {file_path.name}")

            if self.dry_run:
                print(f"      [DRY RUN] Would delete")
                self.deleted_count += 1
                return 'deleted'

            try:
                file_path.unlink()
                print(f"      → Deleted successfully")
                self.deleted_count += 1
                return 'deleted'

            except Exception as e:
                print(f"      → Error: {str(e)}")
                self.errors.append(f"{file_path.name}: {str(e)}")
                return 'error'

    def process_folder(self, folder_path, recursive=True):
        """
        Process all files in folder

        Args:
            folder_path: Path to folder
            recursive: Whether to search subfolders
        """
        folder_path = Path(folder_path)

        if not folder_path.exists():
            print(f"Error: Folder not found: {folder_path}")
            return

        if not folder_path.is_dir():
            print(f"Error: Not a directory: {folder_path}")
            return

        print(f"\n{'='*60}")
        print(f"Keep Only ODF Files")
        print(f"{'='*60}")

        if self.dry_run:
            print("*** DRY RUN MODE - No files will be modified ***")

        print(f"\nSearching in: {folder_path}")
        if recursive:
            print("Mode: Recursive (including subfolders)")
        else:
            print("Mode: Current folder only")

        print(f"\nKeeping extensions: {', '.join(sorted(self.keep_extensions))}")

        # Get all files
        all_files = self.get_all_files(folder_path, recursive)

        if not all_files:
            print(f"\n✓ No files found")
            return

        print(f"\nFound {len(all_files)} file(s)")
        print(f"{'='*60}")

        # Group files by directory for better display
        files_by_dir = {}
        for file_path in all_files:
            parent = file_path.parent
            if parent not in files_by_dir:
                files_by_dir[parent] = []
            files_by_dir[parent].append(file_path)

        # Process each directory
        for directory in sorted(files_by_dir.keys()):
            files = files_by_dir[directory]

            # Show directory name if recursive
            if recursive and directory != folder_path:
                try:
                    rel_path = directory.relative_to(folder_path)
                    print(f"\n[{rel_path}]")
                except:
                    print(f"\n[{directory.name}]")

            # Process files in this directory
            for file_path in sorted(files):
                self.process_file(file_path)

        # Print summary
        self.print_summary()

    def print_summary(self):
        """Print summary of operations"""
        print(f"\n{'='*60}")
        print("SUMMARY")
        print(f"{'='*60}")
        print(f"Files kept: {self.kept_count}")
        print(f"Files deleted: {self.deleted_count}")
        print(f"Errors: {len(self.errors)}")

        if self.errors:
            print(f"\nErrors encountered:")
            for error in self.errors:
                print(f"  - {error}")

        if self.dry_run:
            print(f"\n*** This was a DRY RUN - no files were actually deleted ***")

        print(f"{'='*60}")


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Keep only ODF files, delete all others',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Preview changes (safe)
  python keep_only_odf.py "C:\\Documents" --dry-run

  # Delete all non-ODF files
  python keep_only_odf.py "C:\\Documents"

  # Keep .odt and .pdf files
  python keep_only_odf.py "C:\\Documents" --keep-also .pdf

  # Current folder only (no subfolders)
  python keep_only_odf.py "C:\\Documents" --no-recursive

ODF Extensions Kept by Default:
  .odt  - OpenDocument Text
  .ods  - OpenDocument Spreadsheet
  .odp  - OpenDocument Presentation
  .odg  - OpenDocument Graphics
  .odf  - OpenDocument Formula

Safety Notes:
  - All non-ODF files are DELETED permanently
  - Always use --dry-run first to preview changes
  - Make backups before running
  - Cannot be undone!
        """
    )

    parser.add_argument(
        'folder_path',
        nargs='?',
        help='Path to folder'
    )

    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview changes without deleting files (RECOMMENDED)'
    )

    parser.add_argument(
        '--no-recursive',
        action='store_true',
        help='Do not search in subfolders'
    )

    parser.add_argument(
        '--keep-also',
        nargs='+',
        help='Additional extensions to keep (e.g., .pdf .txt)'
    )

    args = parser.parse_args()

    # Get folder path
    if args.folder_path:
        folder_path = args.folder_path.strip('"').strip("'")
    else:
        print("Keep Only ODF Files")
        print("=" * 60)
        print("⚠ WARNING: This will DELETE all non-ODF files!")
        print("⚠ Use --dry-run first to preview changes")
        print("=" * 60)
        folder_path = input("\nEnter folder path: ").strip('"').strip("'")

        if not folder_path:
            print("Error: No folder path provided")
            sys.exit(1)

        # Confirm action
        confirm = input("\n⚠ Are you sure? All non-ODF files will be DELETED! (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Operation cancelled")
            sys.exit(0)

    # Additional extensions to keep
    keep_extensions = None
    if args.keep_also:
        keep_extensions = [ext if ext.startswith('.') else f'.{ext}' for ext in args.keep_also]

    # Create keeper
    keeper = ODFOnlyKeeper(dry_run=args.dry_run, keep_extensions=keep_extensions)

    # Process folder
    recursive = not args.no_recursive
    keeper.process_folder(folder_path, recursive=recursive)


if __name__ == "__main__":
    main()
