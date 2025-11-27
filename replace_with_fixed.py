#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Replace Original Files with Fixed Versions
Replaces original files with their "_fixed" versions, removing the "_fixed" suffix
"""

import os
import sys
import argparse
from pathlib import Path
import shutil


class FixedFileReplacer:
    """Handles replacing original files with their fixed versions"""

    def __init__(self, dry_run=False):
        """Initialize replacer"""
        self.dry_run = dry_run
        self.replaced_count = 0
        self.skipped_count = 0
        self.errors = []

    def find_fixed_files(self, folder_path, recursive=True):
        """
        Find all files with '_fixed' suffix

        Args:
            folder_path: Path to search
            recursive: Whether to search subfolders

        Returns:
            List of paths to fixed files
        """
        folder_path = Path(folder_path)

        if recursive:
            # All extensions with _fixed suffix
            fixed_files = [
                f for f in folder_path.rglob("*_fixed.*")
                if f.is_file()
            ]
        else:
            fixed_files = [
                f for f in folder_path.glob("*_fixed.*")
                if f.is_file()
            ]

        return fixed_files

    def get_original_path(self, fixed_path):
        """
        Get the original file path from a fixed file path

        Args:
            fixed_path: Path to file with '_fixed' suffix

        Returns:
            Path to original file (without '_fixed')
        """
        fixed_path = Path(fixed_path)

        # Remove '_fixed' from stem
        original_stem = fixed_path.stem.replace('_fixed', '')
        original_path = fixed_path.parent / f"{original_stem}{fixed_path.suffix}"

        return original_path

    def replace_file(self, fixed_path):
        """
        Replace original file with fixed version

        Args:
            fixed_path: Path to fixed file

        Returns:
            bool: True if successful, False otherwise
        """
        fixed_path = Path(fixed_path)
        original_path = self.get_original_path(fixed_path)

        print(f"\n  Checking: {fixed_path.name}")

        # Check if original exists
        if not original_path.exists():
            print(f"    ⚠ Original not found: {original_path.name}")
            print(f"    → Skipping (no original to replace)")
            self.skipped_count += 1
            return False

        # Verify both files exist before proceeding
        if not fixed_path.exists():
            print(f"    ✗ Fixed file missing: {fixed_path.name}")
            self.errors.append(f"Fixed file missing: {fixed_path.name}")
            return False

        print(f"    ✓ Found pair:")
        print(f"      Original: {original_path.name}")
        print(f"      Fixed:    {fixed_path.name}")

        if self.dry_run:
            print(f"    [DRY RUN] Would replace:")
            print(f"      Delete: {original_path.name}")
            print(f"      Rename: {fixed_path.name} → {original_path.name}")
            self.replaced_count += 1
            return True

        try:
            # Delete original
            original_path.unlink()
            print(f"    → Deleted: {original_path.name}")

            # Rename fixed to original name
            fixed_path.rename(original_path)
            print(f"    → Renamed: {fixed_path.name} → {original_path.name}")
            print(f"    ✓ Success!")

            self.replaced_count += 1
            return True

        except Exception as e:
            print(f"    ✗ Error: {str(e)}")
            self.errors.append(f"{fixed_path.name}: {str(e)}")
            return False

    def process_folder(self, folder_path, recursive=True):
        """
        Process all fixed files in folder

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
        print(f"Replace Original Files with Fixed Versions")
        print(f"{'='*60}")

        if self.dry_run:
            print("*** DRY RUN MODE - No files will be modified ***")

        print(f"\nSearching in: {folder_path}")
        if recursive:
            print("Mode: Recursive (including subfolders)")
        else:
            print("Mode: Current folder only")

        # Find all fixed files
        fixed_files = self.find_fixed_files(folder_path, recursive)

        if not fixed_files:
            print(f"\n✓ No '_fixed' files found")
            return

        print(f"\nFound {len(fixed_files)} fixed file(s)")
        print(f"{'='*60}")

        # Process each fixed file
        for fixed_file in fixed_files:
            # Show subfolder if in recursive mode
            if recursive:
                try:
                    rel_path = fixed_file.relative_to(folder_path)
                    if rel_path.parent != Path('.'):
                        print(f"\n[{rel_path.parent}]")
                except:
                    pass

            self.replace_file(fixed_file)

        # Print summary
        self.print_summary()

    def print_summary(self):
        """Print summary of operations"""
        print(f"\n{'='*60}")
        print("SUMMARY")
        print(f"{'='*60}")
        print(f"Files replaced: {self.replaced_count}")
        print(f"Files skipped: {self.skipped_count}")
        print(f"Errors: {len(self.errors)}")

        if self.errors:
            print(f"\nErrors encountered:")
            for error in self.errors:
                print(f"  - {error}")

        if self.dry_run:
            print(f"\n*** This was a DRY RUN - no files were actually modified ***")

        print(f"{'='*60}")


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Replace original files with their "_fixed" versions',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Preview changes (safe)
  python replace_with_fixed.py "C:\\Documents" --dry-run

  # Replace all fixed files
  python replace_with_fixed.py "C:\\Documents"

  # Current folder only (no subfolders)
  python replace_with_fixed.py "C:\\Documents" --no-recursive

Safety Notes:
  - Original files are DELETED permanently
  - Always use --dry-run first to preview changes
  - Make backups before running
  - Cannot be undone!
        """
    )

    parser.add_argument(
        'folder_path',
        nargs='?',
        help='Path to folder containing files'
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

    args = parser.parse_args()

    # Get folder path
    if args.folder_path:
        folder_path = args.folder_path.strip('"').strip("'")
    else:
        print("Replace Original Files with Fixed Versions")
        print("=" * 60)
        print("⚠ WARNING: This will DELETE original files!")
        print("⚠ Use --dry-run first to preview changes")
        print("=" * 60)
        folder_path = input("\nEnter folder path: ").strip('"').strip("'")

        if not folder_path:
            print("Error: No folder path provided")
            sys.exit(1)

        # Confirm action
        confirm = input("\n⚠ Are you sure? Original files will be DELETED! (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Operation cancelled")
            sys.exit(0)

    # Create replacer
    replacer = FixedFileReplacer(dry_run=args.dry_run)

    # Process folder
    recursive = not args.no_recursive
    replacer.process_folder(folder_path, recursive=recursive)


if __name__ == "__main__":
    main()
