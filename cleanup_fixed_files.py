#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cleanup Fixed/Aligned Files
Deletes all files with "_fixed" or "_aligned" in their names
Searches recursively through all subfolders
"""

import sys
from pathlib import Path
import argparse


def find_files_to_delete(folder_path, patterns=['_fixed', '_aligned']):
    """Find all files matching the patterns"""
    folder = Path(folder_path)

    if not folder.exists():
        print(f"Error: Folder not found: {folder}")
        return []

    if not folder.is_dir():
        print(f"Error: Not a directory: {folder}")
        return []

    # Find all files with patterns in their names (recursively)
    files_to_delete = []

    for pattern in patterns:
        # Search for any file containing the pattern
        matching_files = list(folder.rglob(f"*{pattern}*"))
        files_to_delete.extend(matching_files)

    # Remove duplicates and filter out directories
    files_to_delete = [f for f in set(files_to_delete) if f.is_file()]

    return sorted(files_to_delete)


def delete_files(files, dry_run=False):
    """Delete the files (or just show what would be deleted in dry-run mode)"""
    if not files:
        print("No files to delete.")
        return

    print(f"\nFound {len(files)} file(s) to delete:")
    print("=" * 70)

    for file_path in files:
        # Show relative path
        print(f"  {file_path.name} [{file_path.parent}]")

    print("=" * 70)

    if dry_run:
        print("\n[DRY RUN] No files were deleted.")
        return

    # Ask for confirmation
    print(f"\nAre you sure you want to delete these {len(files)} file(s)?")
    confirm = input("Type 'yes' to confirm: ").strip().lower()

    if confirm != 'yes':
        print("Deletion cancelled.")
        return

    # Delete files
    deleted_count = 0
    error_count = 0

    for file_path in files:
        try:
            file_path.unlink()
            deleted_count += 1
            print(f"✓ Deleted: {file_path.name}")
        except Exception as e:
            error_count += 1
            print(f"✗ Error deleting {file_path.name}: {str(e)}")

    # Summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print(f"Successfully deleted: {deleted_count}")
    print(f"Errors: {error_count}")


def main():
    parser = argparse.ArgumentParser(
        description='Delete all files with "_fixed" or "_aligned" in their names',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Preview what would be deleted (dry run)
  python cleanup_fixed_files.py "C:\\Documents\\Arabic" --dry-run

  # Actually delete files
  python cleanup_fixed_files.py "C:\\Documents\\Arabic"

  # Delete only files with "_fixed" in name
  python cleanup_fixed_files.py "C:\\Documents\\Arabic" --pattern "_fixed"

  # Delete files with custom patterns
  python cleanup_fixed_files.py "C:\\Documents" --pattern "_backup" --pattern "_temp"
        """
    )

    parser.add_argument(
        'folder_path',
        nargs='?',
        help='Path to folder to search'
    )

    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview what would be deleted without actually deleting'
    )

    parser.add_argument(
        '--pattern',
        action='append',
        help='Custom pattern to search for (can be used multiple times). Default: _fixed and _aligned'
    )

    args = parser.parse_args()

    # Get folder path
    if args.folder_path:
        folder_path = args.folder_path.strip('"').strip("'")
    else:
        # Interactive mode
        print("Cleanup Fixed/Aligned Files")
        print("=" * 70)
        folder_path = input("Enter folder path to search: ").strip().strip('"').strip("'")

        if not folder_path:
            print("Error: No folder path provided")
            sys.exit(1)

    # Determine patterns
    if args.pattern:
        patterns = args.pattern
    else:
        patterns = ['_fixed', '_aligned']

    print(f"\nSearching for files with patterns: {', '.join(patterns)}")
    print(f"Searching recursively in: {folder_path}")

    if args.dry_run:
        print("\n*** DRY RUN MODE - No files will be deleted ***")

    # Find files
    files = find_files_to_delete(folder_path, patterns)

    # Delete files
    delete_files(files, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
