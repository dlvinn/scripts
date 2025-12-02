#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Delete ODT Files Tool
Deletes all .odt files in a folder and its subfolders.
"""

import sys
import argparse
from pathlib import Path

def delete_odt_files(folder_path, dry_run=False, force=False):
    """Find and delete all .odt files recursively"""
    folder = Path(folder_path)

    if not folder.exists():
        print(f"Error: Folder not found: {folder}")
        return

    if not folder.is_dir():
        print(f"Error: Not a directory: {folder}")
        return

    print(f"Scanning for .odt files in: {folder} ...")
    
    # rglob("*.odt") searches recursively for .odt files
    odt_files = list(folder.rglob("*.odt"))

    if not odt_files:
        print("No .odt files found.")
        return

    # List found files
    print(f"\nFound {len(odt_files)} .odt file(s):")
    print("-" * 50)
    for f in odt_files:
        # Show path relative to the input folder for cleaner output
        try:
            print(f"  {f.relative_to(folder)}")
        except:
            print(f"  {f.name}")
    print("-" * 50)

    # If dry run, stop here
    if dry_run:
        print("\n[DRY RUN] No files were deleted.")
        return

    # Confirmation step (unless forced)
    if not force:
        confirm = input("\nAre you sure you want to PERMANENTLY delete these files? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Operation cancelled.")
            return

    # Delete process
    print("\nDeleting files...")
    deleted_count = 0
    error_count = 0

    for f in odt_files:
        try:
            f.unlink()
            print(f"✓ Deleted: {f.name}")
            deleted_count += 1
        except Exception as e:
            print(f"✗ Error deleting {f.name}: {e}")
            error_count += 1

    print("\n" + "=" * 50)
    print(f"Done. Deleted: {deleted_count} | Errors: {error_count}")
    print("=" * 50)

def main():
    parser = argparse.ArgumentParser(description='Recursively delete all .odt files')
    
    parser.add_argument('folder', nargs='?', help='Folder to scan')
    parser.add_argument('--dry-run', action='store_true', help='Show files without deleting')
    parser.add_argument('--force', action='store_true', help='Delete without asking for confirmation')

    args = parser.parse_args()

    # Get folder path
    if args.folder:
        folder_path = args.folder.strip('"').strip("'")
    else:
        # Interactive mode
        print("Delete Recursive .odt Files")
        folder_path = input("Enter folder path: ").strip().strip('"').strip("'")

    if not folder_path:
        print("No folder provided.")
        return

    delete_odt_files(folder_path, args.dry_run, args.force)

if __name__ == "__main__":
    main()