#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test Number Position Fixes
Demonstrates what the enhanced fixer corrects
"""

import sys
import io

# Ensure UTF-8 output
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

import re


def is_arabic_text(text):
    """Check if text contains Arabic characters"""
    if not text:
        return False
    arabic_chars = sum(1 for char in text if '\u0600' <= char <= '\u06FF')
    return arabic_chars > len(text) * 0.3


def fix_numbered_header(text):
    """Fix numbered header patterns"""
    text = text.strip()

    # Pattern 1: "النطاق.2" → "2. النطاق" (NO SPACE before dot)
    pattern1 = r'^(.+?)\.(\d+(?:\.\d+)*)$'

    # Pattern 2: "المجال .2" → "2. المجال" (WITH SPACE before dot)
    pattern2 = r'^(.+?)\s+\.(\d+(?:\.\d+)*)$'

    match = re.match(pattern1, text) or re.match(pattern2, text)

    if match:
        arabic_text = match.group(1).strip()
        number = match.group(2).strip()

        # Check if text is actually Arabic
        if not is_arabic_text(arabic_text):
            return text, False

        # Reconstruct in proper RTL format
        new_text = f"{number}. {arabic_text}"
        return new_text, True

    return text, False


def main():
    """Test the number fixing"""
    print("=" * 70)
    print("NUMBER POSITION FIX TEST")
    print("=" * 70)

    # Your document's exact patterns
    test_cases = [
        " الغرض.1",          # From your document
        "النطاق.2",         # From your document
        "المسؤوليات.3",      # From your document
        "التحكم بالسجلات.4",  # From your document
        "المجال .5",        # With space (old pattern)
        "الإجراء .10",      # With space (old pattern)
        "example.txt",      # Should NOT be fixed (not Arabic)
        "file.1",           # Should NOT be fixed (not Arabic)
    ]

    print("\nTest Results:")
    print("-" * 70)

    for original in test_cases:
        fixed, was_fixed = fix_numbered_header(original)

        if was_fixed:
            print(f"✓ FIXED")
            print(f"  Original: {original}")
            print(f"  Fixed:    {fixed}")
            print()
        else:
            print(f"○ NOT CHANGED (correct or non-Arabic)")
            print(f"  Text: {original}")
            print()

    print("=" * 70)
    print("\nYour Document Issues - Will Be Fixed:")
    print("=" * 70)

    your_issues = [
        (" الغرض.1", "1. الغرض"),
        ("النطاق.2", "2. النطاق"),
        ("المسؤوليات.3", "3. المسؤوليات"),
        ("التحكم بالسجلات.4", "4. التحكم بالسجلات"),
    ]

    for wrong, correct in your_issues:
        fixed, was_fixed = fix_numbered_header(wrong)
        status = "✓" if fixed == correct else "✗"
        print(f"{status} {wrong.strip():25} → {fixed:25} (expected: {correct})")

    print("=" * 70)


if __name__ == "__main__":
    main()
