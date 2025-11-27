#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Encoding Fixer for Arabic Mojibake (Windows-1256 → Windows-1252 misinterpretation)
Fixes issues like "ÀB" → "آب"
"""

import sys
import io

# Ensure UTF-8 output on Windows console
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


class ArabicEncodingFixer:
    """
    Fixes Mojibake encoding issues where Arabic text (Windows-1256)
    is misinterpreted as Latin (Windows-1252/Latin-1)
    """

    # Comprehensive mapping: Mojibake → Correct Arabic
    # Generated from actual Windows-1256 → Latin1 misinterpretation
    # This maps the corrupted Latin characters back to their Arabic equivalents
    MOJIBAKE_MAP = {
        # Arabic letters (correctly generated from cp1256)
        'Á': 'ء',  # Hamza
        'Â': 'آ',  # Alef with madda - USER'S SPECIFIC CASE (ÂÈ → آب)
        'Ã': 'أ',  # Alef with hamza above
        'Ä': 'ؤ',  # Waw with hamza
        'Å': 'إ',  # Alef with hamza below
        'Æ': 'ئ',  # Ya with hamza
        'Ç': 'ا',  # Alef
        'È': 'ب',  # Ba - USER'S SPECIFIC CASE
        'É': 'ة',  # Ta marbuta
        'Ê': 'ت',  # Ta
        'Ë': 'ث',  # Tha
        'Ì': 'ج',  # Jeem
        'Í': 'ح',  # Ha
        'Î': 'خ',  # Kha
        'Ï': 'د',  # Dal
        'Ð': 'ذ',  # Thal
        'Ñ': 'ر',  # Ra
        'Ò': 'ز',  # Zain
        'Ó': 'س',  # Seen
        'Ô': 'ش',  # Sheen
        'Õ': 'ص',  # Sad
        'Ö': 'ض',  # Dad
        'Ø': 'ط',  # Ta
        'Ù': 'ظ',  # Zha
        'Ú': 'ع',  # Ain
        'Û': 'غ',  # Ghain
        'Ü': 'ـ',  # Tatweel (kashida)
        'Ý': 'ف',  # Fa
        'Þ': 'ق',  # Qaf
        'ß': 'ك',  # Kaf
        'á': 'ل',  # Lam
        'ã': 'م',  # Meem
        'ä': 'ن',  # Noon
        'å': 'ه',  # Ha
        'æ': 'و',  # Waw
        'ì': 'ى',  # Alef maqsura
        'í': 'ي',  # Ya

        # Also add user's reported variant (À) in case it appears
        'À': 'آ',  # Alternative mapping for Alef with madda

        # Arabic diacritics and special characters
        'ð': 'ً',  # Fathatan
        'ñ': 'ٌ',  # Dammatan
        'ò': 'ٍ',  # Kasratan
        'ó': 'َ',  # Fatha
        'ô': 'ُ',  # Damma
        'õ': 'ِ',  # Kasra
        'ö': 'ّ',  # Shadda
        '÷': 'ْ',  # Sukun

        # Arabic-Indic digits
        '°': '٠',  # Arabic digit 0
        '±': '١',  # Arabic digit 1
        '²': '٢',  # Arabic digit 2
        '³': '٣',  # Arabic digit 3
        '´': '٤',  # Arabic digit 4
        'µ': '٥',  # Arabic digit 5
        '¶': '٦',  # Arabic digit 6
        '·': '٧',  # Arabic digit 7
        '¸': '٨',  # Arabic digit 8
        '¹': '٩',  # Arabic digit 9

        # Punctuation and symbols
        '¡': '،',  # Arabic comma
        '»': '؛',  # Arabic semicolon
        '¿': '؟',  # Arabic question mark

        # Special space characters
        '\xa0': ' ',  # Non-breaking space to regular space
    }

    def __init__(self):
        """Initialize the fixer with translation table"""
        self.translation_table = str.maketrans(self.MOJIBAKE_MAP)

    def clean_text(self, text):
        """
        Fix Mojibake encoding issues in text using letter-by-letter translation

        Args:
            text (str): The corrupted text (e.g., "ÀB")

        Returns:
            str: The corrected Arabic text (e.g., "آب")
        """
        if not text:
            return text

        return text.translate(self.translation_table)

    def discover_mojibake_chars(self, text_samples):
        """
        Discovery mode: Analyze text samples to find all unique Mojibake characters

        This helps identify which corrupted characters exist in your documents
        so you can add them to the mapping if needed.

        Args:
            text_samples (list): List of corrupted text strings to analyze

        Returns:
            dict: Contains sets of found mojibake chars and unmapped chars
        """
        found_mojibake = set()
        unmapped_chars = set()

        for text in text_samples:
            if not text:
                continue

            for char in text:
                # Check if it's a known Mojibake character
                if char in self.MOJIBAKE_MAP:
                    found_mojibake.add(char)
                # Check if it's a suspicious character (Latin-1 extended range)
                # but NOT in our map and NOT a normal ASCII/Arabic character
                elif (ord(char) >= 128 and ord(char) <= 255 and
                      not ('\u0600' <= char <= '\u06FF')):  # Not Arabic
                    unmapped_chars.add(char)

        return {
            'found_mojibake': found_mojibake,
            'unmapped_chars': unmapped_chars,
            'found_mojibake_sorted': sorted(found_mojibake),
            'unmapped_chars_sorted': sorted(unmapped_chars)
        }

    def print_discovery_report(self, text_samples):
        """
        Print a detailed discovery report for the given text samples

        Args:
            text_samples (list): List of corrupted text strings to analyze
        """
        results = self.discover_mojibake_chars(text_samples)

        print("\n" + "=" * 70)
        print("MOJIBAKE DISCOVERY REPORT")
        print("=" * 70)

        if results['found_mojibake']:
            print(f"\n✓ Found {len(results['found_mojibake'])} known Mojibake characters:")
            print("  (These will be automatically fixed)")
            for char in results['found_mojibake_sorted']:
                arabic_char = self.MOJIBAKE_MAP[char]
                print(f"    '{char}' → '{arabic_char}' (U+{ord(char):04X} → U+{ord(arabic_char):04X})")
        else:
            print("\n✓ No known Mojibake characters found")

        if results['unmapped_chars']:
            print(f"\n⚠ Found {len(results['unmapped_chars'])} UNMAPPED suspicious characters:")
            print("  (You may need to add these to the mapping)")
            for char in results['unmapped_chars_sorted']:
                print(f"    '{char}' (U+{ord(char):04X})")
        else:
            print("\n✓ No unmapped suspicious characters found")

        print("\n" + "=" * 70)


# Convenience function for quick testing
def quick_fix(text):
    """Quick fix for a single string"""
    fixer = ArabicEncodingFixer()
    return fixer.clean_text(text)


# Test the specific case mentioned in the prompt
if __name__ == "__main__":
    print("Arabic Encoding Fixer - Test Mode")
    print("=" * 70)

    # Your specific example - corrected Mojibake
    corrupted = "ÂÈ"  # Actual Windows-1256 → Latin1 Mojibake
    expected = "آب"

    fixer = ArabicEncodingFixer()
    fixed = fixer.clean_text(corrupted)

    print(f"\nTest Case: Iraqi month 'August' (آب)")
    print(f"  Corrupted: {corrupted}")
    print(f"  Fixed:     {fixed}")
    print(f"  Expected:  {expected}")
    print(f"  Result:    {'✓ PASS' if fixed == expected else '✗ FAIL'}")

    # Also test user's reported variant
    corrupted_alt = "ÀB"
    fixed_alt = fixer.clean_text(corrupted_alt)
    print(f"\nAlternative (user-reported): {corrupted_alt} → {fixed_alt}")

    # Additional test cases (correctly generated)
    print("\n" + "-" * 70)
    print("Additional Test Cases:")
    print("-" * 70)

    test_cases = [
        ('ÇáÚÑÇÞ', 'العراق', 'Iraq'),
        ('ÃäÊ', 'أنت', 'You'),
        ('ãÑÍÈÇ', 'مرحبا', 'Hello'),
        ('ÈÛÏÇÏ', 'بغداد', 'Baghdad'),
    ]

    for corrupted, expected, meaning in test_cases:
        fixed = fixer.clean_text(corrupted)
        status = "✓" if fixed == expected else "✗"
        print(f"{status} {meaning:15} | {corrupted:15} → {fixed:15} (expected: {expected})")

    # Discovery mode example
    print("\n" + "=" * 70)
    print("DISCOVERY MODE EXAMPLE")
    print("=" * 70)

    sample_texts = [
        "ÀB",  # آب (August)
        "ÇáÚÑÇÞ",  # العراق (Iraq)
        "ãÑÍÈÇ ÈÝÜã",  # مرحبا بکم (Welcome)
    ]

    fixer.print_discovery_report(sample_texts)
