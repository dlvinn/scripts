#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate comprehensive Windows-1256 → Latin1 Mojibake mapping
"""

import sys
import io

# Ensure UTF-8 output on Windows console
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Common Arabic characters and their Windows-1256 codes
arabic_chars = {
    0xC1: 'ء',  # Hamza
    0xC2: 'آ',  # Alef with madda
    0xC3: 'أ',  # Alef with hamza above
    0xC4: 'ؤ',  # Waw with hamza
    0xC5: 'إ',  # Alef with hamza below
    0xC6: 'ئ',  # Ya with hamza
    0xC7: 'ا',  # Alef
    0xC8: 'ب',  # Ba
    0xC9: 'ة',  # Ta marbuta
    0xCA: 'ت',  # Ta
    0xCB: 'ث',  # Tha
    0xCC: 'ج',  # Jeem
    0xCD: 'ح',  # Ha
    0xCE: 'خ',  # Kha
    0xCF: 'د',  # Dal
    0xD0: 'ذ',  # Thal
    0xD1: 'ر',  # Ra
    0xD2: 'ز',  # Zain
    0xD3: 'س',  # Seen
    0xD4: 'ش',  # Sheen
    0xD5: 'ص',  # Sad
    0xD6: 'ض',  # Dad
    0xD8: 'ط',  # Ta
    0xD9: 'ظ',  # Zha
    0xDA: 'ع',  # Ain
    0xDB: 'غ',  # Ghain
    0xE1: 'ف',  # Fa
    0xE3: 'ق',  # Qaf
    0xE4: 'ك',  # Kaf
    0xE5: 'ل',  # Lam
    0xE6: 'م',  # Meem
    0xE4: 'ن',  # Noon
    0xE5: 'ه',  # Ha
    0xE6: 'و',  # Waw
    0xEC: 'ى',  # Alef maqsura
    0xED: 'ي',  # Ya
}

print("Mojibake Mapping Generator")
print("=" * 80)
print("\nGenera testing string-by-string approach:\n")

# Test the specific case
test_word = 'آب'  # August in Arabic
print(f"Original Arabic: {test_word}")

# Encode to Windows-1256
cp1256_bytes = test_word.encode('cp1256')
print(f"Windows-1256 bytes: {' '.join(f'{b:02X}' for b in cp1256_bytes)}")

# Decode as Latin1 (this creates the Mojibake)
try:
    mojibake = cp1256_bytes.decode('latin1')
    print(f"Mojibake (Latin1): {mojibake}")
    print(f"Mojibake chars: {' '.join(f'{c}(U+{ord(c):04X})' for c in mojibake)}")
except:
    mojibake = cp1256_bytes.decode('latin1', errors='replace')
    print(f"Mojibake (with errors): {mojibake}")

print("\n" + "=" * 80)
print("Generate Python mapping:")
print("=" * 80)

# Generate all Arabic letters
arabic_letters = 'ءآأؤإئابةتثجحخدذرزسشصضطظعغـفقكلمنهوىي'
mapping = {}

for char in arabic_letters:
    try:
        cp1256_byte = char.encode('cp1256')
        mojibake_char = cp1256_byte.decode('latin1')

        # Only single-byte characters
        if len(cp1256_byte) == 1 and len(mojibake_char) == 1:
            mapping[mojibake_char] = char
            print(f"    '{mojibake_char}': '{char}',  # {char} (U+{ord(char):04X})")
    except:
        pass
