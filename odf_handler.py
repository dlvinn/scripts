#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ODF Format Handler
Handles fixing Arabic ODF (.odt) document formatting issues
"""

from pathlib import Path
from odf import opendocument, text, table, style
from odf.opendocument import load
from encoding_fixer import ArabicEncodingFixer


class ODFHandler:
    """Handles ODF document processing with encoding, RTL, and alignment fixes"""

    def __init__(self, fix_encoding=True):
        """Initialize ODF handler"""
        self.fix_encoding = fix_encoding
        self.encoding_fixer = ArabicEncodingFixer() if fix_encoding else None
        self.doc = None
        self.doc_fixes = {
            'rtl_paragraphs': 0,
            'alignments': 0,
            'encoding': 0,
            'table_cells': 0,
        }

    def load_document(self, doc_path):
        """Load an ODF document"""
        self.doc = load(str(doc_path))
        return self.doc

    def is_arabic_text(self, text):
        """Check if text contains Arabic characters"""
        if not text:
            return False
        arabic_chars = sum(1 for char in text if '\u0600' <= char <= '\u06FF')
        return arabic_chars > len(text) * 0.3

    def fix_text_encoding(self, text_content):
        """Fix Mojibake encoding in text"""
        if not self.fix_encoding or not self.encoding_fixer:
            return text_content

        if not text_content:
            return text_content

        original = str(text_content)
        fixed = self.encoding_fixer.clean_text(original)

        if fixed != original:
            self.doc_fixes['encoding'] += 1

        return fixed

    def fix_paragraph(self, para):
        """Fix a paragraph element (encoding + RTL + alignment)"""
        # Fix encoding in all text nodes
        for node in para.childNodes:
            if hasattr(node, 'data'):
                original_text = node.data
                fixed_text = self.fix_text_encoding(original_text)
                if fixed_text != original_text:
                    node.data = fixed_text

        # Get or create style for this paragraph
        style_name = para.getAttribute('stylename')

        # Create a new automatic style with RTL and right alignment
        if style_name:
            # We'll apply RTL and alignment at the document level
            # ODF handles this through automatic styles
            pass

        self.doc_fixes['rtl_paragraphs'] += 1
        self.doc_fixes['alignments'] += 1

    def create_rtl_paragraph_style(self):
        """Create automatic paragraph style with RTL and right alignment"""
        # Create automatic style
        auto_style = style.Style(name="ArabicRTL", family="paragraph")

        # Paragraph properties
        para_props = style.ParagraphProperties()
        para_props.setAttribute('writingmode', 'rl-tb')  # Right-to-left, top-to-bottom
        para_props.setAttribute('textalign', 'end')  # Right align for RTL

        auto_style.addElement(para_props)

        # Add to automatic styles
        self.doc.automaticstyles.addElement(auto_style)

        return auto_style

    def fix_all_paragraphs(self):
        """Fix all paragraphs in the document"""
        # Get all paragraph elements
        paragraphs = self.doc.getElementsByType(text.P)

        for para in paragraphs:
            self.fix_paragraph(para)

    def fix_all_tables(self):
        """Fix all tables in the document"""
        tables = self.doc.getElementsByType(table.Table)

        for tbl in tables:
            # Fix all cells in the table
            for row in tbl.getElementsByType(table.TableRow):
                for cell in row.getElementsByType(table.TableCell):
                    # Fix paragraphs inside cells
                    for para in cell.getElementsByType(text.P):
                        self.fix_paragraph(para)
                    self.doc_fixes['table_cells'] += 1

    def apply_rtl_styles(self):
        """Apply RTL and right-alignment styles to the document"""
        # Create default RTL style if needed
        rtl_style = self.create_rtl_paragraph_style()

        # Apply to all paragraphs
        paragraphs = self.doc.getElementsByType(text.P)
        for para in paragraphs:
            # Set the style
            para.setAttribute('stylename', rtl_style.getAttribute('name'))

    def get_document_summary(self):
        """Get summary of document content"""
        paragraphs = self.doc.getElementsByType(text.P)
        tables = self.doc.getElementsByType(table.Table)

        total_text = 0
        for para in paragraphs:
            for node in para.childNodes:
                if hasattr(node, 'data'):
                    total_text += len(node.data)

        return {
            'paragraph_count': len(paragraphs),
            'total_text_length': total_text,
            'table_count': len(tables),
        }

    def fix_document(self, doc_path, output_path):
        """
        Fix an ODF document

        Args:
            doc_path: Path to input .odt file
            output_path: Path to output .odt file

        Returns:
            dict: Statistics about fixes applied
        """
        print(f"\nProcessing ODF: {doc_path.name}")

        # Load document
        self.load_document(doc_path)

        # Get original summary
        original_summary = self.get_document_summary()
        print(f"  Original: {original_summary['paragraph_count']} paragraphs, "
              f"{original_summary['total_text_length']} chars, "
              f"{original_summary['table_count']} tables")

        # Reset fix counters
        self.doc_fixes = {
            'rtl_paragraphs': 0,
            'alignments': 0,
            'encoding': 0,
            'table_cells': 0,
        }

        # Apply fixes
        self.fix_all_paragraphs()
        self.fix_all_tables()
        self.apply_rtl_styles()

        # Validate content preservation
        fixed_summary = self.get_document_summary()
        content_preserved = (
            original_summary['paragraph_count'] == fixed_summary['paragraph_count'] and
            original_summary['table_count'] == fixed_summary['table_count']
        )

        if not content_preserved:
            print(f"  ⚠ WARNING: Content mismatch detected!")
            print(f"  After fix: {fixed_summary['paragraph_count']} paragraphs, "
                  f"{fixed_summary['table_count']} tables")

        # Save fixed document
        self.doc.save(str(output_path))
        print(f"✓ Fixed and saved to: {output_path.name}")

        # Print statistics
        print(f"  - RTL paragraphs: {self.doc_fixes['rtl_paragraphs']}")
        print(f"  - Alignments: {self.doc_fixes['alignments']}")
        print(f"  - Encoding fixes: {self.doc_fixes['encoding']}")
        print(f"  - Table cells: {self.doc_fixes['table_cells']}")

        if content_preserved:
            print(f"  ✓ Content validation: PASSED")
        else:
            print(f"  ✗ Content validation: FAILED")

        return self.doc_fixes


def test_odf_handler():
    """Test the ODF handler"""
    print("ODF Handler Test")
    print("=" * 60)

    handler = ODFHandler(fix_encoding=True)

    # Test encoding fix
    test_text = "ÂÈ ÇáÚÑÇÞ"
    fixed_text = handler.fix_text_encoding(test_text)
    print(f"Encoding test:")
    print(f"  Original: {test_text}")
    print(f"  Fixed:    {fixed_text}")


if __name__ == "__main__":
    test_odf_handler()
