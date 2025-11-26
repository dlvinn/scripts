#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Example Usage Script for Arabic Document Tools
Demonstrates how to use both the fixer and generator programmatically
"""

from pathlib import Path
from docx_format_fixer import ArabicDocxFixer
from docx_generator import ArabicDocumentGenerator


def example_1_fix_documents():
    """Example 1: Fix all documents in a folder"""
    print("\n" + "="*60)
    print("EXAMPLE 1: Fix Documents in a Folder")
    print("="*60)

    # Specify the folder containing documents to fix
    folder_path = "."  # Current directory (or specify your path)

    # Create fixer instance and process
    fixer = ArabicDocxFixer(folder_path)
    fixer.process_folder()


def example_2_generate_simple_document():
    """Example 2: Generate a simple document with company name and content"""
    print("\n" + "="*60)
    print("EXAMPLE 2: Generate Simple Document")
    print("="*60)

    # Define document parameters
    company_name = "شركة المثال للتجارة"
    content = """
بسم الله الرحمن الرحيم

تقرير الأداء الشهري

ملخص الأداء:
• المبيعات: زيادة بنسبة 15%
• العملاء الجدد: 25 عميل
• الأرباح: نمو ملحوظ

التوصيات:
• الاستمرار في استراتيجية التسويق الحالية
• زيادة التركيز على خدمة العملاء
• توسيع نطاق المنتجات
    """

    output_path = "example_simple_report.docx"

    # Generate document
    generator = ArabicDocumentGenerator(
        company_name=company_name,
        content=content,
        output_path=output_path
    )

    generator.generate()
    print(f"\nDocument created: {output_path}")


def example_3_generate_with_logo():
    """Example 3: Generate document with logo"""
    print("\n" + "="*60)
    print("EXAMPLE 3: Generate Document with Logo")
    print("="*60)

    # Check if logo exists
    logo_path = "logo.png"  # Update with your logo path

    if not Path(logo_path).exists():
        print(f"Note: Logo file '{logo_path}' not found. Skipping this example.")
        print("To run this example, place a logo.png file in the current directory.")
        return

    # Define document parameters
    company_name = "شركة الجودة المتميزة"
    content = """
تقرير مراقبة الجودة

معلومات عامة:
• التاريخ: 2025-01-15
• المفتش: أحمد محمد علي
• القسم: قسم الإنتاج

نتائج الفحص:
• إجمالي الوحدات المفحوصة: 100 وحدة
• الوحدات المطابقة: 98 وحدة
• الوحدات المرفوضة: 2 وحدة
• نسبة النجاح: 98%

التوصيات:
• الاستمرار في معايير الجودة الحالية
• مراجعة عملية الإنتاج
• تدريب إضافي للموظفين
    """

    output_path = "example_quality_report.docx"

    # Generate document
    generator = ArabicDocumentGenerator(
        logo_path=logo_path,
        company_name=company_name,
        content=content,
        output_path=output_path
    )

    generator.generate()
    print(f"\nDocument created: {output_path}")


def example_4_multiple_documents():
    """Example 4: Generate multiple documents from templates"""
    print("\n" + "="*60)
    print("EXAMPLE 4: Generate Multiple Documents")
    print("="*60)

    # Define templates for different document types
    templates = [
        {
            "name": "daily_report",
            "company": "شركة التقارير اليومية",
            "content": "تقرير يومي\n\n• العمليات المنجزة اليوم\n• المهام المتبقية\n• الملاحظات",
            "output": "example_daily_report.docx"
        },
        {
            "name": "meeting_minutes",
            "company": "إدارة الاجتماعات",
            "content": "محضر اجتماع\n\n• الحضور\n• جدول الأعمال\n• القرارات المتخذة\n• المهام الموكلة",
            "output": "example_meeting_minutes.docx"
        },
        {
            "name": "inspection_report",
            "company": "قسم التفتيش",
            "content": "تقرير التفتيش\n\n• الموقع\n• النتائج\n• التوصيات\n• الإجراءات المطلوبة",
            "output": "example_inspection_report.docx"
        }
    ]

    # Generate each document
    for template in templates:
        print(f"\nGenerating: {template['name']}")

        generator = ArabicDocumentGenerator(
            company_name=template["company"],
            content=template["content"],
            output_path=template["output"]
        )

        generator.generate()

    print(f"\n✓ Generated {len(templates)} documents successfully!")


def example_5_from_file():
    """Example 5: Generate document with content from file"""
    print("\n" + "="*60)
    print("EXAMPLE 5: Generate Document from File")
    print("="*60)

    # Create a sample content file
    content_file = "example_content.txt"
    sample_content = """تقرير تفصيلي

المقدمة:
هذا تقرير تفصيلي يتضمن معلومات شاملة عن الموضوع.

النقاط الرئيسية:
• النقطة الأولى: تحليل البيانات
• النقطة الثانية: النتائج والتوصيات
• النقطة الثالثة: الخطوات التالية

الخاتمة:
تم إعداد هذا التقرير بناءً على المعلومات المتاحة.
"""

    # Write sample content to file
    with open(content_file, 'w', encoding='utf-8') as f:
        f.write(sample_content)

    print(f"Created sample content file: {content_file}")

    # Read content from file
    with open(content_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # Generate document
    generator = ArabicDocumentGenerator(
        company_name="شركة التقارير التفصيلية",
        content=content,
        output_path="example_from_file.docx"
    )

    generator.generate()
    print(f"\nDocument created from file: example_from_file.docx")


def main():
    """Run all examples"""
    print("\n" + "="*60)
    print("Arabic Document Tools - Example Usage")
    print("="*60)
    print("\nThis script demonstrates various ways to use the tools.")
    print("Each example can be run independently.\n")

    # Menu
    print("Available examples:")
    print("1. Fix all documents in current folder")
    print("2. Generate simple document (no logo)")
    print("3. Generate document with logo")
    print("4. Generate multiple documents")
    print("5. Generate document from content file")
    print("6. Run all examples")
    print("0. Exit")

    choice = input("\nEnter your choice (0-6): ").strip()

    if choice == "1":
        example_1_fix_documents()
    elif choice == "2":
        example_2_generate_simple_document()
    elif choice == "3":
        example_3_generate_with_logo()
    elif choice == "4":
        example_4_multiple_documents()
    elif choice == "5":
        example_5_from_file()
    elif choice == "6":
        # Run all examples
        example_2_generate_simple_document()
        example_3_generate_with_logo()
        example_4_multiple_documents()
        example_5_from_file()
        print("\n" + "="*60)
        print("All examples completed!")
        print("="*60)
    elif choice == "0":
        print("Exiting...")
        return
    else:
        print("Invalid choice. Please run again.")

    print("\nDone! Check the generated files in the current directory.")


if __name__ == "__main__":
    main()
