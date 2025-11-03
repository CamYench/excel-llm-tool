#!/usr/bin/env python3
"""
Demo script for Excel-to-LLM Context Feeder Tool.
This script demonstrates the tool's capabilities with a sample Excel file.
"""

import os
import sys
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

from converter import excel_to_formatted_text, excel_to_chunked_text
from utils import detect_sheets, get_data_summary, format_summary

def demo_basic_conversion():
    """Demonstrate basic Excel to text conversion."""
    print("🚀 Demo: Basic Excel to Text Conversion")
    print("=" * 50)
    
    # Check if sample file exists
    sample_file = Path("sample_data/sample_company_data.xlsx")
    if not sample_file.exists():
        print("❌ Sample file not found. Please run 'python sample_data.py' first.")
        return
    
    print(f"📁 Using sample file: {sample_file}")
    
    # Detect sheets
    sheets = detect_sheets(sample_file)
    print(f"📋 Available sheets: {', '.join(sheets)}")
    
    # Convert to different formats
    formats = ['csv', 'json', 'markdown']
    
    for fmt in formats:
        print(f"\n📊 Converting to {fmt.upper()} format...")
        try:
            result = excel_to_formatted_text(
                sample_file,
                sheet_names=['Employee_Data'],
                output_format=fmt,
                include_summary=True
            )
            
            # Show preview
            preview = result[:500] + "..." if len(result) > 500 else result
            print(f"✅ {fmt.upper()} conversion successful!")
            print(f"📏 Output size: {len(result):,} characters")
            print(f"👀 Preview:\n{preview}")
            
        except Exception as e:
            print(f"❌ Error converting to {fmt}: {e}")
    
    print("\n" + "=" * 50)

def demo_multiple_sheets():
    """Demonstrate processing multiple sheets."""
    print("🚀 Demo: Multiple Sheet Processing")
    print("=" * 50)
    
    sample_file = Path("sample_data/sample_company_data.xlsx")
    if not sample_file.exists():
        print("❌ Sample file not found. Please run 'python sample_data.py' first.")
        return
    
    print("📊 Processing multiple sheets (Employee_Data + Financial_Data)...")
    
    try:
        result = excel_to_formatted_text(
            sample_file,
            sheet_names=['Employee_Data', 'Financial_Data'],
            output_format='csv',
            include_summary=True
        )
        
        print(f"✅ Multi-sheet conversion successful!")
        print(f"📏 Output size: {len(result):,} characters")
        
        # Count sheet headers
        sheet_count = result.count("=== SHEET:")
        print(f"📋 Processed {sheet_count} sheets")
        
        # Show summary section
        if "=== DATA SUMMARY ===" in result:
            summary_start = result.find("=== DATA SUMMARY ===")
            summary_end = result.find("=== SHEET:")
            if summary_end == -1:
                summary_end = len(result)
            
            summary = result[summary_start:summary_end]
            print(f"📈 Summary section:\n{summary}")
        
    except Exception as e:
        print(f"❌ Error in multi-sheet conversion: {e}")
    
    print("\n" + "=" * 50)

def demo_chunking():
    """Demonstrate chunking functionality."""
    print("🚀 Demo: Chunking for Large Files")
    print("=" * 50)
    
    sample_file = Path("sample_data/sample_company_data.xlsx")
    if not sample_file.exists():
        print("❌ Sample file not found. Please run 'python sample_data.py' first.")
        return
    
    print("📊 Testing chunking with small chunk size...")
    
    try:
        chunks = excel_to_chunked_text(
            sample_file,
            sheet_names=['Employee_Data'],
            output_format='csv',
            chunk_size=20,  # Small chunks for demo
            include_summary=True
        )
        
        print(f"✅ Chunking successful!")
        print(f"📦 Generated {len(chunks)} chunks")
        
        for i, chunk in enumerate(chunks):
            print(f"  Chunk {i+1}: {len(chunk):,} characters")
            if i == 0:  # Show first chunk preview
                preview = chunk[:300] + "..." if len(chunk) > 300 else chunk
                print(f"    Preview: {preview}")
        
    except Exception as e:
        print(f"❌ Error in chunking: {e}")
    
    print("\n" + "=" * 50)

def demo_custom_prompts():
    """Demonstrate custom prompt templates."""
    print("🚀 Demo: Custom Prompt Templates")
    print("=" * 50)
    
    sample_file = Path("sample_data/sample_company_data.xlsx")
    if not sample_file.exists():
        print("❌ Sample file not found. Please run 'python sample_data.py' first.")
        return
    
    custom_templates = [
        "Please analyze this company data: [formatted_data]",
        "Here is the dataset for your review: [formatted_data]",
        "Data analysis request: [formatted_data]"
    ]
    
    for i, template in enumerate(custom_templates, 1):
        print(f"📝 Template {i}: {template}")
        
        try:
            result = excel_to_formatted_text(
                sample_file,
                sheet_names=['Employee_Data'],
                output_format='csv',
                prompt_template=template,
                include_summary=False
            )
            
            # Check if template was applied
            if template.replace("[formatted_data]", "") in result:
                print(f"✅ Template applied successfully!")
                print(f"📏 Output size: {len(result):,} characters")
            else:
                print(f"⚠️ Template may not have been applied correctly")
            
        except Exception as e:
            print(f"❌ Error with template {i}: {e}")
        
        print()
    
    print("=" * 50)

def main():
    """Run all demos."""
    print("🎯 Excel-to-LLM Context Feeder Tool - Demo Suite")
    print("=" * 60)
    
    # Check if sample data exists
    sample_file = Path("sample_data/sample_company_data.xlsx")
    if not sample_file.exists():
        print("📊 Sample data not found. Generating sample Excel file...")
        try:
            import sample_data
            sample_data.main()
            print()
        except ImportError:
            print("❌ Could not import sample_data.py. Please ensure it exists.")
            return
        except Exception as e:
            print(f"❌ Error generating sample data: {e}")
            return
    
    # Run demos
    try:
        demo_basic_conversion()
        demo_multiple_sheets()
        demo_chunking()
        demo_custom_prompts()
        
        print("🎉 Demo completed successfully!")
        print("\n💡 Next steps:")
        print("   1. Try the GUI: streamlit run src/app.py")
        print("   2. Try the CLI: python src/cli.py sample_data/sample_company_data.xlsx -f json --summary")
        print("   3. Explore different formats and options")
        
    except Exception as e:
        print(f"❌ Demo failed: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
