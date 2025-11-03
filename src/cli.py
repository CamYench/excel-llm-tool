"""
Command Line Interface for Excel-to-LLM Context Feeder Tool.
"""

import argparse
import sys
import os
from pathlib import Path
from converter import excel_to_formatted_text, excel_to_chunked_text, validate_excel_file
from utils import detect_sheets


def main():
    """Main CLI function."""
    parser = argparse.ArgumentParser(
        description="Convert Excel files to LLM-friendly text formats",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s data.xlsx -o output.txt -f csv
  %(prog)s data.xlsx -s "Sheet1,Sheet2" -f json --summary
  %(prog)s data.xlsx -f markdown --chunk --chunk-size 3000
        """
    )
    
    # Required arguments
    parser.add_argument(
        "input_file",
        help="Input Excel file (.xlsx or .xls)"
    )
    
    # Optional arguments
    parser.add_argument(
        "-o", "--output",
        default="output.txt",
        help="Output file path (default: output.txt)"
    )
    
    parser.add_argument(
        "-f", "--format",
        choices=["csv", "json", "markdown"],
        default="csv",
        help="Output format (default: csv)"
    )
    
    parser.add_argument(
        "-s", "--sheets",
        help="Comma-separated list of sheet names (default: all sheets)"
    )
    
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=5000,
        help="Number of rows per chunk (default: 5000)"
    )
    
    parser.add_argument(
        "--chunk",
        action="store_true",
        help="Enable chunking for large files"
    )
    
    parser.add_argument(
        "--summary",
        action="store_true",
        help="Include data summary"
    )
    
    parser.add_argument(
        "--prompt-template",
        default="Analyze this data: [formatted_data]",
        help="Custom prompt template with [formatted_data] placeholder"
    )
    
    parser.add_argument(
        "--evaluate-formulas",
        action="store_true",
        help="Attempt to evaluate Excel formulas (experimental)"
    )
    
    parser.add_argument(
        "--list-sheets",
        action="store_true",
        help="List available sheets and exit"
    )
    
    parser.add_argument(
        "--validate",
        action="store_true",
        help="Validate Excel file and exit"
    )
    
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Verbose output"
    )
    
    args = parser.parse_args()
    
    # Check if input file exists
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' not found.", file=sys.stderr)
        sys.exit(1)
    
    # Validate Excel file
    if not validate_excel_file(args.input_file):
        print(f"Error: '{args.input_file}' is not a valid Excel file.", file=sys.stderr)
        sys.exit(1)
    
    if args.verbose:
        print(f"✅ Valid Excel file: {args.input_file}")
    
    # List sheets if requested
    if args.list_sheets:
        try:
            sheets = detect_sheets(args.input_file)
            print(f"Available sheets in '{args.input_file}':")
            for i, sheet in enumerate(sheets, 1):
                print(f"  {i}. {sheet}")
            sys.exit(0)
        except Exception as e:
            print(f"Error listing sheets: {e}", file=sys.stderr)
            sys.exit(1)
    
    # Validate file if requested
    if args.validate:
        print(f"✅ '{args.input_file}' is a valid Excel file.")
        sys.exit(0)
    
    # Parse sheet names
    sheet_names = None
    if args.sheets:
        sheet_names = [s.strip() for s in args.sheets.split(",")]
        if args.verbose:
            print(f"📋 Processing sheets: {', '.join(sheet_names)}")
    
    # Process the file
    try:
        if args.verbose:
            print(f"🔄 Converting '{args.input_file}' to {args.format.upper()} format...")
        
        if args.chunk:
            # Use chunked processing
            if args.verbose:
                print(f"📦 Using chunked processing with chunk size: {args.chunk_size}")
            
            chunks = excel_to_chunked_text(
                args.input_file,
                sheet_names,
                args.format,
                args.chunk_size,
                args.prompt_template,
                args.summary
            )
            
            # Combine chunks for output
            result = "\n\n".join(chunks)
            
            if args.verbose:
                print(f"✅ Generated {len(chunks)} chunks")
        else:
            # Use single processing
            result = excel_to_formatted_text(
                args.input_file,
                sheet_names,
                args.format,
                args.chunk_size,
                args.prompt_template,
                args.summary,
                args.evaluate_formulas
            )
            
            if args.verbose:
                print("✅ Conversion completed successfully")
        
        # Save to output file
        try:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(result)
            
            # Get file size
            file_size = len(result.encode('utf-8')) / 1024  # KB
            
            print(f"💾 Output saved to: {args.output}")
            print(f"📏 File size: {file_size:.2f} KB")
            print(f"📝 Characters: {len(result):,}")
            
        except Exception as e:
            print(f"Error saving output file: {e}", file=sys.stderr)
            sys.exit(1)
            
    except Exception as e:
        print(f"Error during conversion: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
