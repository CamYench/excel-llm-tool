"""
Unit tests for the Excel-to-LLM Context Feeder Tool converter module.
"""

import unittest
import pandas as pd
import tempfile
import os
from pathlib import Path
import sys

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from converter import excel_to_formatted_text, excel_to_chunked_text, validate_excel_file
from utils import chunk_df, format_prompt, detect_sheets, get_data_summary, format_summary


class TestConverter(unittest.TestCase):
    """Test cases for the converter module."""
    
    def setUp(self):
        """Set up test data."""
        # Create sample DataFrame
        self.sample_data = {
            'Name': ['Alice', 'Bob', 'Charlie', 'Diana'],
            'Age': [25, 30, 35, 28],
            'City': ['New York', 'Los Angeles', 'Chicago', 'Boston'],
            'Salary': [50000, 60000, 70000, 55000]
        }
        self.df = pd.DataFrame(self.sample_data)
        
        # Create temporary Excel file
        self.temp_dir = tempfile.mkdtemp()
        self.excel_path = os.path.join(self.temp_dir, "test_data.xlsx")
        
        with pd.ExcelWriter(self.excel_path, engine='openpyxl') as writer:
            self.df.to_excel(writer, sheet_name='Sheet1', index=False)
            # Add a second sheet
            self.df.to_excel(writer, sheet_name='Sheet2', index=False)
    
    def tearDown(self):
        """Clean up test files."""
        if os.path.exists(self.excel_path):
            os.remove(self.excel_path)
        if os.path.exists(self.temp_dir):
            os.rmdir(self.temp_dir)
    
    def test_validate_excel_file(self):
        """Test Excel file validation."""
        # Valid Excel file
        self.assertTrue(validate_excel_file(self.excel_path))
        
        # Invalid file
        invalid_path = os.path.join(self.temp_dir, "invalid.txt")
        with open(invalid_path, 'w') as f:
            f.write("This is not an Excel file")
        
        self.assertFalse(validate_excel_file(invalid_path))
        
        # Clean up
        os.remove(invalid_path)
    
    def test_detect_sheets(self):
        """Test sheet detection."""
        sheets = detect_sheets(self.excel_path)
        self.assertEqual(len(sheets), 2)
        self.assertIn('Sheet1', sheets)
        self.assertIn('Sheet2', sheets)
    
    def test_chunk_df(self):
        """Test DataFrame chunking."""
        # Test with small DataFrame (should return single chunk)
        chunks = chunk_df(self.df, chunk_size=10)
        self.assertEqual(len(chunks), 1)
        self.assertEqual(len(chunks[0]), 4)
        
        # Test with chunk size smaller than DataFrame
        chunks = chunk_df(self.df, chunk_size=2)
        self.assertEqual(len(chunks), 2)
        self.assertEqual(len(chunks[0]), 2)
        self.assertEqual(len(chunks[1]), 2)
    
    def test_format_prompt(self):
        """Test prompt formatting."""
        data = "Sample data"
        template = "Analyze this: [formatted_data]"
        result = format_prompt(data, template)
        expected = "Analyze this: Sample data"
        self.assertEqual(result, expected)
        
        # Test default template
        result = format_prompt(data)
        expected = "Analyze this data: Sample data"
        self.assertEqual(result, expected)
    
    def test_get_data_summary(self):
        """Test data summary generation."""
        summary = get_data_summary(self.df)
        
        self.assertEqual(summary['rows'], 4)
        self.assertEqual(summary['columns'], 4)
        self.assertIn('Name', summary['column_types'])
        self.assertIn('Age', summary['column_types'])
        self.assertIn('numeric_summary', summary)
    
    def test_format_summary(self):
        """Test summary formatting."""
        summary = get_data_summary(self.df)
        formatted = format_summary(summary)
        
        self.assertIn("=== DATA SUMMARY ===", formatted)
        self.assertIn("Rows: 4", formatted)
        self.assertIn("Columns: 4", formatted)
        self.assertIn("Name: object", formatted)
        self.assertIn("Age: int64", formatted)
    
    def test_excel_to_formatted_text_csv(self):
        """Test Excel to CSV conversion."""
        result = excel_to_formatted_text(
            self.excel_path,
            sheet_names=['Sheet1'],
            output_format='csv'
        )
        
        self.assertIn("=== SHEET: Sheet1 ===", result)
        self.assertIn("Alice,25,New York,50000", result)
        self.assertIn("Name,Age,City,Salary", result)
    
    def test_excel_to_formatted_text_json(self):
        """Test Excel to JSON conversion."""
        result = excel_to_formatted_text(
            self.excel_path,
            sheet_names=['Sheet1'],
            output_format='json'
        )
        
        self.assertIn("=== SHEET: Sheet1 ===", result)
        self.assertIn('"Name": "Alice"', result)
        self.assertIn('"Age": 25', result)
    
    def test_excel_to_formatted_text_markdown(self):
        """Test Excel to Markdown conversion."""
        result = excel_to_formatted_text(
            self.excel_path,
            sheet_names=['Sheet1'],
            output_format='markdown'
        )
        
        self.assertIn("=== SHEET: Sheet1 ===", result)
        self.assertIn("| Name | Age | City | Salary |", result)
        self.assertIn("| Alice | 25 | New York | 50000 |", result)
    
    def test_excel_to_formatted_text_multiple_sheets(self):
        """Test conversion with multiple sheets."""
        result = excel_to_formatted_text(
            self.excel_path,
            sheet_names=['Sheet1', 'Sheet2'],
            output_format='csv'
        )
        
        self.assertIn("=== SHEET: Sheet1 ===", result)
        self.assertIn("=== SHEET: Sheet2 ===", result)
        self.assertIn("Alice,25,New York,50000", result)
    
    def test_excel_to_formatted_text_with_summary(self):
        """Test conversion with summary included."""
        result = excel_to_formatted_text(
            self.excel_path,
            sheet_names=['Sheet1'],
            output_format='csv',
            include_summary=True
        )
        
        self.assertIn("=== DATA SUMMARY ===", result)
        self.assertIn("Rows: 4", result)
        self.assertIn("=== SHEET: Sheet1 ===", result)
    
    def test_excel_to_formatted_text_custom_prompt(self):
        """Test conversion with custom prompt template."""
        custom_template = "Process this dataset: [formatted_data]"
        result = excel_to_formatted_text(
            self.excel_path,
            sheet_names=['Sheet1'],
            output_format='csv',
            prompt_template=custom_template
        )
        
        self.assertIn("Process this dataset:", result)
        self.assertIn("Alice,25,New York,50000", result)
    
    def test_excel_to_chunked_text(self):
        """Test chunked text conversion."""
        chunks = excel_to_chunked_text(
            self.excel_path,
            sheet_names=['Sheet1'],
            output_format='csv',
            chunk_size=2
        )
        
        # Should have 2 chunks (4 rows / 2 rows per chunk)
        self.assertEqual(len(chunks), 2)
        
        # Check chunk headers
        self.assertIn("=== SHEET: Sheet1 - CHUNK 1/2 ===", chunks[0])
        self.assertIn("=== SHEET: Sheet1 - CHUNK 2/2 ===", chunks[1])
    
    def test_excel_to_chunked_text_with_summary(self):
        """Test chunked text conversion with summary."""
        chunks = excel_to_chunked_text(
            self.excel_path,
            sheet_names=['Sheet1'],
            output_format='csv',
            chunk_size=2,
            include_summary=True
        )
        
        # Should have 3 chunks (1 summary + 2 data chunks)
        self.assertEqual(len(chunks), 3)
        self.assertIn("=== DATA SUMMARY ===", chunks[0])
    
    def test_invalid_output_format(self):
        """Test handling of invalid output format."""
        with self.assertRaises(ValueError):
            excel_to_formatted_text(
                self.excel_path,
                sheet_names=['Sheet1'],
                output_format='invalid_format'
            )
    
    def test_empty_sheet_handling(self):
        """Test handling of empty sheets."""
        # Create Excel file with empty sheet
        empty_excel_path = os.path.join(self.temp_dir, "empty_test.xlsx")
        empty_df = pd.DataFrame()
        
        with pd.ExcelWriter(empty_excel_path, engine='openpyxl') as writer:
            empty_df.to_excel(writer, sheet_name='EmptySheet', index=False)
        
        try:
            result = excel_to_formatted_text(
                empty_excel_path,
                sheet_names=['EmptySheet'],
                output_format='csv'
            )
            
            # Should handle empty sheet gracefully
            self.assertIn("=== SHEET: EmptySheet ===", result)
            
        finally:
            os.remove(empty_excel_path)


class TestUtils(unittest.TestCase):
    """Test cases for utility functions."""
    
    def test_chunk_df_edge_cases(self):
        """Test chunk_df with edge cases."""
        # Empty DataFrame
        empty_df = pd.DataFrame()
        chunks = chunk_df(empty_df, chunk_size=5)
        self.assertEqual(len(chunks), 1)
        self.assertEqual(len(chunks[0]), 0)
        
        # DataFrame smaller than chunk size
        small_df = pd.DataFrame({'A': [1, 2]})
        chunks = chunk_df(small_df, chunk_size=5)
        self.assertEqual(len(chunks), 1)
        self.assertEqual(len(chunks[0]), 2)
    
    def test_format_prompt_edge_cases(self):
        """Test format_prompt with edge cases."""
        # Empty data
        result = format_prompt("", "Test: [formatted_data]")
        self.assertEqual(result, "Test: ")
        
        # No placeholder
        result = format_prompt("data", "No placeholder")
        self.assertEqual(result, "No placeholder")
        
        # Multiple placeholders
        result = format_prompt("data", "[formatted_data] and [formatted_data]")
        self.assertEqual(result, "data and data")


if __name__ == '__main__':
    unittest.main()
