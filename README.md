# Excel-to-LLM Context Feeder Tool

A powerful Python tool that converts Excel files (.xlsx, .xls) into LLM-friendly text formats (CSV, JSON, Markdown tables) with a modern Streamlit-based GUI. Perfect for developers and teams building LLM applications that need structured data as context.

## 🚀 Features

- **Multiple Output Formats**: Convert Excel to CSV, JSON, or Markdown tables
- **Smart Sheet Handling**: Process single sheets, multiple sheets, or all sheets automatically
- **Data Chunking**: Handle large files by splitting into manageable chunks
- **Custom Prompt Templates**: Use your own prompt templates with `[formatted_data]` placeholder
- **Data Summaries**: Include statistical summaries and metadata
- **Modern GUI**: Beautiful Streamlit interface for easy file upload and configuration
- **File Validation**: Automatic Excel file validation and error handling
- **Preview & Download**: Preview output and download results directly from the GUI

## 📋 Requirements

- Python 3.8+
- Streamlit 1.38+
- pandas 2.2.2+
- openpyxl 3.1.5+

## 🛠️ Installation

1. **Clone the repository:**
   ```bash
   git clone <repository-url>
   cd excel-llm-tool
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application:**
   ```bash
   streamlit run src/app.py
   ```

4. **Open your browser** and navigate to `http://localhost:8501`

## 🎯 Usage

### GUI Interface

1. **Upload Excel File**: Use the file uploader to select your `.xlsx` or `.xls` file
2. **Configure Options**: 
   - Choose output format (CSV, JSON, Markdown)
   - Select sheets to process
   - Set chunk size for large files
   - Enable/disable data summary
   - Customize prompt template
3. **Convert**: Click the convert button to process your file
4. **Download**: Save the output to a file or download directly from the GUI

### Configuration Options

- **Output Format**: Choose between CSV, JSON, or Markdown
- **Chunk Size**: Set the number of rows per chunk (1000-10000)
- **Data Summary**: Include statistical information about your data
- **Prompt Template**: Customize the output format with your own template
- **Sheet Selection**: Process specific sheets or all sheets automatically

### Advanced Features

- **Chunking**: Enable for very large files to split into manageable pieces
- **Formula Evaluation**: Experimental support for Excel formula evaluation
- **Multiple Sheets**: Process multiple sheets with automatic concatenation
- **Data Cleaning**: Automatic handling of missing values and data types

## 📁 Project Structure

```
excel-llm-tool/
├── src/
│   ├── app.py           # Streamlit GUI entry point
│   ├── converter.py     # Core conversion logic
│   └── utils.py         # Helper functions
├── tests/
│   └── test_converter.py # Unit tests
├── requirements.txt     # Python dependencies
└── README.md           # This file
```

## 🔧 API Usage

You can also use the tool programmatically:

```python
from src.converter import excel_to_formatted_text

# Convert Excel to CSV
result = excel_to_formatted_text(
    file_path="data.xlsx",
    sheet_names=["Sheet1"],
    output_format="csv",
    include_summary=True
)

print(result)
```

### Available Functions

- `excel_to_formatted_text()`: Convert Excel to single formatted text
- `excel_to_chunked_text()`: Convert Excel to multiple chunks
- `validate_excel_file()`: Validate Excel file format
- `detect_sheets()`: Get available sheet names
- `get_data_summary()`: Generate data statistics

## 🧪 Testing

Run the test suite to ensure everything works correctly:

```bash
python -m pytest tests/
```

Or run individual test files:

```bash
python tests/test_converter.py
```

## 📊 Supported Formats

### CSV Output
```
=== SHEET: Sheet1 ===
Name,Age,City,Salary
Alice,25,New York,50000
Bob,30,Los Angeles,60000
```

### JSON Output
```json
=== SHEET: Sheet1 ===
[
  {
    "Name": "Alice",
    "Age": 25,
    "City": "New York",
    "Salary": 50000
  }
]
```

### Markdown Output
```markdown
=== SHEET: Sheet1 ===
| Name | Age | City | Salary |
|------|-----|------|--------|
| Alice | 25 | New York | 50000 |
| Bob | 30 | Los Angeles | 60000 |
```

## 🚨 Error Handling

The tool includes comprehensive error handling for:
- Invalid Excel files
- Missing sheets
- File permission issues
- Memory constraints
- Invalid configuration options

## 🔮 Future Enhancements

- **LLM API Integration**: Direct integration with popular LLM APIs
- **Batch Processing**: Process multiple Excel files at once
- **Advanced Formatting**: More output format options
- **Cloud Storage**: Support for cloud-based file storage
- **Collaboration**: Real-time collaboration features

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

If you encounter any issues or have questions:

1. Check the error messages in the GUI
2. Review the console output for detailed error information
3. Ensure your Excel file is valid and not corrupted
4. Verify you have the correct Python version and dependencies

## 🎉 Quick Start Example

1. **Install and run:**
   ```bash
   pip install -r requirements.txt
   streamlit run src/app.py
   ```

2. **Upload an Excel file** with sample data
3. **Select CSV format** and enable data summary
4. **Click Convert** to see your data transformed
5. **Download the result** for use in your LLM application

---

**Happy converting! 🎯**
