# 🚀 Quick Start Guide

Get up and running with the Excel-to-LLM Context Feeder Tool in under 5 minutes!

## ⚡ Super Quick Start

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Generate sample data:**
   ```bash
   python sample_data.py
   ```

3. **Launch the GUI:**
   ```bash
   streamlit run src/app.py
   ```

4. **Open your browser** to `http://localhost:8501`

5. **Upload the sample Excel file** from `sample_data/sample_company_data.xlsx`

6. **Click Convert** and see the magic happen! ✨

## 🎯 What You'll See

- **Beautiful Streamlit interface** with sidebar configuration
- **File upload** for Excel files (.xlsx, .xls)
- **Multiple output formats**: CSV, JSON, Markdown
- **Sheet selection** for multi-sheet Excel files
- **Data preview** and statistics
- **Download options** for your converted data

## 🔧 Alternative: Command Line

Prefer the command line? No problem!

```bash
# Basic conversion
python src/cli.py sample_data/sample_company_data.xlsx -f csv

# With summary and custom output
python src/cli.py sample_data/sample_company_data.xlsx -f json --summary -o my_output.txt

# Process specific sheets
python src/cli.py sample_data/sample_company_data.xlsx -s "Employee_Data,Financial_Data" -f markdown

# Enable chunking for large files
python src/cli.py sample_data/sample_company_data.xlsx -f csv --chunk --chunk-size 3000
```

## 🧪 Test the Tool

Run the demo to see all features in action:

```bash
python demo.py
```

This will:
- Generate sample Excel data
- Convert to different formats
- Show chunking capabilities
- Demonstrate custom prompts

## 📊 Sample Data Included

The tool comes with sample company data including:
- **Employee Data**: 100 employees with departments, salaries, performance ratings
- **Financial Data**: 24 months of revenue, expenses, and metrics
- **Department Summary**: Aggregated statistics by department
- **City Analysis**: Employee distribution and performance by city

## 🎨 Customization Options

- **Output Formats**: CSV, JSON, Markdown
- **Prompt Templates**: Customize how your data is presented
- **Chunking**: Split large files into manageable pieces
- **Data Summaries**: Include statistical information
- **Sheet Selection**: Choose which sheets to process

## 🚨 Troubleshooting

**"Module not found" errors?**
```bash
pip install -r requirements.txt
```

**Streamlit not working?**
```bash
pip install streamlit --upgrade
```

**Excel file issues?**
- Ensure file is not corrupted
- Check file extension (.xlsx or .xls)
- Verify file is not password protected

**Permission errors?**
- Check file permissions
- Ensure output directory is writable

## 🔗 Next Steps

1. **Try your own Excel files**
2. **Explore different output formats**
3. **Customize prompt templates**
4. **Test with large datasets**
5. **Integrate with your LLM workflows**

## 💡 Pro Tips

- **Large files**: Use chunking for files >10,000 rows
- **Multiple sheets**: Process related data together for better context
- **Custom prompts**: Tailor output for your specific LLM use case
- **Data summaries**: Include summaries for better LLM understanding

---

**Need help?** Check the main README.md for detailed documentation and examples.

**Happy converting! 🎯**
