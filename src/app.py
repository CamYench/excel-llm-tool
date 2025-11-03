"""
Streamlit GUI for Excel-to-LLM Context Feeder Tool.
"""

import streamlit as st
import pandas as pd
import os
from pathlib import Path
from converter import excel_to_formatted_text, excel_to_chunked_text, validate_excel_file
from utils import detect_sheets, get_data_summary, format_summary

# Page configuration
st.set_page_config(
    page_title="Excel-to-LLM Context Feeder",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sidebar-header {
        font-size: 1.2rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
    }
    .info-box {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
        margin: 1rem 0;
        color: #2c3e50;
    }
    .info-box strong {
        color: #1f77b4;
    }
    .success-box {
        background-color: #d4edda;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #dc3545;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    """Main application function."""
    
    # Header
    st.markdown('<h1 class="main-header">📊 Excel-to-LLM Context Feeder</h1>', unsafe_allow_html=True)
    
    # Sidebar for configuration
    with st.sidebar:
        st.markdown('<h3 class="sidebar-header">⚙️ Configuration</h3>', unsafe_allow_html=True)
        
        # Output format selection
        output_format = st.selectbox(
            "Output Format",
            ["CSV", "JSON", "Markdown"],
            help="Choose the format for the converted data"
        )
        
        # Chunk size for large files
        chunk_size = st.slider(
            "Chunk Size (rows)",
            min_value=1000,
            max_value=10000,
            value=5000,
            step=1000,
            help="Number of rows per chunk for large files"
        )
        
        # Include summary option
        include_summary = st.checkbox(
            "Include Data Summary",
            value=True,
            help="Add statistical summary of the data"
        )
        
        # Advanced options
        with st.expander("Advanced Options"):
            # Prompt template
            prompt_template = st.text_area(
                "Prompt Template",
                value="Analyze this data: [formatted_data]",
                help="Template for formatting the output. Use [formatted_data] as placeholder.",
                height=100
            )
            
            # Evaluate formulas option
            evaluate_formulas = st.checkbox(
                "Evaluate Excel Formulas",
                value=False,
                help="Attempt to evaluate Excel formulas (experimental)"
            )
            
            # Chunking option
            use_chunking = st.checkbox(
                "Enable Chunking",
                value=False,
                help="Split large files into multiple chunks"
            )
        
        # Information box
        st.markdown("""
        <div class="info-box">
            <strong>💡 Tips:</strong><br>
            • Upload Excel files (.xlsx, .xls)<br>
            • Select specific sheets or process all<br>
            • Choose output format for your LLM<br>
            • Use chunking for very large files
        </div>
        """, unsafe_allow_html=True)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 📁 File Upload")
        
        # File uploader
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=["xlsx", "xls"],
            help="Upload an Excel file to convert"
        )
        
        if uploaded_file is not None:
            # Validate file
            if not validate_excel_file(uploaded_file):
                st.error("❌ Invalid Excel file. Please upload a valid .xlsx or .xls file.")
                return
            
            # Reset file pointer for reading
            uploaded_file.seek(0)
            
            try:
                # Detect available sheets
                sheets = detect_sheets(uploaded_file)
                uploaded_file.seek(0)  # Reset again for processing
                
                st.success(f"✅ File uploaded successfully! Found {len(sheets)} sheet(s).")
                
                # Sheet selection
                if len(sheets) > 1:
                    selected_sheets = st.multiselect(
                        "Select Sheets to Process",
                        sheets,
                        default=sheets[0],
                        help="Choose which sheets to include in the conversion"
                    )
                else:
                    selected_sheets = sheets
                    st.info(f"📋 Processing sheet: {sheets[0]}")
                
                # Output file path
                st.markdown("### 💾 Output Configuration")
                output_filename = st.text_input(
                    "Output Filename",
                    value=f"excel_context_{output_format.lower()}.txt",
                    help="Name for the output file"
                )
                
                # Convert button
                if st.button("🚀 Convert Excel to LLM Context", type="primary", use_container_width=True):
                    if not selected_sheets:
                        st.error("❌ Please select at least one sheet to process.")
                        return
                    
                    # Process the file
                    with st.spinner("🔄 Processing Excel file..."):
                        try:
                            if use_chunking:
                                # Use chunked processing
                                chunks = excel_to_chunked_text(
                                    uploaded_file,
                                    selected_sheets,
                                    output_format.lower(),
                                    chunk_size,
                                    prompt_template,
                                    include_summary
                                )
                                
                                # Combine chunks for display
                                result = "\n\n".join(chunks)
                                st.success(f"✅ File processed successfully! Generated {len(chunks)} chunks.")
                            else:
                                # Use single processing
                                result = excel_to_formatted_text(
                                    uploaded_file,
                                    selected_sheets,
                                    output_format.lower(),
                                    chunk_size,
                                    prompt_template,
                                    include_summary,
                                    evaluate_formulas
                                )
                                st.success("✅ File processed successfully!")
                            
                            # Save to file
                            try:
                                with open(output_filename, "w", encoding="utf-8") as f:
                                    f.write(result)
                                st.success(f"💾 Saved to: {output_filename}")
                            except Exception as e:
                                st.error(f"❌ Error saving file: {str(e)}")
                            
                            # Display results
                            st.markdown("### 📊 Conversion Results")
                            
                            # File info
                            file_size = len(result.encode('utf-8')) / 1024  # KB
                            st.info(f"📏 Output size: {file_size:.2f} KB")
                            
                            # Preview
                            with st.expander("👀 Preview Output", expanded=True):
                                # Truncate for display if too long
                                preview_length = 3000
                                if len(result) > preview_length:
                                    st.text(result[:preview_length] + "\n\n... (truncated for display)")
                                    st.info(f"📝 Full output is {len(result):,} characters long")
                                else:
                                    st.text(result)
                            
                            # Download button
                            st.download_button(
                                label="⬇️ Download Output",
                                data=result,
                                file_name=output_filename,
                                mime="text/plain",
                                use_container_width=True
                            )
                            
                        except Exception as e:
                            st.error(f"❌ Error during conversion: {str(e)}")
                            st.exception(e)
                
            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)}")
                st.exception(e)
    
    with col2:
        st.markdown("### 📈 Data Preview")
        
        if uploaded_file is not None:
            try:
                # Reset file pointer
                uploaded_file.seek(0)
                
                # Show preview of first sheet
                sheets = detect_sheets(uploaded_file)
                if sheets:
                    uploaded_file.seek(0)
                    df = pd.read_excel(uploaded_file, sheet_name=sheets[0], nrows=10)
                    
                    st.markdown(f"**Preview of '{sheets[0]}' sheet:**")
                    st.dataframe(df, use_container_width=True)
                    
                    # Basic stats
                    uploaded_file.seek(0)
                    full_df = pd.read_excel(uploaded_file, sheet_name=sheets[0])
                    
                    st.markdown("**Quick Stats:**")
                    st.metric("Rows", f"{len(full_df):,}")
                    st.metric("Columns", len(full_df.columns))
                    
                    # Memory usage
                    memory_kb = full_df.memory_usage(deep=True).sum() / 1024
                    st.metric("Memory", f"{memory_kb:.1f} KB")
                    
            except Exception as e:
                st.error(f"Error previewing data: {str(e)}")
    
    # Footer
    st.markdown("---")
    st.markdown(
        "**Excel-to-LLM Context Feeder Tool** | "
        "Convert Excel files to LLM-friendly formats for better context and analysis."
    )

if __name__ == "__main__":
    main()
