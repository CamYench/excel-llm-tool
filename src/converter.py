"""
Core conversion logic for Excel-to-LLM Context Feeder Tool.
"""

import pandas as pd
from typing import List, Union, Any
import json
from utils import chunk_df, format_prompt, get_data_summary, format_summary


def excel_to_formatted_text(
    file_path: Union[str, Any],
    sheet_names: Union[str, List[str]] = None,
    output_format: str = "csv",
    chunk_size: int = 5000,
    prompt_template: str = "Analyze this data: [formatted_data]",
    include_summary: bool = False,
    evaluate_formulas: bool = False
) -> str:
    """
    Convert Excel file to formatted text for LLM context.
    
    Args:
        file_path: Path to Excel file or file-like object
        sheet_names: Sheet name(s) to process (None for all)
        output_format: Output format ('csv', 'json', 'markdown')
        chunk_size: Number of rows per chunk for large files
        prompt_template: Template for prompt formatting
        include_summary: Whether to include data summary
        evaluate_formulas: Whether to evaluate Excel formulas
        
    Returns:
        Formatted text string
    """
    try:
        # Read Excel file
        if sheet_names is None:
            # Read all sheets
            all_sheets = pd.read_excel(
                file_path, 
                sheet_name=None, 
                engine='openpyxl',
                na_filter=False
            )
        elif isinstance(sheet_names, str):
            # Single sheet
            all_sheets = {sheet_names: pd.read_excel(
                file_path, 
                sheet_name=sheet_names, 
                engine='openpyxl',
                na_filter=False
            )}
        else:
            # Multiple specific sheets
            all_sheets = {}
            for sheet in sheet_names:
                try:
                    all_sheets[sheet] = pd.read_excel(
                        file_path, 
                        sheet_name=sheet, 
                        engine='openpyxl',
                        na_filter=False
                    )
                except Exception as e:
                    print(f"Warning: Could not read sheet '{sheet}': {e}")
        
        if not all_sheets:
            raise ValueError("No sheets could be read from the Excel file")
        
        # Process each sheet
        all_formatted_data = []
        
        for sheet_name, df in all_sheets.items():
            if df.empty:
                continue
                
            # Clean the data
            df = df.fillna("")
            
            # Convert to specified format
            if output_format.lower() == "csv":
                formatted_data = df.to_csv(index=False, sep=',')
            elif output_format.lower() == "json":
                # Convert to records format for better LLM consumption
                formatted_data = df.to_json(orient='records', indent=2, force_ascii=False)
            elif output_format.lower() == "markdown":
                formatted_data = df.to_markdown(index=False)
            else:
                raise ValueError(f"Unsupported output format: {output_format}")
            
            # Add sheet header
            sheet_header = f"\n=== SHEET: {sheet_name} ===\n"
            all_formatted_data.append(sheet_header + formatted_data)
        
        # Combine all sheet data
        combined_data = "\n\n".join(all_formatted_data)
        
        # Add summary if requested
        if include_summary:
            # Use first non-empty sheet for summary
            first_df = next((df for df in all_sheets.values() if not df.empty), None)
            if first_df is not None:
                summary = get_data_summary(first_df)
                summary_text = format_summary(summary)
                combined_data = summary_text + "\n\n" + combined_data
        
        # Apply prompt template
        final_output = format_prompt(combined_data, prompt_template)
        
        return final_output
        
    except Exception as e:
        raise Exception(f"Error converting Excel file: {str(e)}")


def excel_to_chunked_text(
    file_path: Union[str, Any],
    sheet_names: Union[str, List[str]] = None,
    output_format: str = "csv",
    chunk_size: int = 5000,
    prompt_template: str = "Analyze this data: [formatted_data]",
    include_summary: bool = False
) -> List[str]:
    """
    Convert Excel file to chunked text for large datasets.
    
    Args:
        file_path: Path to Excel file or file-like object
        sheet_names: Sheet name(s) to process
        output_format: Output format
        chunk_size: Number of rows per chunk
        prompt_template: Template for prompt formatting
        include_summary: Whether to include data summary
        
    Returns:
        List of chunked text strings
    """
    try:
        # Read Excel file
        if sheet_names is None:
            all_sheets = pd.read_excel(
                file_path, 
                sheet_name=None, 
                engine='openpyxl',
                na_filter=False
            )
        elif isinstance(sheet_names, str):
            all_sheets = {sheet_names: pd.read_excel(
                file_path, 
                sheet_name=sheet_names, 
                engine='openpyxl',
                na_filter=False
            )}
        else:
            all_sheets = {}
            for sheet in sheet_names:
                try:
                    all_sheets[sheet] = pd.read_excel(
                        file_path, 
                        sheet_name=sheet, 
                        engine='openpyxl',
                        na_filter=False
                    )
                except Exception as e:
                    print(f"Warning: Could not read sheet '{sheet}': {e}")
        
        if not all_sheets:
            raise ValueError("No sheets could be read from the Excel file")
        
        chunks = []
        
        for sheet_name, df in all_sheets.items():
            if df.empty:
                continue
                
            # Clean the data
            df = df.fillna("")
            
            # Split into chunks
            df_chunks = chunk_df(df, chunk_size)
            
            for i, chunk_df in enumerate(df_chunks):
                # Convert chunk to specified format
                if output_format.lower() == "csv":
                    formatted_chunk = chunk_df.to_csv(index=False, sep=',')
                elif output_format.lower() == "json":
                    formatted_chunk = chunk_df.to_json(orient='records', indent=2, force_ascii=False)
                elif output_format.lower() == "markdown":
                    formatted_chunk = chunk_df.to_markdown(index=False)
                else:
                    raise ValueError(f"Unsupported output format: {output_format}")
                
                # Add chunk header
                chunk_header = f"\n=== SHEET: {sheet_name} - CHUNK {i+1}/{len(df_chunks)} ===\n"
                chunk_text = chunk_header + formatted_chunk
                
                # Apply prompt template
                final_chunk = format_prompt(chunk_text, prompt_template)
                chunks.append(final_chunk)
        
        # Add summary if requested
        if include_summary and chunks:
            first_df = next((df for df in all_sheets.values() if not df.empty), None)
            if first_df is not None:
                summary = get_data_summary(first_df)
                summary_text = format_summary(summary)
                summary_chunk = format_prompt(summary_text, prompt_template)
                chunks.insert(0, summary_chunk)
        
        return chunks
        
    except Exception as e:
        raise Exception(f"Error converting Excel file to chunks: {str(e)}")


def validate_excel_file(file_path: Union[str, Any]) -> bool:
    """
    Validate that the file is a valid Excel file.
    
    Args:
        file_path: Path to file or file-like object
        
    Returns:
        True if valid Excel file, False otherwise
    """
    try:
        pd.read_excel(file_path, nrows=1, engine='openpyxl')
        return True
    except Exception:
        return False
