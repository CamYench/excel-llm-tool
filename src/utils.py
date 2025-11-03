"""
Utility functions for Excel-to-LLM Context Feeder Tool.
"""

import pandas as pd
from typing import List, Dict, Any, Union
import json


def chunk_df(df: pd.DataFrame, chunk_size: int = 5000) -> List[pd.DataFrame]:
    """
    Split a DataFrame into chunks of specified size.
    
    Args:
        df: Input DataFrame
        chunk_size: Number of rows per chunk
        
    Returns:
        List of DataFrame chunks
    """
    if len(df) <= chunk_size:
        return [df]
    
    chunks = []
    for i in range(0, len(df), chunk_size):
        chunk = df.iloc[i:i + chunk_size]
        chunks.append(chunk)
    
    return chunks


def format_prompt(data_text: str, prompt_template: str = "Analyze this data: [formatted_data]") -> str:
    """
    Format data with a custom prompt template.
    
    Args:
        data_text: Formatted data string
        prompt_template: Template with [formatted_data] placeholder
        
    Returns:
        Formatted prompt with data
    """
    return prompt_template.replace("[formatted_data]", data_text)


def detect_sheets(file_path: Union[str, Any]) -> List[str]:
    """
    Detect available sheets in an Excel file.
    
    Args:
        file_path: Path to Excel file or file-like object
        
    Returns:
        List of sheet names
    """
    try:
        xls = pd.ExcelFile(file_path)
        return xls.sheet_names
    except Exception as e:
        raise ValueError(f"Could not read Excel file: {str(e)}")


def get_data_summary(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Generate a summary of DataFrame statistics.
    
    Args:
        df: Input DataFrame
        
    Returns:
        Dictionary with summary statistics
    """
    summary = {
        "rows": len(df),
        "columns": len(df.columns),
        "column_types": df.dtypes.to_dict(),
        "missing_values": df.isnull().sum().to_dict(),
        "memory_usage": df.memory_usage(deep=True).sum()
    }
    
    # Add numeric column statistics
    numeric_cols = df.select_dtypes(include=['number']).columns
    if len(numeric_cols) > 0:
        summary["numeric_summary"] = df[numeric_cols].describe().to_dict()
    
    return summary


def format_summary(summary: Dict[str, Any]) -> str:
    """
    Format summary statistics as readable text.
    
    Args:
        summary: Summary dictionary
        
    Returns:
        Formatted summary string
    """
    lines = [
        "=== DATA SUMMARY ===",
        f"Rows: {summary['rows']:,}",
        f"Columns: {summary['columns']}",
        f"Memory Usage: {summary['memory_usage'] / 1024:.2f} KB",
        "",
        "Column Types:",
    ]
    
    for col, dtype in summary['column_types'].items():
        lines.append(f"  {col}: {dtype}")
    
    lines.append("")
    lines.append("Missing Values:")
    for col, missing in summary['missing_values'].items():
        if missing > 0:
            lines.append(f"  {col}: {missing}")
    
    if 'numeric_summary' in summary:
        lines.append("")
        lines.append("Numeric Column Statistics:")
        for col, stats in summary['numeric_summary'].items():
            lines.append(f"  {col}:")
            for stat, value in stats.items():
                if pd.notna(value):
                    lines.append(f"    {stat}: {value:.4f}")
    
    return "\n".join(lines)
