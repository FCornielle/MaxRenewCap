"""
Input/Output Utilities Module

This module contains utility functions for file I/O operations,
CSV handling, and data management.
"""

import pandas as pd
import os
from pathlib import Path


def load_contingency_results(file_path='Resultados.csv', encoding='latin1'):
    """
    Load contingency analysis results from CSV file.
    
    Args:
        file_path (str): Path to the CSV file (default: 'Resultados.csv')
        encoding (str): File encoding (default: 'latin1')
        
    Returns:
        pandas.DataFrame: Loaded contingency results
        
    Example:
        >>> df = load_contingency_results('results.csv')
    """
    try:
        df = pd.read_csv(file_path, encoding=encoding, low_memory=False)
        print(f"Contingency results loaded from '{file_path}' with {len(df)} rows.")
        return df
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        return None
    except Exception as e:
        print(f"Error loading file '{file_path}': {e}")
        return None


def save_results_to_csv(df, file_path, index=False):
    """
    Save DataFrame to CSV file.
    
    Args:
        df (pandas.DataFrame): DataFrame to save
        file_path (str): Path where to save the CSV file
        index (bool): Whether to include row indices (default: False)
        
    Example:
        >>> save_results_to_csv(results_df, 'optimization_results.csv')
    """
    try:
        df.to_csv(file_path, index=index, encoding='utf-8')
        print(f"Results saved to '{file_path}' with {len(df)} rows.")
    except Exception as e:
        print(f"Error saving file '{file_path}': {e}")


def ensure_directory_exists(directory_path):
    """
    Ensure that a directory exists, create it if it doesn't.
    
    Args:
        directory_path (str): Path to the directory
        
    Returns:
        bool: True if directory exists or was created successfully
        
    Example:
        >>> ensure_directory_exists('output/results')
    """
    try:
        Path(directory_path).mkdir(parents=True, exist_ok=True)
        return True
    except Exception as e:
        print(f"Error creating directory '{directory_path}': {e}")
        return False


def get_safe_filename(filename):
    """
    Convert a string to a safe filename by removing/replacing invalid characters.
    
    Args:
        filename (str): Original filename
        
    Returns:
        str: Safe filename
        
    Example:
        >>> safe_name = get_safe_filename("Test/File:Name.csv")
    """
    # Replace invalid characters with underscores
    invalid_chars = '<>:"/\\|?*'
    safe_filename = filename
    for char in invalid_chars:
        safe_filename = safe_filename.replace(char, '_')
    
    return safe_filename


def create_results_summary(results_df, output_dir='output'):
    """
    Create a summary of optimization results and save to multiple formats.
    
    Args:
        results_df (pandas.DataFrame): Results DataFrame
        output_dir (str): Output directory (default: 'output')
        
    Returns:
        dict: Dictionary with file paths of created files
        
    Example:
        >>> files = create_results_summary(results_df, 'results')
    """
    if results_df is None or results_df.empty:
        print("No results to summarize.")
        return {}
    
    # Ensure output directory exists
    ensure_directory_exists(output_dir)
    
    files_created = {}
    
    try:
        # Save as CSV
        csv_path = os.path.join(output_dir, 'optimization_results.csv')
        save_results_to_csv(results_df, csv_path)
        files_created['csv'] = csv_path
        
        # Create summary statistics
        summary_stats = {
            'Total_Substations': len(results_df),
            'Average_Max_Power': results_df['Potencia Maxima'].mean(),
            'Max_Power_Overall': results_df['Potencia Maxima'].max(),
            'Min_Power_Overall': results_df['Potencia Maxima'].min(),
            'Average_Loading': results_df['Cargabilidad Maxima'].mean(),
            'Max_Loading': results_df['Cargabilidad Maxima'].max()
        }
        
        summary_df = pd.DataFrame(list(summary_stats.items()), columns=['Metric', 'Value'])
        summary_path = os.path.join(output_dir, 'summary_statistics.csv')
        save_results_to_csv(summary_df, summary_path)
        files_created['summary'] = summary_path
        
        # Create text report
        report_path = os.path.join(output_dir, 'optimization_report.txt')
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("PowerFactory Generator Optimization Results\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"Total Substations Analyzed: {len(results_df)}\n")
            f.write(f"Average Maximum Power: {summary_stats['Average_Max_Power']:.2f} MW\n")
            f.write(f"Maximum Power Achieved: {summary_stats['Max_Power_Overall']:.2f} MW\n")
            f.write(f"Minimum Power Achieved: {summary_stats['Min_Power_Overall']:.2f} MW\n")
            f.write(f"Average Loading: {summary_stats['Average_Loading']:.2f}%\n")
            f.write(f"Maximum Loading: {summary_stats['Max_Loading']:.2f}%\n\n")
            
            f.write("Detailed Results:\n")
            f.write("-" * 30 + "\n")
            f.write(results_df.to_string(index=False))
        
        files_created['report'] = report_path
        print(f"Results summary created in '{output_dir}' directory.")
        
    except Exception as e:
        print(f"Error creating results summary: {e}")
    
    return files_created
