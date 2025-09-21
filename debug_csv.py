#!/usr/bin/env python3
"""
Debug script to analyze the CSV file structure and identify issues
"""

import pandas as pd
import numpy as np

def debug_csv_structure():
    """Debug the CSV file structure to understand the data format"""
    print("🔍 Debugging CSV file structure...")
    
    try:
        # Load the CSV file
        df = pd.read_csv('notebooks/Resultados.csv', encoding='latin1', low_memory=False)
        print(f"✅ CSV loaded successfully: {df.shape[0]} rows, {df.shape[1]} columns")
        
        # Check the first few rows
        print("\n📊 First 3 rows:")
        print(df.head(3))
        
        # Check for "Max. loading in %" columns
        print("\n🔍 Looking for 'Max. loading in %' columns...")
        headers = df.iloc[0].values
        max_loading_indices = []
        
        for i, header in enumerate(headers):
            if "Max. loading in %" in str(header):
                max_loading_indices.append(i)
                print(f"  Found at column {i}: '{header}'")
        
        print(f"\n📈 Found {len(max_loading_indices)} 'Max. loading in %' columns")
        
        if len(max_loading_indices) > 0:
            # Check the line names row (row 1)
            line_names_row = df.iloc[1].values
            print("\n🔍 Line names in row 1:")
            for i, idx in enumerate(max_loading_indices[:5]):  # Show first 5
                line_name = line_names_row[idx]
                print(f"  Column {idx}: '{line_name}'")
            
            # Check some data values
            print("\n📊 Sample data values:")
            for i, idx in enumerate(max_loading_indices[:3]):  # Show first 3
                line_data = df.iloc[2:5, idx].values  # First 3 data rows
                print(f"  Column {idx} data: {line_data}")
        
        # Check for other loading-related columns
        print("\n🔍 Looking for other loading-related columns...")
        loading_columns = []
        for i, header in enumerate(headers):
            if "loading" in str(header).lower() or "cargabilidad" in str(header).lower():
                loading_columns.append((i, header))
                print(f"  Column {i}: '{header}'")
        
        return df, max_loading_indices
        
    except Exception as e:
        print(f"❌ Error loading CSV: {e}")
        return None, []

if __name__ == "__main__":
    df, max_loading_indices = debug_csv_structure()
    
    if df is not None and len(max_loading_indices) > 0:
        print(f"\n✅ Found {len(max_loading_indices)} loading columns")
        print("The CSV structure looks correct for processing.")
    else:
        print("\n❌ No loading columns found. This explains why the analysis is not working.")

