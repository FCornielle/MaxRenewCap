#!/usr/bin/env python3
"""
Test script for CSV parser
"""

from src_code.csv_parser import parse_contingency_results

def main():
    print("Testing CSV parser...")
    result = parse_contingency_results()
    print(f"Found {len(result)} lines with loading data")
    
    if not result.empty:
        print("\nFirst 5 results:")
        print(result.head())
    else:
        print("No data found")

if __name__ == "__main__":
    main()
