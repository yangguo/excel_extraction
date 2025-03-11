#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Conclusion Generator

This script extracts and compares design and operation data from Excel files containing IT control tests.
It processes data from multiple sheets and generates a comparison report to identify differences
between design specifications and operational implementation.
"""

import pandas as pd
import os
import numpy as np
from pathlib import Path


def find_excel_files(directory):
    """
    Find all Excel files in the specified directory.
    
    Args:
        directory (Path): Directory path to search for Excel files
        
    Returns:
        list: List of Path objects pointing to Excel files
    """
    return list(directory.glob("*.xlsx"))


def extract_first_sheet_data(df, max_row=47):
    """
    Extract relevant data from the first sheet of an Excel file.
    
    Args:
        df (DataFrame): DataFrame containing the first sheet data
        max_row (int): Maximum row index to extract (default: 47)
        
    Returns:
        list: List of dictionaries containing extracted data
    """
    print("\nExtracting data from first sheet...")
    first_sheet_data = []
    
    if len(df) < 12:
        print("First sheet does not have enough rows to extract data from row 12.")
        return first_sheet_data
    
    # Ensure we don't go beyond the available rows
    max_row = min(max_row, len(df) - 1)
    
    # Extract the data for rows 12 to 48 (0-indexed: 11 to 47)
    for row in range(11, max_row + 1):
        row_data = {
            'No': str(df.iloc[row, 3]) if pd.notna(df.iloc[row, 3]) else '',
            'Description': str(df.iloc[row, 4]) if pd.notna(df.iloc[row, 4]) else '',
            'HDesign': str(df.iloc[row, 7]) if pd.notna(df.iloc[row, 7]) else '',
            'HOperation': str(df.iloc[row, 8]) if pd.notna(df.iloc[row, 9]) else ''
        }
        first_sheet_data.append(row_data)
    
    return first_sheet_data


def process_sheet(df, sheet_name):
    """
    Process a single sheet from an Excel file to extract conclusion data.
    
    Args:
        df (DataFrame): DataFrame containing the sheet data
        sheet_name (str): Name of the sheet being processed
        
    Returns:
        dict: Dictionary containing sheet conclusion data
    """
    print(f"\nProcessing sheet: {sheet_name}")
    
    # Find header values based on sheet type
    if sheet_name == 'PE-6':
        d8_value = df.iloc[9, 3] if len(df) > 9 else None
        start_row = 16
    elif sheet_name in ['PE-3d', 'PE-8']:
        d8_value = df.iloc[7, 3] if len(df) > 7 else None
        start_row = 13
    else:
        d8_value = df.iloc[6, 3] if len(df) > 6 else None
        start_row = 13
    
    # Find the conclusion row by scanning for non-numeric values
    current_row = start_row
    design_value = str(d8_value) if pd.notna(d8_value) else ''
    operation_value = ''
    
    while current_row < len(df):
        b_value = df.iloc[current_row, 1]  # Column B
        print(f"Current B value: {b_value}")
        
        # Check if the value is numeric
        if not pd.isna(b_value) and isinstance(b_value, (int, float)) and b_value == int(b_value):
            current_row += 1
        else:
            # Check the next row before breaking
            if current_row + 1 < len(df):
                next_b_value = df.iloc[current_row + 1, 1]  # Next row's Column B
                print(f"Next B value: {next_b_value}")
                if not pd.isna(next_b_value) and isinstance(next_b_value, (int, float)) and next_b_value == int(next_b_value):
                    # Next row has valid numeric value, continue with this row
                    current_row += 1
                    continue
            
            # Both current and next rows don't have valid numeric values
            # Update conclusion values
            d_value = df.iloc[current_row, 3]  # Column D
            operation_value = str(d_value) if pd.notna(d_value) else ''
            break
    
    print(f"Sheet: {sheet_name} - Extracted conclusion data")
    
    return {
        'Sheet': sheet_name,
        'Type': 'Detail',
        'Design': design_value,
        'Operation': operation_value
    }


def process_excel_file(input_file):
    """
    Process an Excel file to extract conclusion data.
    
    Args:
        input_file (Path): Path to the Excel file to process
        
    Returns:
        tuple: (first_sheet_data, conclusion_data, output_file)
    """
    output_file = input_file.name.replace(".xlsx", "_conclusion.xlsx")
    
    # Skip files that are already processed or are output files
    if (input_file.name == output_file or 
        "_conclusion.xlsx" in input_file.name or 
        "_extracted.xlsx" in input_file.name):
        return None, None, None
    
    print(f"\nProcessing file: {input_file.name}")
    first_sheet_data = []
    conclusion_data = []
    
    # Extract data from the first sheet
    try:
        first_sheet_df = pd.read_excel(input_file, sheet_name=0)
        first_sheet_data = extract_first_sheet_data(first_sheet_df)
    except Exception as e:
        print(f"Error extracting data from first sheet: {str(e)}")
    
    # Process the rest of the sheets
    try:
        xl = pd.ExcelFile(input_file)
        sheet_names = xl.sheet_names[1:]  # Skip the first sheet
    except Exception as e:
        print(f"Error reading file {input_file.name}: {str(e)}")
        return first_sheet_data, conclusion_data, output_file
    
    # Process each sheet
    for sheet_name in sheet_names:
        try:
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            sheet_data = process_sheet(df, sheet_name)
            conclusion_data.append(sheet_data)
            
            # Add additional blank rows for specific sheets
            if sheet_name in ['PE-3d', 'PE-8']:
                conclusion_data.append({
                    'Sheet': sheet_name,
                    'Type': 'Detail',
                    'Design': '',
                    'Operation': ''
                })
            elif sheet_name == 'PE-6':
                # Add three blank rows for PE-6
                for _ in range(3):
                    conclusion_data.append({
                        'Sheet': sheet_name,
                        'Type': 'Detail',
                        'Design': '',
                        'Operation': ''
                    })
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {str(e)}")
    
    return first_sheet_data, conclusion_data, output_file


def analyze_and_save_conclusion(first_sheet_data, conclusion_data, output_file):
    """
    Analyze and save the conclusion data to an Excel file.
    
    Args:
        first_sheet_data (list): Data from the first sheet
        conclusion_data (list): Conclusion data from other sheets
        output_file (str): Path to save the output file
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not first_sheet_data or not conclusion_data:
        print("No data was extracted from sheets")
        return False
    
    # Create DataFrames from extracted data
    first_df = pd.DataFrame(first_sheet_data)
    conclusion_df = pd.DataFrame(conclusion_data)
    
    # Combine first sheet data with conclusion data
    df_combined = pd.merge(first_df, conclusion_df, how='outer', left_index=True, right_index=True)
    
    # Fill empty values with values from previous rows
    df_combined.replace("", np.nan, inplace=True)
    df_combined.fillna(method='ffill', inplace=True)
    
    # Compare the difference between design specifications and actual implementation
    df_combined['Design Difference'] = np.where(df_combined['HDesign'] != df_combined['Design'], 'N', 'Y')
    df_combined['Operation Difference'] = np.where(df_combined['HOperation'] != df_combined['Operation'], 'N', 'Y')
    
    # Print summary statistics
    print("\nDesign Differences:\n", df_combined['Design Difference'].value_counts())
    print("Operation Differences:\n", df_combined['Operation Difference'].value_counts())
    
    # Save to Excel file
    df_combined.to_excel(output_file, sheet_name='Combined Data', index=False)
    print(f"\nData extracted successfully to {output_file} with {len(df_combined)} total rows")
    
    return True


def main():
    """
    Main function to extract and analyze conclusion data from Excel files.
    """
    # Get the current directory
    current_dir = Path.cwd()
    
    # Find all Excel files in the current directory
    excel_files = find_excel_files(current_dir)
    
    total_files_processed = 0
    
    # Process each Excel file
    for input_file in excel_files:
        first_sheet_data, conclusion_data, output_file = process_excel_file(input_file)
        
        if not output_file:
            continue
        
        if analyze_and_save_conclusion(first_sheet_data, conclusion_data, output_file):
            total_files_processed += 1
    
    print(f"\nSummary: Processed {total_files_processed} files successfully.")


if __name__ == "__main__":
    main()