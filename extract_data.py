#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Data Extraction Tool

This script extracts structured data from Excel files containing IT control testing results.
It processes each sheet in the Excel files and extracts header information and detailed data rows,
organizing them into a standardized format and saving to a new Excel file.
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


def combine_row_values(df, row, start_col):
    """
    Combine values from multiple columns starting from start_col until a NaN value is encountered.
    
    Args:
        df (DataFrame): DataFrame containing the data
        row (int): Row index to process
        start_col (int): Starting column index
        
    Returns:
        str: Combined values as a string
    """
    combined_value = ""
    col = start_col
    num_cols = len(df.columns)
    
    while col < num_cols:
        cell_value = df.iloc[row, col]
        if pd.isna(cell_value):
            break
        
        # Add a newline between values if not the first value
        if combined_value:
            combined_value += "\n"
            
        combined_value += str(cell_value)
        col += 1
    
    return combined_value


def process_sheet(df, sheet_name):
    """
    Process a single sheet from an Excel file to extract data.
    
    Args:
        df (DataFrame): DataFrame containing the sheet data
        sheet_name (str): Name of the sheet being processed
        
    Returns:
        tuple: (header_data, detail_rows, rows_processed)
            - header_data: Dictionary containing header information
            - detail_rows: List of dictionaries with detail row data
            - rows_processed: Number of detail rows processed
    """
    print(f"\nProcessing sheet: {sheet_name}")
    detail_rows = []
    
    # remove spaces from sheet name
    sheet_name = sheet_name.replace(" ", "")

    # Find header values (B6, D6, B8, D8) based on sheet type
    if sheet_name == 'PE-6':
        b6_value = df.iloc[7, 1] if len(df) > 7 else None
        d6_value = df.iloc[7, 3] if len(df) > 7 else None
        b8_value = df.iloc[9, 1] if len(df) > 9 else None
        d8_value = df.iloc[9, 3] if len(df) > 9 else None
        start_row = 16
    elif sheet_name in ['PE-3d', 'PE-8']:
        b6_value = df.iloc[5, 1] if len(df) > 5 else None
        d6_value = df.iloc[5, 3] if len(df) > 5 else None
        b8_value = df.iloc[7, 1] if len(df) > 7 else None
        d8_value = df.iloc[7, 3] if len(df) > 7 else None
        start_row = 13
    else:
        b6_value = df.iloc[4, 1] if len(df) > 4 else None
        d6_value = df.iloc[4, 3] if len(df) > 4 else None
        b8_value = df.iloc[6, 1] if len(df) > 6 else None
        d8_value = df.iloc[6, 3] if len(df) > 6 else None
        start_row = 13
    
    # Create header data dictionary
    header_data = {
        'Sheet': sheet_name,
        'Type': 'Header',
        'Number': 0,
        'Description': str(b6_value) if pd.notna(b6_value) else '',
        'Details': str(d6_value) if pd.notna(d6_value) else '',
        'Control': str(b8_value) if pd.notna(b8_value) else '',
        'Conclusion': str(d8_value) if pd.notna(d8_value) else ''
    }
    
    # Process detail rows
    current_row = start_row
    while current_row < len(df):
        b_value = df.iloc[current_row, 1]  # Column B
        print(f"Current B value: {b_value}")
        
        # Check if the value is numeric (a row number)
        if not pd.isna(b_value) and isinstance(b_value, (int, float)) and b_value == int(b_value):
            c_value = df.iloc[current_row, 2]  # Column C
            
            # Combine values from column D onwards until NaN is encountered
            d_value = combine_row_values(df, current_row, 3)  # Start at column D (index 3)
            
            detail_rows.append({
                'Sheet': sheet_name,
                'Type': 'Detail',
                'Number': int(b_value),
                'Description': str(c_value) if pd.notna(c_value) else '',
                'Details': d_value if d_value else '',
                'Control': '',
                'Conclusion': ''
            })
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
            # Update control and conclusion values for the last detail row
            b_value = df.iloc[current_row, 1]
            d_value = df.iloc[current_row, 3]
            
            if detail_rows:  # Only update if we have detail rows
                detail_rows[-1]['Control'] = str(b_value) if pd.notna(b_value) else ''
                detail_rows[-1]['Conclusion'] = str(d_value) if pd.notna(d_value) else ''
            break
    
    rows_processed = current_row - start_row
    print(f"Sheet: {sheet_name} - Extracted header and {rows_processed} detail rows")
    return header_data, detail_rows, rows_processed


def process_excel_file(input_file):
    """
    Process an Excel file to extract data from all sheets.
    
    Args:
        input_file (Path): Path to the Excel file to process
        
    Returns:
        list: List of dictionaries containing extracted data
    """
    output_file = input_file.name.replace(".xlsx", "_extracted.xlsx")
    # Skip if output file matches input file
    if input_file.name == output_file:
        return []
    
    print(f"\nProcessing file: {input_file.name}")
    all_data = []
    
    try:
        xl = pd.ExcelFile(input_file)
        sheet_names = xl.sheet_names[1:]  # Skip the first sheet
    except Exception as e:
        print(f"Error reading file {input_file.name}: {str(e)}")
        return []
    
    # Process each sheet
    for sheet_name in sheet_names:
        try:
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            header_data, detail_rows, _ = process_sheet(df, sheet_name)
            
            # Add header data and detail rows to results
            all_data.append(header_data)
            all_data.extend(detail_rows)
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {str(e)}")
    
    return all_data, output_file


def post_process_data(all_data):
    """
    Post-process extracted data to fill empty values and ensure consistency.
    
    Args:
        all_data (list): List of dictionaries containing extracted data
        
    Returns:
        list: Processed data with filled values
    """
    if not all_data:
        return []
    
    # Fill "details" column with upper cell value if empty
    for i in range(len(all_data)):
        if all_data[i]['Details'] == '' and i > 0:
            all_data[i]['Details'] = all_data[i-1]['Details']
    
    # Fill "control" and "conclusion" columns with down cell value if empty (loop backwards)
    for i in range(len(all_data)-1, 0, -1):
        if all_data[i]['Control'] == '' and i < len(all_data)-1:
            all_data[i]['Control'] = all_data[i+1]['Control']
        if all_data[i]['Conclusion'] == '' and i < len(all_data)-1:
            all_data[i]['Conclusion'] = all_data[i+1]['Conclusion']
    
    return all_data


def main():
    """
    Main function to extract data from Excel files in the current directory.
    """
    # Get the current directory
    current_dir = Path.cwd()
    
    # Find all Excel files in the current directory
    excel_files = find_excel_files(current_dir)
    
    total_files_processed = 0
    total_rows_extracted = 0
    
    # Process each Excel file
    for input_file in excel_files:
        result = process_excel_file(input_file)
        
        if not result:
            continue
            
        all_data, output_file = result
        
        if all_data:
            # Post-process data to fill empty values
            processed_data = post_process_data(all_data)
            
            # Create DataFrame and save to Excel
            df_combined = pd.DataFrame(processed_data)
            df_combined.to_excel(output_file, sheet_name='Combined Data', index=False)
            
            total_files_processed += 1
            total_rows_extracted += len(processed_data)
            
            print(f"\nData extracted successfully to {output_file} with {len(processed_data)} total rows")
        else:
            print("No data was extracted from any sheet")
    
    print(f"\nSummary: Processed {total_files_processed} files, extracted {total_rows_extracted} total rows")


if __name__ == "__main__":
    main()