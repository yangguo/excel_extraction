# Excel Data Extractor

A Python utility for extracting and analyzing data from IT control testing Excel files.

## Overview

This project provides tools to extract and process specific data from Excel files containing IT general control test results. It's designed to:

1. Extract detailed data from multiple sheets of Excel files
2. Process and organize the extracted information 
3. Compare and analyze differences between design and operation controls

## Features

- **Data Extraction**: Extract data from specific cells and ranges in Excel files
- **Batch Processing**: Process multiple Excel files in a directory
- **Data Comparison**: Compare design specifications with actual operations
- **Automatic Output**: Generate formatted Excel files with extracted data

## Installation

### Prerequisites

- Python 3.10+
- Virtual environment (recommended)

### Setup

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/excel-data-extractor.git
   cd excel-data-extractor
   ```

2. Create and activate a virtual environment:
   ```
   python -m venv myenv
   # On Windows
   myenv\Scripts\activate
   # On Unix or MacOS
   source myenv/bin/activate
   ```

3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

### Extract Data

To extract detailed data from Excel files:

```python
python extract_data.py
```

This will process all Excel files in the current directory and create output files with "_extracted" suffix.

### Generate Conclusions

To compare and analyze data from Excel files:

```python
python extract_conclusion.py
```

This will process all Excel files, compare design specifications with operations, and create output files with "_conclusion" suffix.

## Project Structure

```
excel-data-extractor/
├── README.md                 # Project documentation
├── extract_data.py           # Data extraction script
├── extract_conclusion.py     # Analysis and comparison script
├── requirements.txt          # Package dependencies
├── .gitignore                # Git ignore file
└── myenv/                    # Virtual environment (not tracked in git)
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.