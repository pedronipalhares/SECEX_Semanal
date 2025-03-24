# SECEX Weekly Data Processing

This repository contains a Python script for processing weekly export and import data from SECEX (Secretaria de Comércio Exterior). The script handles both mid-month updates and new month data, calculating appropriate deltas and maintaining historical records.

## Features

- Dynamic data extraction from Excel files
- Automatic calculation of weekly deltas
- Protection against duplicate entries
- Backup system for historical data
- Separate processing for exports and imports
- Support for both mid-month updates and new month data

## Directory Structure

```
SECEX_Semanal/
├── datasets/
│   ├── current/                 # Current week's processed data
│   ├── downloads/              # Downloaded Excel files
│   └── historicals/            # Historical data
│       ├── backups/           # Automatic backups of historical data
│       ├── Brazil_Secex_Weekly_Exports.csv
│       └── Brazil_Secex_Weekly_Imports.csv
├── format_exports.py           # Main processing script
├── .gitignore
└── README.md
```

## How to Use

1. Run the script:
   ```bash
   python3 format_exports.py
   ```

2. The script will:
   - Create backups of existing historical data
   - Download the latest Excel file
   - Extract week number and working days
   - Process export and import data
   - Update historical records
   - Save current week's data

## Data Processing Logic

### Mid-Month Updates
- Calculates the difference between current and previous values
- Considers working days delta
- Updates daily averages based on the delta

### New Month Data
- Uses raw values for new months
- Calculates appropriate daily averages
- Maintains data consistency

## Duplicate Protection

The script includes protection against accidental duplicate entries:
- Checks if data for the current week already exists
- Creates backups before any modifications
- Warns users about existing entries
- Prevents unintended data corruption

## File Handling

- Historical data is preserved in CSV format
- Automatic backups are created with timestamps
- Downloaded Excel files are saved with timestamps
- Current week's data is saved separately

## Requirements

- Python 3.x
- Required Python packages:
  - pandas
  - numpy
  - requests
  - openpyxl

## Notes

- The script automatically handles both export and import data
- Historical files are tracked in version control
- Temporary and backup files are ignored by git
- Excel files are not committed to the repository 