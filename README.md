# script-xlsx

A Python script that creates an interactive Excel workbook with dynamic filtering from multiple CSV data sources.

## Overview

`xlsx-create.py` processes healthcare provider data from multiple CSV files and combines them into a single Excel workbook with a powerful search interface. The script is designed to handle BCBS (Blue Cross Blue Shield) provider monitoring data across multiple databases.

## Features

- **Multi-source data consolidation**: Combines 7+ different data types into one workbook
- **Dynamic search interface**: Interactive Search_Tab with Excel FILTER formulas
- **Smart dropdown filtering**: Automatically generated issuer dropdown for license data
- **Automatic date stamping**: Sheet names include current date (e.g., `license_25-11-05`)
- **Scalable validation lists**: Handles large dropdown lists via hidden helper sheets

## Data Types Processed

The script processes the following provider data types:

| Type | CSV Prefix | Columns | Search Row |
|------|-----------|---------|------------|
| License | `dnpi_license_bcbs_sc_` | 28 | Row 3 |
| NPI | `dnpi_npi_bcbs_sc_` | 41 | Row 101 |
| Exclusion | `dnpi_exclusion_bcbs_sc_` | 43 | Row 110 |
| Preclusion | `dnpi_preclusion_bcbs_sc_` | 22 | Row 120 |
| Opt-Out | `dnpi_opt_out_bcbs_sc_` | 18 | Row 130 |
| OFAC | `dnpi_ofac_bcbs_sc_` | 22 | Row 140 |
| SSDMF | `dnpi_ssdmf_bcbs_sc_` | 15 | Row 150 |
| Missing | `npi_missing_license_creds` | 1 | Row 160 |

## Requirements

```bash
pip install pandas numpy openpyxl
```

## Configuration

Before running, update these variables in `xlsx-create.py`:

```python
input_dir = '/Users/tthreatt/Desktop/BCBS-SC'  # Directory containing CSV files
output_file = 'BCBS-251110-new.xlsx'           # Output filename
```

## Usage

1. Place all CSV files in the input directory
2. Run the script:
   ```bash
   python xlsx-create.py
   ```
3. Open the generated Excel file
4. Use the Search_Tab:
   - Enter an NPI or provider ID in cell **A1**
   - (For license data) Select an issuer filter in cell **B1**
   - View filtered results across all data sources

## How It Works

1. **CSV Processing**: Scans input directory for matching CSV files
2. **Sheet Creation**: Creates dated sheets for each data type (e.g., `license_25-11-05`)
3. **Search Tab**: Generates a Search_Tab with:
   - Bold headers for each data type
   - FILTER formulas that dynamically pull matching records
   - Data validation dropdowns where applicable
4. **Formula Generation**: Creates Excel FILTER formulas with proper range references
5. **Validation Lists**: Extracts unique issuers and creates dropdown (hidden sheet if needed)

## Output Structure

The generated Excel workbook contains:
- **Search_Tab**: Interactive search interface (first sheet)
- **Data Sheets**: One sheet per CSV file with dated names
- **IssuerList** (if needed): Hidden helper sheet for large dropdown lists

## Use Case

This tool is designed for healthcare compliance teams who need to:
- Monitor provider credentials across multiple databases
- Quickly search for provider information by NPI or ID
- Filter license data by issuing authority
- Maintain audit trails of provider status across exclusion lists, licenses, and credentials

## License

See LICENSE file for details.