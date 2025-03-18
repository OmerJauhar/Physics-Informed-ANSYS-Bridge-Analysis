# ANSYS PDF Report Extractor

This Python tool extracts engineering parameters from ANSYS PDF reports using OpenRouter API and populates an Excel file with the extracted data. Designed specifically for the "Max Failure load of spaghetti bridge using PINN" project.

## Features

- Searches a directory for ANSYS report PDF files
- Uses pdfplumber to extract text content from PDF files
- Uses OpenRouter API to intelligently extract parameters from the reports
- Performs calculations for derived parameters
- Populates an Excel file with extracted data
- Increments Bridge ID for each new report

## Requirements

- Python 3.6+
- Required packages:
  - pandas
  - requests
  - numpy
  - openpyxl (for Excel file handling)
  - pdfplumber (for PDF text extraction)
  - reportlab (for test PDF generation)

## Installation

1. Clone the repository or download the source code
2. Install the required packages:

```bash
pip install pandas requests numpy openpyxl pdfplumber reportlab
```

## Usage

Run the script with the following command:

```bash
python ansys_report_extractor.py --api_key YOUR_API_KEY --excel_template path/to/template.xlsx --reports_dir path/to/reports --output path/to/output.xlsx
```

### Parameters:

- `--api_key`: Your OpenRouter API key (required)
- `--excel_template`: Path to the Excel template file (required)
- `--reports_dir`: Directory containing ANSYS PDF reports (required)
- `--output`: Path where the populated Excel file will be saved (required)

## Testing

You can test the functionality using the included test script:

```bash
python test_ansys_extractor.py
```

This will create sample ANSYS PDF reports, process them, and generate a test output file.

## Parameter Extraction

The script extracts the following parameters from ANSYS PDF reports:

- Bridge Length (m)
- Bridge Width (m) 
- Bridge Height (m)
- Cross-Sectional Diameter (m)
- Cross-Sectional Area (m²)
- Cross-Sectional Moment of Inertia (m⁴)
- Number of Strands (always set to 6)
- Number of Beams (always set to 2)
- Angle of Inclination (°) (always set to 0)
- Angle of Declination (°) (always set to 0)
- Young's Modulus (Pa)
- Poisson's Ratio
- Density (kg/m³)
- Tensile Yield Strength (Pa)
- Shear Modulus (Pa)
- Applied Force (N)
- Mesh Elements
- Mesh Density (elements/m³)
- Max Equivalent Stress (Pa)
- Max Principal Stress (Pa)
- Min/Max Deformation (m)
- Safety Factor
- Reaction Forces (N)
- Strain Energy (J)

It also calculates these derived parameters:

- Work Done (J): 0.5 * |Applied Force| * Max Deformation
- Energy Residual (J): |Strain Energy - Work Done|
- Yield Constraint Residual: ReLU(Max Equivalent Stress - Tensile Yield Strength)
- Max Failure Load (kg): Applied Force (N) / 9.8

## Notes

- The script looks for PDF files in the specified directory and its subdirectories.
- If a parameter cannot be extracted from the report, it will be marked as 'N/A' in the Excel file.
- The OpenRouter API call is configured to use the Anthropic Claude 3 Opus model for optimal parameter extraction.