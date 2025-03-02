# Crude Petroleum Export Data Analysis

This project analyzes crude petroleum export data for European countries, generating statistical insights and visualizations.

## Requirements

- Python 3.8 or higher
- Conda package manager

## Setup

1. Create a new conda environment:
```bash
conda create -n oil-analysis python=3.8
conda activate oil-analysis
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Data Requirements

Place your Excel file named `data.xlsx` in the project root directory. The file should contain:

- Sheet name: `Exporters-of-Crude-Petroleum-2`
- Required columns:
  - `Continent`
  - `Country`
  - `ISO 3`
  - `Trade Value`

## Running the Analysis

1. Ensure your Excel file is in the project directory
2. Run the analysis:
```bash
python main.py
```

## Outputs

The script will:
1. Print summary statistics to the console
2. Print a list of top exporters (within 10% of maximum value)
3. Generate visualizations in the `output` directory:
   - Bar chart of export values by country
   - Histogram of export value distribution

## Error Handling

The script includes error handling for:
- Missing input file
- Invalid data formats
- Missing values in the Trade Value column 