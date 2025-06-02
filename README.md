# MMA Analytics Project

This project automates analysis and reporting of MMA club member data using Python.

## Project Structure

```
EDVANTIS/
├── project1_data_processing.py       # Parse and analyze raw Excel data, calculate efficiency
├── project2_charts_excel.py          # Generate charts (bar, pie, scatter) in Excel format
├── project3_presentation.py          # Create PowerPoint presentation with stats and summary
├── pyproject.toml                    # Poetry config file with dependencies
├── README.md                         # This file
└── outputs/                          # Generated files (Excel + PowerPoint)
    ├── processed_stats.xlsx
    ├── MMA_Club_Charts.xlsx
    └── MMA_Club_Stats_Presentation.pptx
```

## Setup Instructions

Make sure you have [Poetry](https://python-poetry.org/) installed.

### 1. Install dependencies

```bash
poetry install
```

### 2. Activate environment

```bash
poetry shell
```

### 3. Run scripts

Each script will automatically create the `outputs/` folder and save results there.

#### Step 1: Process original Excel data

```bash
python project1_data_processing.py
```

#### Step 2: Generate charts in Excel

```bash
python project2_charts_excel.py
```

#### Step 3: Generate PowerPoint presentation

```bash
python project3_presentation.py
```

## Notes

- You should place the original Excel file `MMA_Club_Stats_20_Students.xlsx` in the root directory before running `project1_data_processing.py`.
- All generated files will be stored in the `outputs/` folder.
