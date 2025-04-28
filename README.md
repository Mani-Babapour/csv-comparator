# CSV Comparator

A Python tool to automatically compare two CSV files inside each subfolder, highlight the differences, and generate a detailed Excel report summarizing the results.

## Features

- Scans all subfolders within a main directory.
- Compares two CSV files inside each subfolder.
- Detects:
  - Full matches ✅
  - Differences in data or structure ❌
  - Highlights differences with **red background** for easy identification.
- Generates a final `comparison_result.xlsx` report summarizing the comparison results.
- Saves individual difference files for discrepancies.
- Supports Persian (Farsi) language in CSV files.

## Requirements

- Python 3.7+
- Libraries:
  - `pandas`
  - `openpyxl`

To install the required dependencies, run the following:

```bash
pip install pandas openpyxl
