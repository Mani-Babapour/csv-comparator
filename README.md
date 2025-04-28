# CSV Comparator

A Python tool to automatically compare two CSV files inside each subfolder and generate a detailed Excel report summarizing the differences.

## Features

- Scans all subfolders within a main directory.
- Compares two CSV files inside each subfolder.
- Detects:
  - Full matches ✅
  - Differences in data or structure ❌
- Generates a final `comparison_result.xlsx` report.
- Supports Persian (Farsi) language in CSV files.

## Requirements

- Python 3.7+
- Libraries:
  - `pandas`

Install dependencies:

```bash
pip install pandas
