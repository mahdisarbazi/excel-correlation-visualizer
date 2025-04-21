An open-source Python tool to visualize correlation matrices from Excel files instantly. Perfect for researchers and analysts!

# Excel Correlation Visualizer

A Python desktop application that allows users to visualize correlations between numeric variables in Excel files. The application provides a simple graphical user interface to select Excel files, preview data, and generate correlation heatmaps and scatter plots.

## Features

- Load and preview Excel files
- Auto-detect header rows and data starting points
- Generate correlation matrix heatmap
- Create scatter plots for all pairs of numeric variables
- Display regression lines and correlation coefficients on scatter plots
- Save individual or all plots to image files

## Requirements

- Python 3.6+
- Dependencies:
  - pandas
  - matplotlib
  - numpy
  - seaborn
  - tkinter (usually comes with Python)
  - Pillow (PIL)

## Installation

1. Clone this repository or download the source code
2. Install the required packages:

```bash
pip install pandas matplotlib numpy seaborn Pillow
```

## Usage

Run the application by executing:

```bash
python excel-correlation-visualizer.py
```

### Steps to Use:

1. Click "Select Excel File" to choose an Excel file
2. Select a sheet from the file
3. Preview the data and specify header row and data start row (or use auto-detection)
4. View the correlation matrix and scatter plots
5. Save plots individually or all at once

## Data Handling

- The application attempts to convert non-numeric columns to numeric when possible
- Only numeric columns are included in the correlation analysis
- NaN values are automatically handled in the correlation calculations

## Credits

- Original creator: Mahdi Sarbazi
- Modified for English interface

## License

This project is open source and available under the [MIT License](LICENSE).
