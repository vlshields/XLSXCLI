# XLSXCLI

This Python script provides a class `MyFormatter` for formatting Excel files. It includes methods for checking data format, adjusting column widths, coloring column headers, adding borders to cells, and freezing panes.

## Usage

The script can be run from the command line with the following arguments:

- `filename`: The Excel file to process. This must be a .xlsx file.
- `--add_borders`: Add borders to cells.
- `--panes`: Freeze the first row and column.
- `--color_columns`: Color the column headers.
- `--spacing`: Adjust column widths based on the maximum length of the data in each column.

  ```shell
  python frmtxlsx.py "test.xlsx" --add_borders --panes --spacing --color_columns
  ```

## Class MyFormatter

The `MyFormatter` class takes a pandas DataFrame, a workbook, and a worksheet as input. It provides the following methods:

- `check_format()`: Checks the format of the data and performs necessary transformations if needed.
- `spacing()`: Adjusts the column widths in the worksheet based on the data length.
- `color_columns(grey_bottom=False, color="#ffb3b3")`: Adds background color to the column headers. Optionally colors the bottom row as well.
- `add_borders()`: Adds borders to the cells in the worksheet.
- `panes()`: Freezes the first row and column of the worksheet.

## Exceptions

The script defines the following exception:

- `InvalidFileTypeError`: Raised when the input file is not a .xlsx file.

## Dependencies

This script requires the following Python packages:

- pandas
- argparse

## Note

This script uses the xlsxwriter engine for pandas ExcelWriter, which provides more formatting options but may not be compatible with all Excel features. If you encounter issues, you may need to adjust the script to use a different engine.
