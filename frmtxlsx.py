#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Formatter CLI Tool

A command-line tool for formatting Excel files with various styling options.

@author: VShields

Usage:
    python frmtxlsx.py input.xlsx --output output.xlsx --spacing --borders --color-columns
"""

import argparse
import logging
import sys
from pathlib import Path
from typing import Optional, Union

import pandas as pd


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class InvalidFileTypeError(Exception):
    """Raised when an invalid file type is provided."""
    pass


class FileNotFoundError(Exception):
    """Raised when the input file doesn't exist."""
    pass


class ExcelFormatterError(Exception):
    """Base exception for Excel formatting errors."""
    pass


class MyFormatter:
    """
    Excel formatter class for applying various formatting styles to Excel worksheets.
    
    Attributes:
        data: pandas DataFrame containing the data
        workbook: xlsxwriter workbook object
        worksheet: xlsxwriter worksheet object
    """
    
    def __init__(self, data: pd.DataFrame, workbook, worksheet):
        self.data = data
        self.workbook = workbook
        self.worksheet = worksheet
        self._validated = False

    def __repr__(self) -> str:
        return f"MyFormatter(rows={len(self.data)}, cols={len(self.data.columns)})"

    def __str__(self) -> str:
        return f"MyFormatter with {len(self.data)} rows and {len(self.data.columns)} columns"

    def _validate_data(self) -> None:
        """
        Validates and prepares the data for formatting operations.
        
        Raises:
            TypeError: If the data is not a supported type.
            ExcelFormatterError: If data validation fails.
        """
        if self._validated:
            return
            
        try:
            if isinstance(self.data, pd.DataFrame):
                if self.data.index.nlevels > 1:
                    logger.info("Resetting multi-level index")
                    self.data = self.data.reset_index()
            elif isinstance(self.data, pd.core.groupby.DataFrameGroupBy):
                logger.info("Processing GroupBy object")
                self.data = self.data.apply(lambda x: x.reset_index(drop=True))
                self.data = self.data.reset_index()
            else:
                raise TypeError(
                    f"Unsupported data type: {type(self.data)}. "
                    "Expected pd.DataFrame or pd.core.groupby.DataFrameGroupBy."
                )
            
            if self.data.empty:
                raise ExcelFormatterError("Cannot format empty DataFrame")
                
            self._validated = True
            logger.debug(f"Data validation successful: {self.data.shape}")
            
        except Exception as e:
            logger.error(f"Data validation failed: {e}")
            raise ExcelFormatterError(f"Data validation failed: {e}")

    def apply_spacing(self) -> None:
        """
        Adjusts column widths based on content length.
        
        Calculates the maximum width needed for each column and applies it to the worksheet.
        Includes a small padding for better readability.
        """
        try:
            self._validate_data()
            
            logger.info("Applying column spacing")
            for i, col in enumerate(self.data.columns):
                # Calculate max width including header
                col_data = self.data[col].astype(str)
                max_width = max(
                    col_data.str.len().max() if not col_data.empty else 0,
                    len(str(col))
                )
                # Add padding and set reasonable limits
                width = min(max(max_width + 2, 8), 50)  # Min 8, max 50 chars
                self.worksheet.set_column(i, i, width)
                
            logger.debug(f"Applied spacing to {len(self.data.columns)} columns")
            
        except Exception as e:
            logger.error(f"Failed to apply spacing: {e}")
            raise ExcelFormatterError(f"Spacing operation failed: {e}")

    def apply_column_colors(self, grey_bottom: bool = False, header_color: str = "#ffb3b3") -> None:
        """
        Applies background colors to column headers and optionally the bottom row.

        Args:
            grey_bottom: Whether to color the bottom row grey (useful for totals)
            header_color: CSS color value for column headers

        Raises:
            ExcelFormatterError: If color application fails
        """
        try:
            self._validate_data()
            
            logger.info(f"Applying column colors (header: {header_color}, grey_bottom: {grey_bottom})")
            
            # Apply grey bottom row if requested
            if grey_bottom and len(self.data) > 0:
                last_row = len(self.data)  # Adjust for 0-indexed + header
                last_column = len(self.data.columns) - 1
                last_row_format = self.workbook.add_format({
                    "bg_color": "#d9d9d9", 
                    "bold": True
                })
                self.worksheet.conditional_format(
                    last_row, 0, last_row, last_column,
                    {"type": "no_errors", "format": last_row_format}
                )

            # Apply header colors
            if len(self.data.columns) > 0:
                header_format = self.workbook.add_format({"bg_color": header_color})
                self.worksheet.conditional_format(
                    0, 0, 0, len(self.data.columns) - 1,
                    {"type": "no_errors", "format": header_format}
                )
                
            logger.debug("Column colors applied successfully")
            
        except Exception as e:
            logger.error(f"Failed to apply column colors: {e}")
            raise ExcelFormatterError(f"Color operation failed: {e}")

    def apply_borders(self, border_style: int = 1, border_color: str = "black") -> None:
        """
        Adds borders to all data cells (excluding headers).

        Args:
            border_style: Border style (1=thin, 2=medium, etc.)
            border_color: Border color

        Raises:
            ExcelFormatterError: If border application fails
        """
        try:
            self._validate_data()
            
            if len(self.data) == 0:
                logger.warning("No data rows to apply borders to")
                return
                
            logger.info("Applying cell borders")
            border_format = self.workbook.add_format({
                "border": border_style,
                "border_color": border_color
            })
            
            # Apply to data rows only (skip header row 0)
            self.worksheet.conditional_format(
                1, 0,  # Start from row 1 (after header)
                len(self.data), len(self.data.columns) - 1,
                {"type": "no_errors", "format": border_format}
            )
            
            logger.debug(f"Applied borders to {len(self.data)} rows")
            
        except Exception as e:
            logger.error(f"Failed to apply borders: {e}")
            raise ExcelFormatterError(f"Border operation failed: {e}")

    def freeze_panes(self, row: int = 1, col: int = 0) -> None:
        """
        Freezes panes at the specified position.
        
        Args:
            row: Row to freeze after (default 1 for header)
            col: Column to freeze after (default 0 for no column freeze)
        """
        try:
            logger.info(f"Freezing panes at row {row}, column {col}")
            self.worksheet.freeze_panes(row, col)
            logger.debug("Panes frozen successfully")
            
        except Exception as e:
            logger.error(f"Failed to freeze panes: {e}")
            raise ExcelFormatterError(f"Freeze panes operation failed: {e}")


def validate_input_file(filepath: str) -> Path:
    """
    Validates the input file path and type.
    
    Args:
        filepath: Path to the input file
        
    Returns:
        Path object for the validated file
        
    Raises:
        InvalidFileTypeError: If file is not .xlsx
        FileNotFoundError: If file doesn't exist
    """
    file_path = Path(filepath)
    
    if not file_path.exists():
        raise FileNotFoundError(f"Input file not found: {filepath}")
    
    if file_path.suffix.lower() != '.xlsx':
        raise InvalidFileTypeError(
            f"Invalid file type: {file_path.suffix}. Only .xlsx files are supported."
        )
    
    return file_path


def load_excel_data(filepath: Path, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """
    Loads Excel data with error handling.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Specific sheet to load (default: first sheet)
        
    Returns:
        DataFrame containing the Excel data
        
    Raises:
        ExcelFormatterError: If loading fails
    """
    try:
        logger.info(f"Loading Excel file: {filepath}")
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        logger.info(f"Loaded data: {df.shape[0]} rows, {df.shape[1]} columns")
        return df
        
    except Exception as e:
        logger.error(f"Failed to load Excel file: {e}")
        raise ExcelFormatterError(f"Could not load Excel file {filepath}: {e}")


def create_output_path(input_path: Path, output_path: Optional[str] = None) -> Path:
    """
    Creates the output file path.
    
    Args:
        input_path: Input file path
        output_path: Specified output path (optional)
        
    Returns:
        Path object for output file
    """
    if output_path:
        return Path(output_path)
    else:
        # Create default output name
        stem = input_path.stem
        return input_path.parent / f"{stem}_formatted.xlsx"


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Format Excel files with various styling options.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s data.xlsx --spacing --borders
  %(prog)s input.xlsx --output formatted.xlsx --color-columns --freeze-panes
  %(prog)s file.xlsx --all --verbose
        """
    )
    
    # Required arguments
    parser.add_argument(
        "filename", 
        type=str, 
        help="Path to the Excel file to process (.xlsx only)"
    )
    
    # Optional arguments
    parser.add_argument(
        "--output", "-o",
        type=str,
        help="Output file path (default: input_formatted.xlsx)"
    )
    
    parser.add_argument(
        "--sheet",
        type=str,
        help="Specific sheet name to process (default: first sheet)"
    )
    
    # Formatting options
    parser.add_argument(
        "--borders", "--add-borders",
        action="store_true",
        help="Add borders to data cells"
    )
    
    parser.add_argument(
        "--freeze-panes", "--panes",
        action="store_true",
        help="Freeze the header row"
    )
    
    parser.add_argument(
        "--color-columns",
        action="store_true",
        help="Apply background color to column headers"
    )
    
    parser.add_argument(
        "--spacing",
        action="store_true",
        help="Auto-adjust column widths based on content"
    )
    
    parser.add_argument(
        "--grey-bottom",
        action="store_true",
        help="Color the bottom row grey (useful for totals)"
    )
    
    parser.add_argument(
        "--all",
        action="store_true",
        help="Apply all formatting options"
    )
    
    # Utility arguments
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose logging"
    )
    
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress all output except errors"
    )

    args = parser.parse_args()
    
    # Configure logging level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    elif args.quiet:
        logging.getLogger().setLevel(logging.ERROR)
    
    try:
        # Validate input
        input_path = validate_input_file(args.filename)
        output_path = create_output_path(input_path, args.output)
        
        # Load data
        df = load_excel_data(input_path, args.sheet)
        
        # Create formatted Excel file
        logger.info(f"Creating formatted file: {output_path}")
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            workbook = writer.book
            df.to_excel(writer, sheet_name="Sheet1", index=False)
            worksheet = writer.sheets["Sheet1"]
            
            formatter = MyFormatter(df, workbook, worksheet)
            
            # Apply formatting options
            if args.all:
                formatter.apply_spacing()
                formatter.apply_borders()
                formatter.apply_column_colors(grey_bottom=args.grey_bottom)
                formatter.freeze_panes()
            else:
                if args.spacing:
                    formatter.apply_spacing()
                if args.borders:
                    formatter.apply_borders()
                if args.color_columns:
                    formatter.apply_column_colors(grey_bottom=args.grey_bottom)
                if args.freeze_panes:
                    formatter.freeze_panes()
        
        logger.info(f"Successfully created formatted file: {output_path}")
        print(f"Formatted Excel file saved as: {output_path}")
        
    except (InvalidFileTypeError, FileNotFoundError, ExcelFormatterError) as e:
        logger.error(f"Error: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        logger.info("Operation cancelled by user")
        sys.exit(130)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        if args.verbose:
            raise
        sys.exit(1)


if __name__ == "__main__":
    main()