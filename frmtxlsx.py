
# -*- coding: utf-8 -*-
"""
Created on Tue Apr 11 10:59:16 2023

@author: VShields
"""
import pandas as pd
import argparse

class InvalidFileTypeError(Exception):
    pass

class MyFormatter:
    
    def __init__(self, data, workbook, worksheet):
        self.data = data
        self.workbook = workbook
        self.worksheet = worksheet
    
    def __repr__(self):
        return f'MyFormatter(data={self.data}, workbook={self.workbook}, worksheet={self.worksheet})'
    
    def __str__(self):
        return f'MyFormatter(data={self.data}, workbook={self.workbook}, worksheet={self.worksheet})'
    
    def check_format(self):
        """
        Checks the format of the data and performs necessary transformations if needed.

        If the data is a pandas DataFrame with multiple index levels, it resets the index.
        If the data is a pandas DataFrameGroupBy object, it applies a lambda function to reset the index and then resets the index.
        This is so the spacing and formatting functions can be applied to the data in a groupby object.

        Raises:
            TypeError: If the data is not a pd.DataFrame or a pd.core.groupby.DataFrameGroupBy.

        """
        if isinstance(self.data, pd.DataFrame):
            if self.data.index.nlevels > 1:
                self.data = self.data.reset_index()
        elif isinstance(self.data, pd.core.groupby.DataFrameGroupBy):
            self.data = self.data.apply(lambda x: x.reset_index(drop=True))
            self.data = self.data.reset_index()
        else:
            raise TypeError("The data must be a pd.DataFrame or a pd.core.groupby.DataFrameGroupBy.")
    
    def spacing(self):
        """
        Adjusts the column widths in the worksheet based on the data length.

        This method calculates the maximum width of each column in the data and sets the corresponding column width in the worksheet.
        The width is determined by the length of the longest value in the column, including the column header.

        Returns:
            None
        """
        self.check_format()

        for i, col in enumerate(self.data.columns):
            max_width = max(self.data[col].astype(str).str.len().max(), len(col))
            self.worksheet.set_column(i, i, max_width + 1)

    def color_columns(self, grey_bottom=False, color='#ffb3b3'):
        """
        Option to add background color to the column headers. 
        

        Parameters
        ----------
        grey_bottom : bool, default False
            Specifies whether to color the bottom row as well.
            Useful for pivot/summary tables.
            If True, the bottom row will be grey.
            If False, only the column headers will be colored.

        color : str, default '#ffb3b3' (light red)
            The background color to be applied to the column headers.
            This should be a valid CSS color value.
            See also: https://xlsxwriter.readthedocs.io/working_with_colors.html

        Returns
        -------
        Formatted worksheet with colored column headers and optional grey bottom row.
        """

        if grey_bottom:
            last_row = self.data.shape[0]
            last_column = self.data.shape[1]
            last_row_format = self.workbook.add_format({'bg_color': '#d9d9d9', 'bold': True})
            self.worksheet.conditional_format(last_row, 0, last_row, last_column, {'type': 'no_blanks', 'format': last_row_format})

        column_format = self.workbook.add_format({'bg_color': color})
        self.worksheet.conditional_format(0, 0, 0, len(self.data.columns), {'type': 'no_blanks', 'format': column_format})

    def add_borders(self):
        """
        Adds borders to the cells in the worksheet.

        This method checks the format of the workbook, creates a border format with black color, and applies it to the cells in the worksheet.

        Parameters:
            None

        Returns:
            None
        """
        self.check_format()

        border_format = self.workbook.add_format({'border': 1, 'border_color': 'black'})
        border_format.set_locked(False)
        self.worksheet.conditional_format(1, 0, len(self.data), self.data.shape[1]-1, {'type': 'no_errors', 'format': border_format})
        
    def panes(self):
        self.worksheet.freeze_panes(1,1)



def main():
    parser = argparse.ArgumentParser(description='Process an Excel file.')
    parser.add_argument('filename', type=str, help='The Excel file to process.')
    parser.add_argument('--add_borders', action='store_true', help='Add borders to cells.')
    parser.add_argument('--panes', action='store_true', help='Freeze panes.')
    parser.add_argument('--color_columns', action='store_true', help='Color columns.')
    parser.add_argument('--spacing', action='store_true', help='add spacing to Excel columns based on the maximum length of the data in each column.')

    args = parser.parse_args()
    
    if not args.filename.endswith('.xlsx'):
        raise InvalidFileTypeError('The file to format must be a .xlsx file.')
    
    df = pd.read_excel(args.filename,index=False)
    with pd.ExcelWriter(args.filename) as writer:
        
        workbook = writer.book
        df.to_excel(writer, sheetname="Sheet1", index=False)
        worksheet = writer.sheets['Sheet1']

        formatter = MyFormatter(data, workbook, worksheet)

        # Call the appropriate functions based on the arguments
        if args.add_borders:
            formatter.add_borders()
        if args.panes:
            formatter.panes()
        if args.color_columns:
            formatter.color_columns()
        if args.color_columns:
            formatter.spacing()

if __name__ == '__main__':
    main()             
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    
