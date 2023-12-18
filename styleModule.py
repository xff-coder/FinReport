import numpy as np
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, NamedStyle, Border, Side, Color
from openpyxl.styles.numbers import FORMAT_CURRENCY_USD_SIMPLE, FORMAT_PERCENTAGE
from openpyxl.utils import get_column_letter
from openpyxl.cell import MergedCell
from string import ascii_uppercase

# Define the currency format style
currency_style = NamedStyle(name='currency_style')
currency_style.number_format = FORMAT_CURRENCY_USD_SIMPLE
currency_style.alignment = Alignment(horizontal='right')


# Define a style for fiscal year
fiscal_year_style = NamedStyle(name='fiscal_year_style')
fiscal_year_style.font = Font(color=Color(rgb='002596BE'), bold=True) # RGB for the desired blue color
fiscal_year_style.alignment = Alignment(horizontal='center', vertical='center') # Center alignment
# Define a thick top border style
thick_top_border = Border(top=Side(style='thick'))


#define styling for fixed assets
header_style = NamedStyle(name="header_style")
header_style.font = Font(bold=True, size=11)
header_style.alignment = Alignment(horizontal="center", vertical="center")
header_style.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
header_style.border = Border(bottom=Side(border_style="thin"))


data_style = NamedStyle(name="data_style")
data_style.font = Font(size=11)
data_style.alignment = Alignment(horizontal="right", vertical="center")
data_style.number_format = '#,##0.00'


title_style = NamedStyle(name="title_style")
title_style.font = Font(bold=True, size=11, color=Color("FFFFFF"))
title_style.alignment = Alignment(horizontal="center", vertical="center")
title_style.fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")


def SetOtherStyle(worksheet):
   # Define fill colors for alternating rows
   grey_fill = PatternFill(start_color='00CCCCCC', end_color='00CCCCCC', fill_type='solid')
   white_fill = PatternFill(start_color='00FFFFFF', end_color='00FFFFFF', fill_type='solid')
                  
   # Apply the alternating row colors
   for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, max_row=worksheet.max_row), start=2):
       fill = grey_fill if row_idx % 2 == 0 else white_fill
       for cell in row:
           cell.fill = fill

   # Set font and alignment for header
   for cell in worksheet['1:1']:
       cell.font = Font(bold=True)
       cell.alignment = Alignment(horizontal='center')

   # Set number format for data cells and align column A to the left
   for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
       for cell in row:
           if cell.column == 1:  # Column A
               cell.alignment = Alignment(horizontal='left')
           else:
               cell.number_format = FORMAT_CURRENCY_USD_SIMPLE

   # Auto-adjust column widths
   for col in worksheet.columns:
       max_length = 0
       column_letter = None

       for cell in col:
           # Skip merged cells
           if isinstance(cell, MergedCell):
               continue

           column_letter = cell.column_letter  # Get the column name for non-merged cells
           try:
               if len(str(cell.value)) > max_length:
                   max_length = len(str(cell.value))
           except:
               pass

       # Adjust the column width if the column letter was found
       if column_letter:
           adjusted_width = (max_length + 2)
           worksheet.column_dimensions[column_letter].width = adjusted_width

  # Add a note at the bottom of each sheet
   last_row = worksheet.max_row + 2
   note_cell = worksheet.cell(row=last_row, column=1)
   note_cell.value = "*$ Expressed in millions"
   note_cell.font = Font(italic=True)
   note_cell.alignment = Alignment(horizontal='left')
  
   # Apply the fiscal year style to all cells in the Fiscal Year column
   for cell in worksheet[2]:  # Iterate over all cells in the second row
       if cell.column_letter != 'A':  # Skip the header cell
           cell.value = f'FY {cell.value}'
           cell.style = fiscal_year_style 
