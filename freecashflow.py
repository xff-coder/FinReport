import numpy as np
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, NamedStyle, Border, Side, Color
from openpyxl.styles.numbers import FORMAT_CURRENCY_USD_SIMPLE, FORMAT_PERCENTAGE
from openpyxl.utils import get_column_letter
from string import ascii_uppercase
import styleModule

def increment_letter(letter, increment):
   # Function to increment a letter by a specified number
   if increment == 0:
       return letter
   else:
       alpha_index = ascii_uppercase.index(letter.upper()) + increment
       return ascii_uppercase[alpha_index % 26] + "A" * (alpha_index // 26)

categories = [
   "Revenue", "COGS", "Gross Profit", "Operating Expenses",
   "Selling, General, Administrative", "Total Operating Expenses",
   "EBITDA", "Depreciation & Amortization", "Operating Profit (EBIT)",
   "Operating Taxes", "NOPAT (Net Operating Profit After Taxes)",
   "(+) Depreciation & Amortization", "(-) Capital Expenditures",
   "(-) Change in NWC", "NWC", "Current Assets", "Current Liabilities"
]


def find_last_non_estimated_ending_PPE_cell(start_year, end_year):
   # Determine the last non-estimated fiscal year
   last_non_estimated_year = end_year


   # Calculate the number of increments needed from the base column ('C' in this case)
   increments = last_non_estimated_year - (start_year + 1)


   # Find the column letter
   col_letter = increment_letter('C', increments)


   # Assuming "Ending PP&E" is in a specific row (e.g., row 12)
   row_number = 12


   # Construct the cell reference
   cell_reference = f"{col_letter}{row_number}"
   return cell_reference
  
def AddFreeCashFlowSheet(start_year, end_year):
   # Adjust the start year for the free cash flow sheet
   adjusted_start_year = start_year + 1


   # Generate fiscal years starting from adjusted_start_year
   fiscal_years = list(range(adjusted_start_year, end_year + 1)) + [f"{year}E" for year in range(end_year + 1, end_year + 6)]
   fiscal_years = list(map(str, fiscal_years))
  
   data = {"Category": categories}
  
   for i, year in enumerate(fiscal_years):
       is_estimated_year = year.endswith("E")
       data[year] = []
       for category in categories:
           col_letter = increment_letter('B', i)
           if category == "Revenue":
               formula = f"='Income Statement'!{col_letter}3" if not is_estimated_year else "Rev Formula"
           elif category == "COGS":
               formula = f"='Income Statement'!{col_letter}4" if not is_estimated_year else "COGS Formula"
           elif category == "Gross Profit":
               formula = f"={col_letter}3 + {col_letter}4" if not is_estimated_year else "GP Formula"
           elif category == "Operating Expenses":
               formula = f"='Income Statement'!{col_letter}6" if not is_estimated_year else "OpEx Formula"
           elif category == "Selling, General, Administrative":
               formula = f"='Income Statement'!{col_letter}7" if not is_estimated_year else "SG&A Formula"
           elif category == "Total Operating Expenses":
               formula = f"=SUM('Income Statement'!{col_letter}6:'Income Statement'!{col_letter}9)" if not is_estimated_year else "TotalOpEx Formula"
           elif category == "EBITDA":
               formula = f"={col_letter}5 + {col_letter}8" if not is_estimated_year else "EBITDA Formula"
           elif category == "Depreciation & Amortization":
               formula = f"='Balance Sheet'!{col_letter}4" if not is_estimated_year else "D&A Formula"
           elif category == "Operating Profit (EBIT)":
               formula = f"={col_letter}9 - {col_letter}10" if not is_estimated_year else "EBIT Formula"
           elif category == "Operating Taxes":
               formula = '=0.21'
           elif category == "NOPAT (Net Operating Profit After Taxes)":
               formula = f"={col_letter}11 * (1 - {col_letter}12)" if not is_estimated_year else "NOPAT Formula"
           elif category == "(+) Depreciation & Amortization":
               formula = f"='Balance Sheet'!{col_letter}4" if not is_estimated_year else "AddBack D&A Formula"
           elif category == "(-) Capital Expenditures":
               formula = f"='Fixed Assets'!{col_letter}5" if not is_estimated_year else "CapEx Formula"
           else:   
               formula = "..."  # Placeholder for other categories
           data[year].append(formula)
          


   df = pd.DataFrame(data)
   df.set_index('Category', inplace=True)
   return df

def CreateSheet(writer, sheet, sheetName, start_year, end_year):
      # Free Cash Flow Sheet

    fiscal_years = list(range(start_year + 1, end_year + 1)) + [f"{year}E" for year in range(end_year + 1, end_year + 6)]
    fiscal_years = list(map(str, fiscal_years))
    
    sheet.to_excel(writer, sheet_name=sheetName, startrow=1)
    fcf_worksheet = writer.sheets[sheetName]

    header_row_num = 2
    for col_num in range(3, len(fiscal_years) + 2):  # Adjust the range based on your fiscal year columns
        cell = fcf_worksheet.cell(row=header_row_num, column=col_num)
        cell.style = styleModule.fiscal_year_style

    # Apply the same style as other sheets to the Free Cash Flow sheet
    styleModule.SetOtherStyle(fcf_worksheet)  # Assuming this is the function for applying general styles

    # Adjust column widths for the Free Cash Flow sheet
    for column_cells in fcf_worksheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells if cell.value) + 2
        fcf_worksheet.column_dimensions[column_cells[0].column_letter].width = length


    # Set the title in cell A1
    title_cell = fcf_worksheet.cell(row=1, column=1)
    title_cell.value = "Unlevered Free Cash Flow"

    # Style the title cell: white text on a lighter shade of blue
    title_cell.font = Font(color=Color("FFFFFF"), bold=True, size=14)  # White text, bold, and larger font size
    title_cell.fill = PatternFill(start_color="1A759C", end_color="1A759C", fill_type="solid")  # Light blue background
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    # Net Working Capital sheet

