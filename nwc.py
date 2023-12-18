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
  
# Define the categories for the Net Working Capital sheet from the image provided
categories = [
       "Cash & Cash Equivalents",
       "Accounts Receivable",
       "Inventory",
       "Other Current Assets",
       "Total Current Assets",
       " ",  # Blank row for spacing
       "Accounts Payable",
       "Other Payables & Accruals",
       "Short Term Debt",
       "Other Short Term Liabilities",
       "Total Current Liabilities",
       "", "", "", # Multiple blank rows to separate the tables visually
       "Assumptions",
       "Fiscal Year",
       "Revenue",
       "COGS",
       " ",
       "Days Sales Outstanding (DSO)",
       "Days Inventory Outstanding (DIO)",
       "Days Payable Outstanding (DPO)",
       " ",
       "Other Current Assets as a % of Revenue",
       "Short Term Debt as a % of Revenue",
       "Other Short Term Liabilities as a % of Revenue",
]


def AddNetWorkingCapitalSheet(start_year, end_year):
   # Generate fiscal years starting from start_year to end_year inclusive
   fiscal_years = list(range(start_year, end_year + 1)) + [f"{year}E" for year in range(end_year + 1, end_year + 6)]
   fiscal_years = list(map(str, fiscal_years))
  
   data = {"Category": categories}
  
   for i, year in enumerate(fiscal_years):
       is_estimated_year = year.endswith("E")
       data[year] = []
       for category in categories:
           col_letter = increment_letter('B', i)
           if category == "Accounts Receivable":
               formula = f"='Balance Sheet'!{col_letter}7" if not is_estimated_year else "AR Formula"
           elif category == "Cash & Cash Equivalents":
               formula = f"='Balance Sheet'!{col_letter}3" if not is_estimated_year else "C&CE Formula"
           elif category == "Inventory":
               formula = f"='Balance Sheet'!{col_letter}8" if not is_estimated_year else "Inventory Formula"
           elif category == "Other Current Assets":
               formula = f"='Balance Sheet'!{col_letter}9" if not is_estimated_year else "OCA Formula"
           elif category == "Total Current Assets":
               formula = f"='Balance Sheet'!{col_letter}11" if not is_estimated_year else "OpEx Formula"
           elif category == "Accounts Payable":
               formula = f"='Balance Sheet'!{col_letter}22 " if not is_estimated_year else "AP Formula"
           elif category == "Other Payables & Accruals":
               formula = f"='Balance Sheet'!{col_letter}23 " if not is_estimated_year else "OPA Formula"
           elif category == "Short Term Debt":
               formula = f"='Balance Sheet'!{col_letter}24 " if not is_estimated_year else "STD Formula"
           elif category == "Other Short Term Liabilities":
               formula = f"='Balance Sheet'!{col_letter}25 " if not is_estimated_year else "OSTL Formula"
           elif category == "Total Current Liabilities":
               formula = f"='Balance Sheet'!{col_letter}27 " if not is_estimated_year else "TCL Formula"
           elif category == "Revenue":
               formula = f"='Income Statement'!{col_letter}3 " if not is_estimated_year else "Revenue Formula"
           elif category == "COGS":
               formula = f"='Income Statement'!{col_letter}4 " if not is_estimated_year else "OSTL Formula"
           elif category == "Days Sales Outstanding (DSO)":
               formula = f"={col_letter}4 / {col_letter}19 * 365 " if not is_estimated_year else "DSO Formula"
           elif category == "Days Inventory Outstanding (DIO)":
               formula = f"={col_letter}5 / {col_letter}19 * 365 " if not is_estimated_year else "DIO Formula"
           elif category == "Days Payable Outstanding (DPO)":
               formula = f"={col_letter}9 / {col_letter}19 * 365 " if not is_estimated_year else "DPO Formula"
           elif category == "Other Current Assets as a % of Revenue":
               formula = f"={col_letter}6 / {col_letter}19 " if not is_estimated_year else "OCA%Rev Formula"
           elif category == "Short Term Debt as a % of Revenue":
               formula = f"={col_letter}11 / {col_letter}19 " if not is_estimated_year else "STD%Rev Formula"
           elif category == "Other Short Term Liabilities as a % of Revenue":
               formula = f"={col_letter}12 / {col_letter}19 " if not is_estimated_year else "OSTL%Rev Formula"
           else:   
               formula = "..."  # Placeholder for other categories
           data[year].append(formula)
          
   df = pd.DataFrame(data)
   df.set_index('Category', inplace=True)
   return df

def CreateSheet(writer, sheet, sheetName):
    sheet.to_excel(writer, sheet_name=sheetName, startrow=1)
    nwc_worksheet = writer.sheets[sheetName]
    styleModule.SetOtherStyle(nwc_worksheet)
    
    # Adjust column widths
    for column_cells in nwc_worksheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells if cell.value) + 2
        nwc_worksheet.column_dimensions[column_cells[0].column_letter].width = length
  
    # Set the title in cell A1
    title_cell = nwc_worksheet.cell(row=1, column=1)
    title_cell.value = "Net Working Capital"

    # Style the title cell: white text on a lighter shade of blue
    title_cell.font = Font(color=Color("FFFFFF"), bold=True, size=14)  # White text, bold, and larger font size
    title_cell.fill = PatternFill(start_color="1A759C", end_color="1A759C", fill_type="solid")  # Light blue background
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
