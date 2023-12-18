import numpy as np
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, NamedStyle, Border, Side, Color
from openpyxl.styles.numbers import FORMAT_CURRENCY_USD_SIMPLE, FORMAT_PERCENTAGE
from openpyxl.utils import get_column_letter
from string import ascii_uppercase


def increment_letter(letter, increment):
   # Function to increment a letter by a specified number
   if increment == 0:
       return letter
   else:
       alpha_index = ascii_uppercase.index(letter.upper()) + increment
       return ascii_uppercase[alpha_index % 26] + "A" * (alpha_index // 26)




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




def AddSheet(start_year, end_year):
   # Adjust the start year for the fixed assets sheet
   adjusted_start_year = start_year + 1
   row_number_for_DA_percentage = 7
   row_number_for_CapEx_percentage = 8


   # Generate fiscal years starting from adjusted_start_year
   fiscal_years = list(range(adjusted_start_year, end_year + 1)) + [f"{year}E" for year in range(end_year + 1, end_year + 6)]
   fiscal_years = list(map(str, fiscal_years))


   categories = ["Beginning PP&E", "D&A", "CapEx", "Ending PP&E", "D&A as a % of Beginning PP&E", "CapEx as a % of Beginning PP&E"]


   data = {"Category": categories}


   for i, year in enumerate(fiscal_years):
       is_estimated_year = year.endswith("E")
       data[year] = []
       for category in categories:
           col_letter = increment_letter('B', i)
           if category == "Beginning PP&E":
               formula = f"='Balance Sheet'!{col_letter}12" if not is_estimated_year else "Estimated formula for Beginning PP&E"
           elif category == "D&A":
               formula = f"='Cash Flow'!{col_letter}5" if not is_estimated_year else "Estimated formula for D&A"
           elif category == "CapEx":
               formula = f"={col_letter}6 - {col_letter}3 + {col_letter}4" if not is_estimated_year else "Estimated formula for CapEx"
           elif category == "Ending PP&E":
               col_letter = increment_letter('C', i)
               formula = f"='Balance Sheet'!{col_letter}12" if not is_estimated_year else "Estimated formula for Ending PP&E"
           elif category == "D&A as a % of Beginning PP&E":
               formula = f"={col_letter}4 / {col_letter}3" if not is_estimated_year else "Estimated formula for D&A as a % of Beginning PP&E"
           elif category == "CapEx as a % of Beginning PP&E":
               formula = f"={col_letter}5 / {col_letter}3" if not is_estimated_year else "Estimated formula for CapEx as a % of Beginning PP&E"
           else:
               formula = "..."  # Placeholder for other categories
           data[year].append(formula)
          
   # Additional logic for estimated fiscal years
   for i, year in enumerate(fiscal_years):
       if year.endswith("E"):  # Focus only on estimated years
           col_letter = increment_letter('B', i)  # Column for the current year
           prev_year_col_letter = increment_letter('C', i - 2)  # Column for the previous year's "Ending PP&E"


           for category in categories:
               if category == "Beginning PP&E":
                   # Reference the 'Ending PP&E' from the previous fiscal year in the same sheet
                   previous_year_ending_ppe_cell = f"{prev_year_col_letter}6"
                   data[year][categories.index(category)] = f"={previous_year_ending_ppe_cell}"


   # Calculate the average for 'D&A as a % of Beginning PP&E' for E fiscal years
   da_percentage_cells = []
   for i, year in enumerate(fiscal_years):
       if year.endswith("E"):
           col_letter = increment_letter('B', i - 5)
           da_percentage_cell = f"{col_letter}{row_number_for_DA_percentage}"
           da_percentage_cells.append(da_percentage_cell)


   # Formula to calculate the average
   if da_percentage_cells:
       average_formula = f"=AVERAGE({','.join(da_percentage_cells)})"


       # Apply this average to each 'E' fiscal year in 'D&A as a % of Beginning PP&E'
       for i, year in enumerate(fiscal_years):
           if year.endswith("E"):
               data[year][categories.index("D&A as a % of Beginning PP&E")] = average_formula


   # Logic for 'CapEx as a % of Beginning PP&E'
   capex_percentage_cells = []
   for i, year in enumerate(fiscal_years):
       if year.endswith("E"):
           col_letter = increment_letter('B', i - 5)
           capex_percentage_cell = f"{col_letter}{row_number_for_CapEx_percentage}"
           capex_percentage_cells.append(capex_percentage_cell)


   # Formula to calculate the average for CapEx
   if capex_percentage_cells:
       average_formula_capex = f"=AVERAGE({','.join(capex_percentage_cells)})"


       # Apply this average to each 'E' fiscal year in 'CapEx as a % of Beginning PP&E'
       for i, year in enumerate(fiscal_years):
           if year.endswith("E"):
               data[year][categories.index("CapEx as a % of Beginning PP&E")] = average_formula_capex
              
   # Logic for the D&A 'E' formula
   for i, year in enumerate(fiscal_years):
       if year.endswith("E"):  # Target only the 'E' fiscal years
           col_letter = increment_letter('B', i)  # Column for the current 'E' fiscal year
           value_in_row_3 = f"{col_letter}3"
           value_in_row_7 = f"{col_letter}7"


           # Formula to multiply the values in row 3 and row 7
           multiplication_formula = f"={value_in_row_3} * {value_in_row_7}"


           # Assign the formula to the 'D&A' category for the current 'E' fiscal year
           data[year][categories.index("D&A")] = multiplication_formula
          
   # Logic for the CapEx 'E' formula
   for i, year in enumerate(fiscal_years):
       if year.endswith("E"):  # Target only the 'E' fiscal years
           col_letter = increment_letter('B', i)  # Column for the current 'E' fiscal year
           value_in_row_3 = f"{col_letter}3"
           value_in_row_8 = f"{col_letter}8"


           # Formula to multiply the values in row 3 and row 8
           multiplication_formula = f"={value_in_row_3} * {value_in_row_8}"


           # Assign the formula to the 'CapEx' category for the current 'E' fiscal year
           data[year][categories.index("CapEx")] = multiplication_formula
          
   # Logic for the Ending PP&E 'E' formula
   for i, year in enumerate(fiscal_years):
       if year.endswith("E"):  # Target only the 'E' fiscal years
           col_letter = increment_letter('B', i)  # Column for the current 'E' fiscal year
           value_in_row_3 = f"{col_letter}3"
           value_in_row_4 = f"{col_letter}4"
           value_in_row_5 = f"{col_letter}5"


           # Formula for Ending PP&E: (Year 'E' Row 3) - (Year 'E' Row 4) + (Year 'E' Row 5)
           formula_ending_ppe = f"={value_in_row_3} - {value_in_row_4} + {value_in_row_5}"


           # Assign the formula to the 'Ending PP&E' category for the current 'E' fiscal year
           data[year][categories.index("Ending PP&E")] = formula_ending_ppe


                  
   df = pd.DataFrame(data)
   df.set_index('Category', inplace=True)
   return df


