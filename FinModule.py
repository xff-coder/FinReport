
import numpy as np
import requests
import json
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, NamedStyle, Border, Side, Color
from openpyxl.styles.numbers import FORMAT_CURRENCY_USD_SIMPLE
from openpyxl.utils import get_column_letter
 
# Function to process, filter, and format data for a given statement
def process_statement(data, columns, selected_columns):
    df = pd.DataFrame(data, columns=columns)
    df.set_index('Fiscal Year', inplace=True)
    df_filtered = df[selected_columns]
    # Convert to millions (without formatting as string)
    for col in df_filtered.columns:
        df_filtered[col] = df_filtered[col] / 1e6
    return df_filtered.T