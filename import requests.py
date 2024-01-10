import requests
import json
import pandas as pd
import numpy as np

url = "https://backend.simfin.com/api/v3/companies/statements/compact?ticker=AMZN&statements=PL,BS,CF&period=FY&fyear=2017%2C2018%2C2019%2C2020%2C2021%2C2022"

headers = {
    "accept": "application/json",
    "Authorization": "b1c124d9-f078-4887-abb9-d3504b54b23b"
}

response = requests.get(url, headers=headers)

jsonStr = json.loads(response.text)
data = jsonStr[0]['statements']
arr = np.empty(len(data), dtype=object)

for i in range(len(data)):    
    columns = data[i]['columns']
    # Extract data
    table_data = data[i]['data']
    # Create a DataFrame
    arr[i] = pd.DataFrame(table_data, columns=columns)

# Convert DataFrame to Excel
excel_file_path = 'output.xlsx'
#arr[0].to_excel(excel_file_path, index=False)

with pd.ExcelWriter('financial_data_separate_rows.xlsx') as writer:
    arr[0].to_excel(writer, index=False, sheet_name='sheet1')
    arr[1].to_excel(writer, index=False, sheet_name='sheet2')
    arr[2].to_excel(writer, index=False, sheet_name='sheet3')



#jsonStr = json.dumps(jsonArr)
#df = pd.json_normalize(jsonStr)
#df.to_csv("flattened_output.csv", index=False)