import pandas as pd
import logging

logger = logging.getLogger(__name__)

try:
    df = pd.read_excel(r'C:\Users\benm\PycharmProjects\AutoCheckActive\Assets\Excels\Data.xlsx')
except Exception as e:
    logger.error(f"Failed to load or validate Excel file: {e}")
    raise SystemExit("Error: Could not process the Excel file.")

columns = ['Name', 'Customer', 'UDID', 'Version', 'Serial']
version = 808  

file_num = 0
row_count = 0
rows_to_add = []  
forbidden_customers = [
    "3466",
    "3208",
    "1742",
    "3456",
    "3589",
    "1099",
    "3244",
    "3124",
    "4446",
    "1248",
    "4423",
    "4598",
    "1433",
    "4497",
    "1580",
    "1351",
    "4807",
    "1679",
    "1946",
    "3448",
    "3283",
    "1438",
    "4923",
    "1852",
    "4318"
]

for num, row in df.iterrows():
    if str(row['Customer']) in forbidden_customers:
        continue
    elif str(row['serial']).startswith('EF500'):
        continue
    else:
        if int(row['ver']) < version and int(row['to_ver']) < version:
            new_row = {
                'Name': row['Name'],
                'Customer': row['Customer'], 
                'UDID': row['UDID'],           
                'Version': row['ver'],
                'Serial': row['serial']
            }
            rows_to_add.append(new_row)
            row_count += 1

        if row_count == 300:
            excel_file = pd.DataFrame(rows_to_add, columns=columns)
            output_file = f'C:\\Users\\benm\\PycharmProjects\\AutoCheckActive\\Assets\\Excels\\toUpdate\\output_{file_num}.xlsx'
            excel_file.to_excel(output_file, index=False)
            
            rows_to_add = []
            row_count = 0
            file_num += 1
if rows_to_add:
    excel_file = pd.DataFrame(rows_to_add, columns=columns)
    output_file = f'C:\\Users\\benm\\PycharmProjects\\AutoCheckActive\\Assets\\Excels\\toUpdate\\output_{file_num}.xlsx'
    excel_file.to_excel(output_file, index=False)
