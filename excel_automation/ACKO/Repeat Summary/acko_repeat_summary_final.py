import pandas as pd
import xlwings as xw
from win32com.client import Dispatch,win32

# File paths
acko_inbound_dump = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Repeat\dumps\Acko Inbound Summary Oct'25_output.xlsb"

#template files 
template_path = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Repeat\template\Repeat Summary Final.xlsb"
output_file_path = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Repeat\template\Repeat Summary Final_output.xlsb"

# Read the dump file into a DataFrame
call_detail_dump = pd.read_excel(acko_inbound_dump, sheet_name='Call Detail Dump')

# Remove unwanted columns from the dump DataFrame
#call_detail_dump = call_detail_dump.drop(['ASR agents', 'Interwal'], axis=1)

####### filtering out values where electronics = 1

call_detail_dump1 = call_detail_dump[
    (call_detail_dump['Electronics'] == 1) &
    (call_detail_dump['Call Type'] == 'Inbound') &
    (call_detail_dump['Status'] == 'Answered')
]

####### filtering out values where Internet = 1
call_detail_dump2 = call_detail_dump[
    (call_detail_dump['Internet'] == 1) &
    (call_detail_dump['Call Type'] == 'Inbound') &
    (call_detail_dump['Status'] == 'Answered')
]


call_detail_dump_health = call_detail_dump[
    (call_detail_dump['Health'] == 1) &
    (call_detail_dump['Call Type'] == 'Inbound') &
    (call_detail_dump['Status'] == 'Answered')
]
######## concatenating elec dump and internet dump

# Combining dataframes
print("Combining dataframes...")
combined_elec_inter = pd.concat([call_detail_dump1, call_detail_dump2], ignore_index=True)
print(f"Combined data: {len(combined_elec_inter)} rows")
############################################
# Using xlwings for pasting data into Excel
############################################

try:
    wb = xw.Book(template_path)
    
    # Access the sheet named 'Dump'
    Dump_sheet = wb.sheets['Dump']
    Hlth_Dump_sheet = wb.sheets['Dump Health']    
   
   
    Dump_sheet.range('H2').options(pd.DataFrame, index = False, header = False ,expand = 'table').value = combined_elec_inter
    Hlth_Dump_sheet.range('H2').options(pd.DataFrame, index = False, header = False ,expand = 'table').value = call_detail_dump_health
       
    # Save the workbook with the new output file path and close it.
    wb.save(output_file_path)
    wb.close()

    print(f"Data successfully written to {output_file_path}")
    
except Exception as e:
    print(f"Error while saving data to Excel using xlwings: {e}")
    
############################################
### using pywin32 to drag formulas
excel = Dispatch("Excel.Application")
excel.Visible = True

source = excel.Workbooks.Open(output_file_path)
 
ws1 = source.Worksheets("Dump")
ws2 = source.Worksheets("Dump Health")

try:
    ws1.Activate()
    
    le1 = str(len(combined_elec_inter) + 1)    
    #fill down    
    formula_range1 = ws1.Range("A2:G2")    
    destination_range1 = ws1.Range(f"A2:G{le1}")  # Define where to fill formulas   
    destination_range1.FillDown()  # Copies formulas down
    
    destination_range2 = ws1.Range(f"CE2:CE{le1}")  # Define where to fill formulas   
    destination_range2.FillDown()  # Copies formulas down
    
    # Define the full data range for sorting
    sort_range = ws1.Range(f"A1:AA{le1}")  # Adjust 'AL' if your sheet has more/less columns
    key1 = ws1.Range(f"AM2:AM{le1}")  # 'Cleaned Nos'
    key2 = ws1.Range(f"CE2:CE{le1}")  # 'Date'
    key3 = ws1.Range(f"I2:I{le1}")  # 'Date'

    # Perform sort
    ws1.Sort.SortFields.Clear()
    ws1.Sort.SortFields.Add(
        Key=key1,
        SortOn=win32.constants.xlSortOnValues,
        Order=win32.constants.xlAscending,
        DataOption=win32.constants.xlSortNormal
    )
    ws1.Sort.SortFields.Add(
        Key=key2,
        SortOn=win32.constants.xlSortOnValues,
        Order=win32.constants.xlAscending,
        DataOption=win32.constants.xlSortNormal
    )
    ws1.Sort.SortFields.Add(
        Key=key3,
        SortOn=win32.constants.xlSortOnValues,
        Order=win32.constants.xlAscending,
        DataOption=win32.constants.xlSortNormal
    )

    ws1.Sort.SetRange(sort_range)
    ws1.Sort.Header = win32.constants.xlYes
    ws1.Sort.Apply()

    
    ws2.Activate()
    
    le1 = str(len(call_detail_dump_health) + 1)    
    #fill down    
    formula_range2 = ws2.Range("A2:G2")    
    destination_range2 = ws2.Range(f"A2:G{le1}")  # Define where to fill formulas   
    destination_range2.FillDown()  # Copies formulas down
    
    destination_range2 = ws1.Range(f"CE2:CE{le1}")  # Define where to fill formulas   
    destination_range2.FillDown()  # Copies formulas down
    
    # Define the full data range for sorting
    sort_range = ws1.Range(f"A1:AA{le1}")  # Adjust 'AL' if your sheet has more/less columns
    key1 = ws2.Range(f"AM2:AM{le1}")  # 'Cleaned Nos'
    key2 = ws2.Range(f"CE2:CE{le1}")  # 'Date'
    key3 = ws2.Range(f"I2:I{le1}")  # 'Date'

    # Perform sort
    ws2.Sort.SortFields.Clear()
    ws2.Sort.SortFields.Add(
        Key=key1,
        SortOn=win32.constants.xlSortOnValues,
        Order=win32.constants.xlAscending,
        DataOption=win32.constants.xlSortNormal
    )
    ws2.Sort.SortFields.Add(
        Key=key2,
        SortOn=win32.constants.xlSortOnValues,
        Order=win32.constants.xlAscending,
        DataOption=win32.constants.xlSortNormal
    )
    ws2.Sort.SortFields.Add(
        Key=key3,
        SortOn=win32.constants.xlSortOnValues,
        Order=win32.constants.xlAscending,
        DataOption=win32.constants.xlSortNormal
    )

    ws2.Sort.SetRange(sort_range)
    ws2.Sort.Header = win32.constants.xlYes
    ws2.Sort.Apply()
    
except Exception as e:
    print(f"Error while applying formulas on sheet using pywin32: {e}")
# Save and Close
source.Save()
source.Close()
excel.Quit()







 

