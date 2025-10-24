# -*- coding: utf-8 -*-
"""
Created on Thu Oct  9 16:55:01 2025

@author: JASKIRAT
"""

import pandas as pd
from datetime import timedelta, datetime
import xlwings as xw

####################################################################################
#DUMPS
####################################################################################

df = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\SWIGGY\BDV\08-10-2025.xlsb",sheet_name="Layer")

####################################################################################
#TEMPLATE
####################################################################################

template = r"C:\Users\8242K\Desktop\WFM\SWIGGY\BDV\Swiggy Store BDV Report.xlsx"
template_output = r"C:\Users\8242K\Desktop\WFM\SWIGGY\BDV\Swiggy Store BDV Report_OUTPUT.xlsx"

####################################################################################
#MODIFICATIONS
####################################################################################

df = df[['NODELABEL','ORDER_DATE','EFFORTSCORE']]

# ORDER_DATE is int64
base_date = datetime(1899, 12, 30)
df['ORDER_DATE'] = df['ORDER_DATE'].apply(lambda x: base_date + timedelta(days=int(x)))

df['ORDER_DATE'] = pd.to_datetime(df['ORDER_DATE'],format('%m/%d/%Y'))
# df['ORDER_DATE'] = df['ORDER_DATE'].dt.strftime('%m/%d/%Y')

df['C SAT'] = df['EFFORTSCORE'].apply(lambda x : 1 if x > 2 else 0)

df['D SAT'] = df['EFFORTSCORE'].apply(lambda x : 1 if x < 3 else 0)


# Group by NODELABEL and ORDER_DATE for per-date counts and sums
pivot_df = df.groupby(['NODELABEL', 'ORDER_DATE'], dropna=False).agg(
    TOTAL_ENTRIES=('NODELABEL', 'count'),                               # Count all rows per node per date
    C_SAT_SUM=('C SAT', 'sum'),                                         # Sum of C SAT per date
    D_SAT_SUM=('D SAT', 'sum')                                          # Sum of D SAT per date
).reset_index()
pivot_df['TOTAL SUM'] = pivot_df['C_SAT_SUM']+pivot_df['D_SAT_SUM']

# Export to Excel if needed
pivot_df.to_excel(r"C:\Users\8242K\Desktop\WFM\SWIGGY\BDV\pivot_table.xlsx", index=False)

####################################################################################
#WRITING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(template)
    
    sheet = wb.sheets['Sheet1']
    
    has_error = False  
    
    # SHEET 1 ------------------------------------------------------------
    try:
        last_row1 = sheet.range('B' + str(sheet.cells.last_cell.row)).end('up').row
        next_row = last_row1 + 1
        sheet[f"B{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = pivot_df
        
        print("✅ SUCCESS: Data successfully written to 'SHEET 1'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'SHEET 1': {e} \n")
        
    wb.save(template_output)
    wb.close()
    
    if has_error:
        print(f"⚠️ COMPLETED WITH ERRORS: Some sheets failed to update in {template_output}\n")
    else:
        print(f"😄 ALL GOOD: Excel update completed without errors at {template_output}\n")

except Exception as e:
    print(f"🔥 FAILURE: Could not save workbook {template_output}: {e}")
finally:
    app.quit()
    

