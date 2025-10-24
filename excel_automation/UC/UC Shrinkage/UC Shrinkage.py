# -*- coding: utf-8 -*-
"""
Created on Wed Oct 15 17:24:23 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw
from datetime import date,timedelta
from win32com.client import Dispatch

####################################################################################
#DUMPS
####################################################################################

hc_dump = pd.read_csv(r"C:\Users\8242K\Desktop\WFM\UC\UC Shrinkage\Dump\head_count_job.csv")

####################################################################################
#TEMPLATE
####################################################################################

template = r"C:\Users\8242K\Desktop\WFM\UC\UC Shrinkage\Template\Urban Company Shrinkage & FTE Oct'25.xlsb"
template_output = r"C:\Users\8242K\Desktop\WFM\UC\UC Shrinkage\Template\Urban Company Shrinkage & FTE Oct'25_OUTPUT.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

# DATE
yesterday = (date.today()-timedelta(days=1))
yesterday_date = yesterday.strftime("%d/%b/%y")

# HC DUMP

hc_dump = hc_dump[hc_dump['OU Name'] == 'Urban company Phygital']
hc_dump = hc_dump[hc_dump['Grade'].isin(['G1','G2'])]
hc_dump = hc_dump[hc_dump['Department Name'] == 'Operations']
hc_dump = hc_dump[hc_dump['Current Status'] == 'Active']

hc_dump = hc_dump[['Employee ID','Employee Name','OU Name','Location Name','Designation Name','Department Name',
                  'Reporting To Name','Functional Reporting To Name','Current Status','Workplace Category']]

hc_dump.insert(0,"DATE",yesterday_date)
hc_dump.insert(9,"BLANK","")
hc_dump.insert(11,"BLANK1","")

####################################################################################
#WRITING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(template)
    
    hc_dump_sheet = wb.sheets['HC Dump']
    
    has_error = False  
        
    # HC DUMP SHEET ------------------------------------------------------------
    try:
        last_row1 = hc_dump_sheet.range('C' + str(hc_dump_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row1 + 1
        hc_dump_sheet[f"C{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = hc_dump
        
        print("✅ SUCCESS: Data successfully written to 'HC DUMP SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'HC DUMP SHEET': {e} \n")
          
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
    
####################################################################################
# DRAGING FORMULA
####################################################################################

excel = Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open(template_output)

ws1 = wb.Worksheets("HC Dump")

try:
    # -------------------------------------------------------------------------------------------
    # HC DUMP SHEET
    
    ws1.Activate()
    last_row = ws1.Cells(ws1.Rows.Count, 3).End(-4162).Row 

    formula_range1 = ws1.Range(f"A{last_row1}:B{last_row}")
    formula_range2 = ws1.Range(f"L{last_row1}:L{last_row}")
    formula_range3 = ws1.Range(f"N{last_row1}:N{last_row}")
    formula_range4 = ws1.Range(f"P{last_row1}:V{last_row}")
    
    formula_range1.FillDown()
    formula_range2.FillDown()
    formula_range3.FillDown()
    formula_range4.FillDown()
    
    print(f"🎯 Formulas successfully applied to range A2:A{last_row} & AA2:AB{last_row}")
    
    wb.Save()
    wb.Close()
    
    print("🚀💥 FORMULAS SUCCESSFULLY APPLIED TO ALL SHEETS USING FILLDOWN! MASTERED BY JASKIRAT! 🎯✅")

except Exception as e:
    print(f"❌ ERROR: {e}")
    
finally:
    excel.Quit()
    
    
