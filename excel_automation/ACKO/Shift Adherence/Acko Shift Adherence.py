# -*- coding: utf-8 -*-
"""
Created on Fri Oct 17 18:14:16 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw
from datetime import date,timedelta
from win32com.client import Dispatch
import numpy as np

####################################################################################
#DUMPS
####################################################################################

agent_details_path = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Shift Adherence\Dumps\Acko Agent Hygiene Report.xlsb"
####################################################################################
#TEMPLATE
####################################################################################

template = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Shift Adherence\Template\Acko Shift Adherence Oct'25.xlsx"
template_output = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Shift Adherence\Template\Acko Shift Adherence Oct'25_OUTPUT.xlsx"

####################################################################################
#COPYING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(agent_details_path,password="Acko@2024")
    
    agent_details_sheet = wb.sheets['Agent Details']
    
    # AGENT DETAILS SHEET ------------------------------------------------------------
    try:
        agent_details_dump = agent_details_sheet["B1"].options(pd.DataFrame,header=True,index=False, expand="table").value

        print("✅ SUCCESS: Data successfully copied from to 'AGENT DETAILS SHEET'")
        
    except Exception as e:
        print(f"❌ ERROR: Failed to copy from 'AGENT DETAILS SHEET': {e} \n")
        
    wb.close()

except Exception as e:
    print(f"🔥 FAILURE: Could not close workbook : {e}")
finally:
    app.quit()


####################################################################################
#MODIFICATIONS
####################################################################################

# DATE
yesterday = date.today() - timedelta(days=35)
yesterday_date = yesterday.strftime("%Y-%m-%d")

# AGENT DETAILS DUMP
agent_details_dump["Date"] = pd.to_datetime(agent_details_dump["Date"], errors="coerce").dt.strftime("%Y-%m-%d")
# agent_details_dump["FHD"] = pd.to_datetime(df["FHD"], errors="coerce").dt.strftime("%Y-%m-%d")
     
df = agent_details_dump[agent_details_dump["Date"] == yesterday_date]

df.iloc[:,53:58] = np.nan

col_to_move = df.columns[58]
col_data = df.pop(col_to_move)
df.insert(53, col_to_move, col_data)

df = df.drop(" Neutral",axis = 1)

####################################################################################
#WRITING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(template)
    
    raw_sheet = wb.sheets["Raw"]    

    has_error = False  
    
    # IVR SHEET ------------------------------------------------------------
    try:
        last_row1 = raw_sheet.range('B' + str(raw_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row1 + 1
        raw_sheet[f"B{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = df
        
        print("✅ SUCCESS: Data successfully written to 'RAW SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'RAW SHEET': {e} \n")
        
        
    wb.save(template_output)
    # wb.save(template)
    wb.close()
    
    if has_error:
        print("⚠️ COMPLETED WITH ERRORS: Some sheets failed to update ! \n")
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

ws1 = wb.Worksheets("Raw")

try:
    # -------------------------------------------------------------------------------------------
    # IVR Dump
    
    ws1.Activate()
    
    last_row = last_row1+len(df)
    
    formula_range1 = ws1.Range(f"A{last_row1}:A{last_row}")
    formula_range2 = ws1.Range(f"BC{last_row1}:BG{last_row}") 
    formula_range1.FillDown()
    formula_range2.FillDown()
    
    print(f"🎯 Formulas successfully applied to range A{last_row1}:A{last_row} and BC{last_row1}:BG{last_row}")

    wb.Save()
    # wb.SaveAs(template)
    wb.Close()
    
    print("🚀💥 FORMULAS SUCCESSFULLY APPLIED TO ALL SHEETS USING FILLDOWN! MASTERED BY JASKIRAT! 🎯✅")

except Exception as e:
    print(f"❌ ERROR: {e}")
finally:
    excel.Quit()
    
