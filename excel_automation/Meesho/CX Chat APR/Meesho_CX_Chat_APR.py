# -*- coding: utf-8 -*-
"""
Created on Mon Nov  3 16:21:48 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw

####################################################################################
#DUMPS
####################################################################################

sheet = "31-10"

####################################################################################
#DUMPS
####################################################################################

master_report  = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\Meesho\CX Chat APR\Dump\Master Report-Kochar.xlsx",sheet_name=sheet)
attendence  = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\Meesho\CX Chat APR\Dump\Daily Attendace Cx Chat.xlsx",sheet_name="OverAll ")

####################################################################################
#TEMPLATE
####################################################################################

template = r"C:\Users\8242K\Desktop\WFM\Meesho\CX Chat APR\Temaplate\Meesho CX APR Report Oct'25.xlsx"
template_output = r"C:\Users\8242K\Desktop\WFM\Meesho\CX Chat APR\Temaplate\Meesho CX APR Report Oct'25_OUTPUT.xlsx"

####################################################################################
#MODIFICATIONS
####################################################################################

# MASTER REPORT DUMP ----------------------------------------------------------------------------

master_report = master_report[
    ['RoomCode', 'RoomUrl', 'RoomStatus', 'ChatStartTime', 'ChatEndTime', 'CsatScore', 'ClosedBy', 'AgentStats', 
     'AgentEmail', 'FirstAgentFirstResponseTime', 'AgentAssignmentTimestamp', 'AverageAgentResponseTime', 'L1', 'L2']
    ]

# Convert all datetime.time columns to string (HH:MM:SS)
time_cols = ['FirstAgentFirstResponseTime', 'AverageAgentResponseTime']

for col in time_cols:
    if col in master_report.columns:
        master_report[col] = master_report[col].astype(str)


# ATTENDENCE DUMP ----------------------------------------------------------------------------

attendence = attendence[
    ['Date', 'EmpId','Meesho Email ID', 'Emp Name','TL Name', 'Scheduled Shift','Attendance']
    ]

attendence.insert(5,"BLANK1","")
attendence.insert(5,"BLANK2","")

####################################################################################
#WRITING DATA
####################################################################################

try:
    app = xw.App()
    wb = app.books.open(template)
        
    has_error = False  
    
    sheet1 = wb.sheets["Raw"] 
    sheet2 = wb.sheets["Performance"] 

    # RAW SHEET ------------------------------------------------------------
    
    try:
        last_row = sheet1.range('I' + str(sheet1.cells.last_cell.row)).end('up').row
        next_row = last_row + 1
        
        sheet1[f"I{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = master_report

        print("✅ SUCCESS: Data successfully written to 'RAW SHEET'")
        
        drag = sheet1.range('I' + str(sheet1.cells.last_cell.row)).end('up').row

        sheet1.range(f"A{last_row}:H{last_row}").api.AutoFill(Destination=sheet1.range(f"A{last_row}:H{drag}").api)
        sheet1.range(f"W{last_row}:Y{last_row}").api.AutoFill(Destination=sheet1.range(f"W{last_row}:Y{drag}").api)

        print(f"🎯 Formulas successfully applied to range A{last_row}:H{drag} & W{last_row}:Y{drag}")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'RAW SHEET': {e} \n")
        
    # PERFORMANCE SHEET ------------------------------------------------------------
        
    try:
        
        sheet2["B2"].options(pd.DataFrame,header=False,index=False, expand="table").value = attendence

        print("✅ SUCCESS: Data successfully written to 'PERFORMANCE SHEET'")
        
        drag = sheet1.range('B' + str(sheet1.cells.last_cell.row)).end('up').row

        sheet2.range("A2:A2").api.AutoFill(Destination=sheet2.range(f"A2:A{drag}").api)
        sheet2.range("K2:M2").api.AutoFill(Destination=sheet2.range(f"K2:M{drag}").api)

        print(f"🎯 Formulas successfully applied to range A2:A{drag} & K2:M{drag}")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'PERFORMANCE SHEET': {e} \n")

    wb.save(template_output)
    wb.close()
    
    if has_error:
        print("⚠️ COMPLETED WITH ERRORS: Some sheets failed to update \n")
    else:
        print(f"😄 ALL GOOD: Excel update completed without errors at {template}\n")
    
    print("🚀💥 MASTERED BY JASKIRAT! 🎯✅")

except Exception as e:
    print(f"🔥 FAILURE: Could not save workbook: {e}")
finally:
    app.quit()
    
    
    
