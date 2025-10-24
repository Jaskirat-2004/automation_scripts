# -*- coding: utf-8 -*-
"""
Created on Wed Oct 22 18:24:24 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw
from datetime import date,timedelta
from win32com.client import Dispatch

####################################################################################
#DUMPS
####################################################################################

hc_dump = pd.read_csv(r"C:\Users\8242K\Desktop\WFM\Boat\Boat IB Shrinkage\Dump\head_count_job.csv")
agent_availability_dump = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\Boat\Boat IB Shrinkage\Dump\AgentAvailabiltyReport.xlsx",skiprows=2)

####################################################################################
#TEMPLATE
####################################################################################

template = r"C:\Users\8242K\Desktop\WFM\Boat\Boat IB Shrinkage\Template\Boat Voice Shrinkage Oct'25.xlsb"
template_output = r"C:\Users\8242K\Desktop\WFM\Boat\Boat IB Shrinkage\Template\Boat Voice Shrinkage Oct'25_OUTPUT.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

# DATE
yesterday = (date.today()-timedelta(days=1))
yesterday_date = yesterday.strftime("%Y-%m-%d")

# HC DUMP
hc_dump = hc_dump[hc_dump["OU Name"] == "Boat IB"]
hc_dump = hc_dump[hc_dump["Grade"] == "G1"]
hc_dump = hc_dump[hc_dump["Department Name"] == "Operations"]
hc_dump = hc_dump[hc_dump["Current Status"] == "Active"]

hc_dump = hc_dump[['Employee ID', 'Employee Name', 'OU Name', 'Location Name', 'Designation Name', 'Grade',
                   'Department Name', 'Reporting To Name', 'Functional Reporting To Name', 'Date Of Joining',
                   'Job Category', 'Current Status',  'Batch Number', 'Workplace Category']]

hc_dump.insert(9,"Blank","")
hc_dump.insert(0,"Date",yesterday_date)

# AGENT AVAILABILTY DUMP
agent_availability_dump["Agent Email Id"] = agent_availability_dump["Agent Email Id"].str.split("@").str[0]
agent_availability_dump["Agent Email Id"] = agent_availability_dump["Agent Email Id"].str.replace(r"\.", " ", regex=True)

agent_availability_dump.insert(0,"blank","")
agent_availability_dump.insert(0,"Date",yesterday_date)

agent_availability_dump["Time in Status (Selected Date) (Custom) (SUM)"] = agent_availability_dump["Time in Status (Selected Date) (Custom) (SUM)"].apply(lambda x: str(x).split()[-1].split('.')[0])


####################################################################################
#WRITING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(template)
    
    hc_sheet = wb.sheets["HC"]    
    sprinkler_sheet = wb.sheets["Sprinkler Login Hours"]

    has_error = False  
    
    # HC SHEET ------------------------------------------------------------
    try:
        last_row1 = hc_sheet.range('D' + str(hc_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row1 + 1
        hc_sheet[f"D{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = hc_dump
        
        print("✅ SUCCESS: Data successfully written to 'HC SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'HC SHEET': {e} \n")
        
    # SPRINKLER SHEET ------------------------------------------------------------
    try:
        last_row2 = sprinkler_sheet.range('B' + str(sprinkler_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row2 + 1
        sprinkler_sheet[f"B{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = agent_availability_dump
        
        print("✅ SUCCESS: Data successfully written to 'SPRINKLER SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'SPRINKLER SHEET': {e} \n")

    wb.save(template_output)
    # wb.save(template)
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

ws1 = wb.Worksheets("HC")
ws2 = wb.Worksheets("Sprinkler Login Hours")


try:
    # -------------------------------------------------------------------------------------------
    # HC
    
    ws1.Activate()
    
    last_row = ws1.Cells(ws1.Rows.Count, 4).End(-4162).Row 
    
    formula_range1 = ws1.Range(f"A{last_row1}:C{last_row}")
    formula_range2 = ws1.Range(f"N{last_row1}:N{last_row}")
    formula_range3 = ws1.Range(f"T{last_row1}:AN{last_row}")
    formula_range1.FillDown()
    formula_range2.FillDown()
    formula_range3.FillDown()
    
    print(f"🎯 Formulas successfully applied to range A{last_row1}:C{last_row} & N{last_row1}:N{last_row} & T{last_row1}:AN{last_row}")
    
    # -------------------------------------------------------------------------------------------
    # SPRINKLER LOGIN HOURS
    
    ws2.Activate()
    
    last_row = ws2.Cells(ws2.Rows.Count, 5).End(-4162).Row
    
    formula_range1 = ws2.Range(f"A{last_row2}:A{last_row}")
    formula_range2 = ws2.Range(f"C{last_row2}:C{last_row}")
    formula_range3 = ws2.Range(f"G{last_row2}:G{last_row}")
    formula_range1.FillDown()
    formula_range2.FillDown()
    formula_range3.FillDown()
    
    print(f"🎯 Formulas successfully applied to range A{last_row2}:A{last_row} & C{last_row2}:C{last_row} & G{last_row2}:G{last_row}")
    
    # -------------------------------------------------------------------------------------------
    
    wb.Save()
    wb.SaveAs(template)

    print("🚀💥 FORMULAS SUCCESSFULLY APPLIED TO ALL SHEETS USING FILLDOWN! MASTERED BY JASKIRAT! 🎯✅")

except Exception as e:
    print(f"❌ ERROR: {e}")
finally:
    wb.Close()
    excel.Quit()
    
