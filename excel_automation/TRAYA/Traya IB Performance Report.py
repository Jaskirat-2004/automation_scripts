# -*- coding: utf-8 -*-
"""
Created on Tue Oct  7 12:50:03 2025

@author: JASKIRAT

"""

import pandas as pd
import xlwings as xw
from datetime import date,timedelta
from win32com.client import Dispatch

####################################################################################
#DUMPS
####################################################################################

acd_dump = pd.read_csv(r"C:\Users\8242K\Desktop\WFM\TRAYA\Inbound Performance Report\Dumps\ACD_Call_Details.csv")
apr_dump = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\TRAYA\Inbound Performance Report\Dumps\APR.xlsb")

####################################################################################
#TEMPLATE
####################################################################################

template = r"C:\Users\8242K\Desktop\WFM\TRAYA\Inbound Performance Report\Template\Traya Inbound Performance Report- Oct'25.xlsb"
template_output = r"C:\Users\8242K\Desktop\WFM\TRAYA\Inbound Performance Report\Template\Traya Inbound Performance Report- Oct'25_OUTPUT.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

# DATE
yesterday = date.today()-timedelta(days=1)
yesterday_date = yesterday.strftime("%Y-%m-%d")

# ACD DUMP

acd_dump = acd_dump[acd_dump['User ID'].str.endswith('maxicus.com',na=False)]

# APR DUMP 

apr_dump.insert(2,"BLANK1","")
apr_dump.insert(3,"BLANK2","")
apr_dump = apr_dump.iloc[:,1:]

####################################################################################
#WRITING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(template,password="Traya@123")
    # wb = app.books.open(temp)
        
    # s1 = wb.sheets['Sheet1']
    # s2 = wb.sheets['Sheet2']
    
    acd_dump_sheet = wb.sheets['ACD Dump']
    apr_dump_sheet = wb.sheets['APR Dump']
    
    has_error = False  
    
    # ACD SHEET ------------------------------------------------------------
    try:
        last_row1 = acd_dump_sheet.range('M' + str(acd_dump_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row1 + 1
        acd_dump_sheet[f"M{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = acd_dump
        
        print("✅ SUCCESS: Data successfully written to 'ACD DUMP SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'ACD DUMP SHEET': {e} \n")
        
    # APR SHEET ------------------------------------------------------------
    try:
        last_row2 = apr_dump_sheet.range('C' + str(apr_dump_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row2 + 1
        apr_dump_sheet[f"C{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = apr_dump
        
        print("✅ SUCCESS: Data successfully written to 'APR DUMP SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'APR DUMP SHEET': {e} \n")
        
    wb.save(template_output)
    # wb.save(temp)
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
wb = excel.Workbooks.Open(template_output,Password="Traya@123")
# wb = excel.Workbooks.Open(template_output)

# ws1 = wb.Worksheets("Sheet1")
# ws2 = wb.Worksheets("Sheet2")

ws1 = wb.Worksheets("ACD Dump")
ws2 = wb.Worksheets("APR Dump")
ws3 = wb.Worksheets("Summary")


try:
    # -------------------------------------------------------------------------------------------
    # ACD DUMP SHEET
    
    ws1.Activate()
    
    last_row = ws1.Cells(ws1.Rows.Count, 13).End(-4162).Row 
    
    formula_range1 = ws1.Range(f"A{last_row1}:L{last_row}")  # Expanding A{last_row1}:F down to the last used row
    formula_range1.FillDown()
    
    print(f"🎯 Formulas successfully applied to range A{last_row1}:F{last_row}")
    
    # -------------------------------------------------------------------------------------------
    # APR DUMP SHEET
    
    ws2.Activate()
    
    last_row = ws2.Cells(ws2.Rows.Count, 6).End(-4162).Row
    
    formula_range1 = ws2.Range(f"A{last_row2}:B{last_row}")  # Expanding A{last_row2}:B down to the last used row
    formula_range2 = ws2.Range(f"D{last_row2}:E{last_row}")  # Expanding H{last_row2}:I down to the last used row
    formula_range1.FillDown()
    formula_range2.FillDown()
    

    print(f"🎯 Formulas successfully applied to range A{last_row2}:B{last_row} & H2:I{last_row}")
    
    # -------------------------------------------------------------------------------------------
    # SUMMARY SHEET
    
    ws3.Activate()

    try:
        found = False
    
        for col in range(1, 40):
            cell_value = ws3.Cells(2, col).Value  # may be datetime or None
    
            if not cell_value:
                continue
            
            if str(yesterday) in str(cell_value):
                ws3.Columns(col).Hidden = False
                print(f"✅ Unhid column {col} ({cell_value}) in 'Summary'")
                found = True
                break

            # if isinstance(cell_value, (datetime, date)):
            #     # convert both to pure date
            #     if cell_value.date() == yesterday.date():
            #         ws3.Columns(col).Hidden = False
            #         print(f"✅ Unhid column {col} ({cell_value}) in 'Summary'")
            #         found = True
            #         break
    
        if not found:
            print(f"⚠️ No column found for yesterday ({yesterday}) in 'Summary' sheet")
    
    except Exception as e:
        print(f"❌ ERROR while unhiding Summary column: {e}")


    wb.Save()
    wb.Close()
    
    print("🚀💥 FORMULAS SUCCESSFULLY APPLIED TO ALL SHEETS USING FILLDOWN! MASTERED BY JASKIRAT! 🎯✅")

except Exception as e:
    print(f"❌ ERROR: {e}")
    
finally:
    excel.Quit()
    
    
