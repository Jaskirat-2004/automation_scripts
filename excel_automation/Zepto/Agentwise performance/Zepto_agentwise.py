# -*- coding: utf-8 -*-
"""
Created on Fri Nov  7 13:41:27 2025

@author: 8242k
"""

import pandas as pd
import xlwings as xw
from datetime import date ,timedelta

####################################################################################
#DUMPS
####################################################################################

shift_wise = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\Zepto\Agentwise performance\Dump\Shift Wise Login Hours.xlsb",sheet_name="Shift Plotters & Login",skiprows = 1)
ticket_dump = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\Zepto\Agentwise performance\Dump\Ticket Dump - Nov'25.xlsx",sheet_name="For Agent Wise")
shift_adherence = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\Zepto\Agentwise performance\Dump\Zepto Shift Adherence  Nov'2025 Updated_Output.xlsb",sheet_name = "Raw",skiprows=1)
csat = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\Zepto\Agentwise performance\Dump\C-Sat Performance Report - Nov'25.xlsb",sheet_name="C-Sat Raw Data")

####################################################################################
#TEMPLATE
####################################################################################

template = r"C:\Users\8242K\Desktop\WFM\Zepto\Agentwise performance\template\Zepto Agent Wise Performance  Report - Sep'25.xlsb"
template_output = r"C:\Users\8242K\Desktop\WFM\Zepto\Agentwise performance\template\Zepto Agent Wise Performance  Report - Sep'25_OUTPUT.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

# DATE
today = date.today()
yesterday = today-timedelta(days=1)
yesterday_date = yesterday.strftime("%Y-%m-%d")

# SHIFT WISE DUMP ----------------------------------------------------------------------------

shift_wise = shift_wise.iloc[:,:14]
cols = ['Date Wise','Overall Login','Total Login Within Shift']
shift_wise['Date Wise'] = pd.to_datetime(shift_wise['Date Wise'],origin='1899-12-30', unit='D')

# TICKET DUMP ----------------------------------------------------------------------------

ticket_dump = ticket_dump[ticket_dump['Date for Agent Wise'] == yesterday_date]
ticket_dump['Date for Agent Wise'] = ticket_dump['Date for Agent Wise'].astype(str).str.split().str[0]

cols = [' AHT', ' FRS', ' Wait Time',' Wait Time+FRS']

for col in cols:
    ticket_dump[col] = ticket_dump[col].astype(str).str.extract(r'(\d{2}:\d{2}:\d{2})')[0]
    
# SHIFT ADHERENECE DUMP ----------------------------------------------------------------------------

shift_adherence = shift_adherence.iloc[:,:39]

####################################################################################
#WRITING DATA
####################################################################################

try:
    app = xw.App()
    wb = app.books.open(template)
        
    has_error = False  
    
    sheet1 = wb.sheets["APR"] 
    sheet2 = wb.sheets["Ticket Raw"] 
    sheet3 = wb.sheets["Shift Adh%"] 
    sheet4 = wb.sheets["Csat Raw"] 

    # APR SHEET ------------------------------------------------------------
    
    try:
        
        sheet1["A2"].options(pd.DataFrame,header=False,index=False, expand="table").value = shift_wise

        print("✅ SUCCESS: Data successfully written to 'APR SHEET'")
        
        drag = len(shift_wise)

        sheet1.range("O2:AL2").api.AutoFill(Destination=sheet1.range(f"O2:AL{drag}").api)
        
        print(f"🎯 Formulas successfully applied to range O2:AL{drag}")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'APR SHEET': {e} \n")
        
    # TICKET RAW SHEET ------------------------------------------------------------
        
    try:
        last_row = sheet2.range('A' + str(sheet2.cells.last_cell.row)).end('up').row
        next_row = last_row + 1
        
        sheet2[f"A{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = ticket_dump

        print("✅ SUCCESS: Data successfully written to 'TICKET RAW SHEET'")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'TICKET RAW SHEET': {e} \n")
            
    # SHIFT% SHEET ------------------------------------------------------------
    
    try:
        sheet3["A2"].options(pd.DataFrame,header=False,index=False, expand="table").value = shift_adherence

        print("✅ SUCCESS: Data successfully written to 'SHIFT% SHEET'")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'SHIFT% SHEET': {e} \n")
    
    # CSAT RAW SHEET ------------------------------------------------------------
        
    try:
        
        sheet4["C2"].options(pd.DataFrame,header=False,index=False, expand="table").value = csat

        print("✅ SUCCESS: Data successfully written to 'CSAT SHEET'")
        
        drag = sheet4.range('C' + str(sheet4.cells.last_cell.row)).end('up').row

        sheet4.range("A2:B2").api.AutoFill(Destination=sheet4.range(f"A2:B{drag}").api)

        print(f"🎯 Formulas successfully applied to range A2:B{drag}")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'CSAT SHEET': {e} \n")

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
    
    
    
