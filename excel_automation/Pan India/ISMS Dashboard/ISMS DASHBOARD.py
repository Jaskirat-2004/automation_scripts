# -*- coding: utf-8 -*-
"""
Created on Mon Oct 13 15:09:41 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw
from datetime import date,timedelta
from win32com.client import Dispatch

####################################################################################
#DUMPS
####################################################################################

information_dump = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\Pan India\ISMS Dashboard\Dump\Information Security Awareness.xlsx")
hc_dump = pd.read_csv(r"C:\Users\8242K\Desktop\WFM\Pan India\ISMS Dashboard\Dump\head_count_job.csv")

####################################################################################
#TEMPLATE
####################################################################################

template = r"C:\Users\8242K\Desktop\WFM\Pan India\ISMS Dashboard\Template\ISMS Coverage  report.xlsb"
template_output = r"C:\Users\8242K\Desktop\WFM\Pan India\ISMS Dashboard\Template\ISMS Coverage  report_OUTPUT.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

# DATE
today = date.today()
date_list = []
for i in range(1,7):
    date_list.append(str(today-timedelta(days=i)))
    
today = today.strftime("%Y-%m-%d")

# INFORMATION DUMP
information_dump.insert(1,"BLANK","")
information_dump.insert(0,"BLANK1","")
information_dump.insert(0,"Type","New LMS Link")

information_dump["DateTime"] = pd.to_datetime(
    information_dump["DateTime"]
    .astype(str)
    .str.split(":")
    .str[0],
    format="%d-%m-%Y", 
    errors="coerce"
).dt.strftime("%d/%m/%Y")

information_dump["Username"] = information_dump["Username"].apply(lambda x : x.strip().lower())

# HC DUMP 

hc_dump = hc_dump.drop(['Employee Type','Reporting To ID','Functional Reporting To Name'],axis=1)
hc_dump = hc_dump[hc_dump['Location Name'].isin(['Amritsar','Bangalore','Gurgaon','Vadodara','Kolkata'])]
hc_dump.loc[(hc_dump['OU Name'] == 'GP Social Media') & (hc_dump['Location Name'] == 'Gurgaon'),'Location Name'] = 'Amritsar'

# hc_dump.loc[(hc_dump['Department Name'] == 'Training') & (hc_dump['Date Of Joining'].isin(date_list))] 

rows_to_remove = hc_dump[
    (hc_dump['Department Name'] == 'Training') &
    (hc_dump['Date Of Joining'].isin(date_list))
].index

hc_dump.drop(rows_to_remove, inplace=True)

# NEW HC DUMP

#  employee id se TRA and GIG remeove krna !!
new = hc_dump[(hc_dump["OU Name"] != "One Card") & (~hc_dump['Employee ID'].str.startswith(('TRA','GIG')))]


####################################################################################
#WRITING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(template,password="Kipl1234")
    
    data_sheet = wb.sheets['data']
    hc_dump_sheet = wb.sheets['HC Dump']
    scores_sheet = wb.sheets['Scores']
    
    has_error = False  
    
    # DATA SHEET ------------------------------------------------------------
    try:
        df = data_sheet["A1"].options(pd.DataFrame,header=True,index=False, expand="table").value
        df = df[df["Type"] == "OLD LMS Link"]
        df = df.drop(['Type','Month'],axis=1)
        length = len(df)+1+1
        
        # make the no of columns same
        df = df.iloc[:, :len(information_dump.columns)]

        data_sheet["C2"].options(pd.DataFrame,header=False,index=False, expand="table").value = df
        
        data_sheet[f"A{length}"].options(pd.DataFrame,header=False,index=False, expand="table").value = information_dump
        
        print("✅ SUCCESS: Data successfully written to 'DATA SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'DATA SHEET': {e} \n")
        
    # HC DUMP SHEET ------------------------------------------------------------
    try:
        hc_dump_sheet["B2"].options(pd.DataFrame,header=False,index=False, expand="table").value = hc_dump
        
        print("✅ SUCCESS: Data successfully written to 'HC DUMP SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'HC DUMP SHEET': {e} \n")
        
    # SCORES SHEET ------------------------------------------------------------
    try:
        scores_sheet["A2"].options(pd.DataFrame,header=False,index=False, expand="table").value = new
        
        print("✅ SUCCESS: Data successfully written to 'SCORES SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'SCORES SHEET': {e} \n")
        
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
wb = excel.Workbooks.Open(template_output,Password="Kipl1234")

ws1 = wb.Worksheets("data")
ws2 = wb.Worksheets("HC Dump")
ws3 = wb.Worksheets("Scores")

try:
    # -------------------------------------------------------------------------------------------
    # DATA SHEET
    
    ws1.Activate()
    
    last_row = ws1.Cells(ws1.Rows.Count, 1).End(-4162).Row 
    
    formula_range1 = ws1.Range(f"B{length-1}:B{last_row}")
    formula_range2 = ws1.Range(f"D{length-1}:D{last_row}")
    formula_range3 = ws1.Range(f"O{length-1}:AA{last_row}")  
    formula_range1.FillDown()
    formula_range2.FillDown()
    formula_range3.FillDown()
    
    print(f"🎯 Formulas successfully applied to range B{length-1}:B{last_row} and D{length-1}:D{last_row} and O{length-1}:AA{last_row}")
    
    # -------------------------------------------------------------------------------------------
    # HC DUMP SHEET
    
    ws2.Activate()
    last_row = len(hc_dump) + 1

    formula_range1 = ws2.Range(f"A2:A{last_row}")
    formula_range2 = ws2.Range(f"AA2:AB{last_row}")
    formula_range1.FillDown()
    formula_range2.FillDown()
    

    print(f"🎯 Formulas successfully applied to range A2:A{last_row} & AA2:AB{last_row}")
    
    # -------------------------------------------------------------------------------------------
    # SCORES SHEET
    
    ws3.Activate()
    last_row = len(new) + 1
    
    formula_range1 = ws3.Range(f"Z2:AH{last_row}")
    formula_range1.FillDown()

    print(f"🎯 Formulas successfully applied to range Z2:AH{last_row}")

    wb.Save()
    wb.Close()
    
    print("🚀💥 FORMULAS SUCCESSFULLY APPLIED TO ALL SHEETS USING FILLDOWN! MASTERED BY JASKIRAT! 🎯✅")

except Exception as e:
    print(f"❌ ERROR: {e}")
    
finally:
    excel.Quit()
    
    
