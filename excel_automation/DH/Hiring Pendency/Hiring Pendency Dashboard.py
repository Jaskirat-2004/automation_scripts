# -*- coding: utf-8 -*-
"""
Created on Thu Oct 16 12:42:50 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw
from win32com.client import Dispatch

####################################################################################
#DUMPS
####################################################################################

candidate_dump = pd.read_csv(r"C:\Users\8242K\Desktop\WFM\DH\Hiring Pendency\CANDIDATE_APPLICATIONS.csv")

####################################################################################
#TEMPLATE
####################################################################################

template = r"C:\Users\8242K\Desktop\WFM\DH\Hiring Pendency\Hiring_Pendency Dashboard.xlsb"
template_output = r"C:\Users\8242K\Desktop\WFM\DH\Hiring Pendency\Hiring_Pendency Dashboard_OUTPUT.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

candidate_dump = candidate_dump[['id', '_id', 'title', 'location', 'Name', 'Email', 'Phone Number',
                                 'source', 'dateApplied', 'position', 'formDataid', 'TRF_Name', 
                                 'processName', 'Test_Score', 'Test_Status', 'Lead_Credit_Date', 
                                 'Call_Status', 'Call_Reason', 'Call_Reason2', 'l1', 'l1 Rating', 
                                 'l1 InterView Schedule Date Time', 'l1 Schedule On', 'l1 Attempt Date', 
                                 'l1 Assign To', 'l1 Assign By', 'l1 Modified Date Time', 
                                 'l1 Rejection Reason', 'l2', 'l2 Rating', 'l2 InterView Schedule Date Time', 
                                 'l2 Schedule On', 'l2 Attempt Date', 'l2 Assign To', 'l2 Assign By', 
                                 'l2 Modified Date Time', 'l2 Rejection Reason', 'l3', 'l3 Rating', 
                                 'l3 InterView Schedule Date Time', 'l3 Schedule On', 'l3 Attempt Date', 
                                 'l3 Assign To', 'l3 Assign By', 'l3 Modified Date Time', 'l3 Rejection Reason', 
                                 'l4', 'l4 Rating', 'l4 InterView Schedule Date Time', 'l4 Schedule On', 
                                 'l4 Attempt Date', 'l4 Assign To', 'l4 Assign By', 'l4 Modified Date Time', 
                                 'l4 Rejection Reason', 'Application Status', 'Recruiter Name', 'Onboard Status', 
                                 'UTM Medium', 'UTM Compaign', 'UTM Term', 'Candidate Status', 'Created At', 'Tin No', 
                                 'Tin Type', 'Form Name', 'Emp Id', 'Document Type', 'designation', 'TAT', 'accidental', 
                                 'mediclaim', 'concentForm', 'isConcentFormPdfExist', 'Source_Application_Form', 
                                 'Date of Birth', 'Preferred_Location', 'What is your preferred mode of interview', 
                                 'Do you identify yourself as having a physical disability']]


####################################################################################
#WRITING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(template,password="DH@2025")
    
    overall_raw_sheet = wb.sheets['Overall Raw']

    has_error = False  
    
    # OVERALL RAW SHEET ------------------------------------------------------------
    try:
        overall_raw_sheet["B2"].options(pd.DataFrame,header=False,index=False, expand="table").value = candidate_dump
        
        print("✅ SUCCESS: Data successfully written to 'OVERALL RAW SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'OVERALL RAW SHEET': {e} \n")
        
        
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
wb = excel.Workbooks.Open(template_output,Password="DH@2025")

ws1 = wb.Worksheets("Overall Raw")

try:
    # -------------------------------------------------------------------------------------------
    # OVERALL RAW SHEET
    
    ws1.Activate()
    
    last_row = ws1.Cells(ws1.Rows.Count, 2).End(-4162).Row 
    
    formula_range1 = ws1.Range(f"A2:A{last_row}")
    formula_range2 = ws1.Range(f"CC2:CK{last_row}")
    formula_range1.FillDown()
    formula_range2.FillDown()
    
    print(f"🎯 Formulas successfully applied to range A2:A{last_row} and CC2:CK{last_row}")
    
    wb.Save()
    wb.Close()
    
    print("🚀💥 FORMULAS SUCCESSFULLY APPLIED TO ALL SHEETS USING FILLDOWN! MASTERED BY JASKIRAT! 🎯✅")

except Exception as e:
    print(f"❌ ERROR: {e}")
    
finally:
    excel.Quit()
    
    
