

# -*- coding: utf-8 -*-
"""
Created on Wed Sep 24 15:34:08 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw
import datetime
from win32com.client import Dispatch

####################################################################################
#DUMPS
####################################################################################

ivr_dump = pd.read_excel(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Dumps\IVRFeedbackReport.xlsx",skiprows=2)

agent_break_dump = pd.read_csv(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Dumps\Agent Break.csv")

agent_state_dump = pd.read_csv(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Dumps\Agent State Default Details.csv")

call_generate_dump = pd.read_excel(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Dumps\CallGenerate.xls")

shrinkage_dump = pd.read_excel(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Dumps\Acko IB Shrinkage & FTE.xlsx",sheet_name="HC Dump")

activity_dump = pd.read_csv(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Dumps\activity.csv")

agent_performance_dump = pd.read_csv(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Dumps\agent_performance_-_overall.csv")

email_productivity_dump = pd.read_excel(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Dumps\Email Productivity (Responses).xlsx",sheet_name="Form Responses 1")

agent_login = pd.read_excel(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Dumps\Agent Login.xlsx")

####################################################################################
#TEMPLATE
####################################################################################

template = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Template\Acko Agent Hygiene Report.xlsb"
template_output = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Acko Reports\Acko Agent Hygine Report\Template\Acko Agent Hygiene Report_Output.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

# DATE
yesterday = (datetime.date.today()-datetime.timedelta(days=5))
yesterday_date = yesterday.strftime("%d/%m/%Y")

# IVR DUMP
ivr_dump = ivr_dump.iloc[:-1,:13]
cols_to_str = ['Call ID','Caller Number','Called Number']
for col in cols_to_str:
    ivr_dump[col] = pd.to_numeric(ivr_dump[col], errors='coerce')
    # ivr_dump[col] = ivr_dump[col].apply(lambda x: "'" + str(int(x)) if pd.notna(x) else x)

# AGEND BREAK DUMP
agent_break_dump = agent_break_dump.iloc[:-1,:4]
agent_break_dump.insert(0,"Date",yesterday_date)

# AGENT STATE DUMP
agent_state_dump = agent_state_dump.drop(["Agent ID","Total Idle Time"],axis=1)
agent_state_dump.insert(0,"Date",yesterday_date)

# CALL GENERTATE
col_to_drop = ['Caller_E164','Queue Time','Feedback', 'Customer Ring Time','Agent ID', 'Ratings', 'Rating Comments', 'DynamicDid', 'DID']
call_generate_dump = call_generate_dump.drop(col_to_drop,axis=1,errors='ignore')
# call_generate_dump['Call ID'] = call_generate_dump['Call ID'].apply(lambda x: "'" + str(int(x)) if pd.notna(x) else x)
call_generate_dump['Call ID'] = pd.to_numeric(call_generate_dump['Call ID'], errors='coerce')
call_generate_dump['Caller No'] = pd.to_numeric(call_generate_dump['Caller No'], errors='coerce')

new_cols = call_generate_dump["Skill"].str.split("->",expand = True)
call_generate_dump = call_generate_dump.join(new_cols)

# SHRINKAGE DUMP
shrinkage_dump = shrinkage_dump[['Date','Login ID','Name','TL Name','Department']]
shrinkage_dump = shrinkage_dump[shrinkage_dump['Date'].dt.date == yesterday]

# ACTIVITY DUMP
activity_dump['Tickets reassigned from agent'] = ""
activity_dump.insert(loc = 0, column = 'Process', value="tech")
activity_dump.insert(0,"Date",yesterday_date)

# AGENT PERFORMANCE DUMP
agent_performance_dump.insert(loc = 8, column = "", value="")
agent_performance_dump.insert(loc = 0, column = 'Process', value="GI")
agent_performance_dump = agent_performance_dump.drop("Source",axis=1)
agent_performance_dump.insert(0,"Date",yesterday_date)

# EMAIL PRODUCTIVITY RESPOSES
email_productivity_dump = email_productivity_dump[['Date','Ticket ID','Your Name','Current Status of the Ticket','Email Address']]
email_productivity_dump['Date'] = pd.to_datetime(email_productivity_dump['Date'], errors='coerce')
email_productivity_dump = email_productivity_dump.dropna(subset=['Date'])
email_productivity_dump = email_productivity_dump[email_productivity_dump['Date'].dt.date == yesterday]

# AGENT LOGIN DUMP
# All modifications done below in mapper

####################################################################################
#WRITING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(template,password="Acko@2024")
    
    ivr_sheet = wb.sheets["IVR Dump"]    
    break_details_sheet = wb.sheets["Break Details"]
    agent_state_sheet = wb.sheets["Agent State summary"]
    cdr_sheet = wb.sheets["CDR Raw"]
    agent_details_sheet = wb.sheets["Agent Details"]
    responses_sheet = wb.sheets["Responses Data"]
    g_sheet = wb.sheets["G sheet Responses"]
    mapper_sheet = wb.sheets["Mapper"]
    login_logout_sheet = wb.sheets["Login Logout Dump"]
    
    has_error = False  
    
    # IVR SHEET ------------------------------------------------------------
    try:
        last_row1 = ivr_sheet.range('G' + str(ivr_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row1 + 1
        ivr_sheet[f"G{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = ivr_dump
        
        print("✅ SUCCESS: Data successfully written to 'IVR SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'IVR SHEET': {e} \n")
        
    # BREAK DETAILS SHEET ------------------------------------------------------------
    try:
        last_row2 = break_details_sheet.range('C' + str(break_details_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row2 + 1
        break_details_sheet[f"C{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = agent_break_dump
        
        print("✅ SUCCESS: Data successfully written to 'BREAK DETAILS SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'BREAK DETAILS SHEET': {e} \n")
    
    # AGENT STATE SUMMARY SHEET ------------------------------------------------------------
    try:
        last_row3 = agent_state_sheet.range('E' + str(agent_state_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row3 + 1
        agent_state_sheet[f"E{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = agent_state_dump
        
        print("✅ SUCCESS: Data successfully written to 'AGENT STATE SUMMARY SHEET")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'AGENT STATE SUMMARY SHEET': {e} \n")
    
    # CDR RAW SHEET ------------------------------------------------------------
    try:
        last_row4 = cdr_sheet.range('AC' + str(cdr_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row4 + 1
        cdr_sheet[f"AC{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = call_generate_dump
        
        num_rows = len(call_generate_dump)
        caller_col_index = call_generate_dump.columns.get_loc("Caller No") + 1  # Excel is 1-indexed
        cdr_sheet.range(f"AC{next_row}:AC{next_row + num_rows - 1}").number_format = "0"

        print("✅ SUCCESS: Data successfully written to 'CDR RAW SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'CDR RAW SHEET': {e} \n")
    
    # AGENT DETAILS SHEET ------------------------------------------------------------
    try:
        last_row5 = agent_details_sheet.range('C' + str(agent_details_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row5 + 1
        agent_details_sheet[f"C{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = shrinkage_dump[['Date','Login ID','Name']]
        agent_details_sheet[f"BH{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = shrinkage_dump[['TL Name','Department']]
        print("✅ SUCCESS: Data successfully written to AGENT DETAILS SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'AGENT DETAILS SHEET': {e}\n")
    
    # RESPONSES DATA SHEET ------------------------------------------------------------
    try:
        
        last_row6 = responses_sheet.range('C' + str(responses_sheet.cells.last_cell.row)).end('up').row
        next_row = last_row6 + 1
        responses_sheet[f"C{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = activity_dump
        
        next_row = last_row6 + len(activity_dump) + 1
        responses_sheet[f"C{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = agent_performance_dump
        
        print("✅ SUCCESS: Data successfully written to 'RESPONSES DATA SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'RESPONSES DATA SHEET': {e}\n")
    
    # G SHEET RESPONSES ------------------------------------------------------------
    try:
        last_row7 = g_sheet.range('C' + str(g_sheet.cells.last_cell.row)).end('up').row   #### since pasting is starting from 'AC'
        next_row = last_row7 + 1
        g_sheet[f"C{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = email_productivity_dump
        
        print("✅ SUCCESS: Data successfully written to 'G SHEEET RESPONSES'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'G SHEET RESPONSES': {e}\n")
    
    # MAPPER SHEET ------------------------------------------------------------
    try:
        df = mapper_sheet["A1"].options(pd.DataFrame,header=True,index=False, expand="table").value
        df = df[~df['Status'].isin(['-','Inhouse','Test'])]
        df = df.iloc[:,1:2]
        
        dump = agent_login[agent_login['Agent Name'].isin(df['Name'])].copy()
        
        print("✅ SUCCESS: Data successfully copied from 'MAPPER SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'MAPPER SHEEET': {e}\n")
       
    # LOGIN LOGOUT DUMP SHEET ------------------------------------------------------------
    try:
        last_row8 = login_logout_sheet.range('E' + str(login_logout_sheet.cells.last_cell.row)).end('up').row   
        next_row = last_row8 + 1
        login_logout_sheet[f"E{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = dump
        
        print("✅ SUCCESS: Data successfully written to 'LOGIN LOGOUT SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'LOGIN LOGOUT SHEEET': {e}\n")
       
        
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
wb = excel.Workbooks.Open(template_output,Password="Acko@2024")

ws1 = wb.Worksheets("IVR Dump")
ws2 = wb.Worksheets("Break Details")
ws3 = wb.Worksheets("Agent State summary")
ws4 = wb.Worksheets("CDR Raw")
ws5 = wb.Worksheets("Agent Details")
ws6 = wb.Worksheets("Responses Data")
ws7 = wb.Worksheets("G sheet Responses")
ws8 = wb.Worksheets("Login Logout Dump")

try:
    # -------------------------------------------------------------------------------------------
    # IVR Dump
    
    ws1.Activate()
    
    last_row_IVR_Dump = ws1.Cells(ws1.Rows.Count, 5).End(-4162).Row 
    
    formula_range1 = ws1.Range(f"A{last_row1}:F{last_row_IVR_Dump}")  # Expanding A{last_row1}:F down to the last used row
    formula_range1.FillDown()
    
    print(f"🎯 Formulas successfully applied to range A{last_row1}:F{last_row_IVR_Dump}")
    
    # -------------------------------------------------------------------------------------------
    # Break Details
    
    ws2.Activate()
    
    last_row_Break_Details = ws2.Cells(ws2.Rows.Count, 5).End(-4162).Row
    
    formula_range1 = ws2.Range(f"A{last_row2}:B{last_row_Break_Details}")  # Expanding A{last_row2}:B down to the last used row
    formula_range2 = ws2.Range(f"H{last_row2}:I{last_row_Break_Details}")  # Expanding H{last_row2}:I down to the last used row
    formula_range1.FillDown()
    formula_range2.FillDown()
    
    print(f"🎯 Formulas successfully applied to range A{last_row2}:B{last_row_Break_Details} & H2:I{last_row_Break_Details}")
     
    # -------------------------------------------------------------------------------------------
    # Agent State summary
    
    ws3.Activate()
    
    last_row_Agent = ws3.Cells(ws3.Rows.Count, 5).End(-4162).Row
    
    formula_range1 = ws3.Range(f"A{last_row3}:D{last_row_Agent}")  # Expanding A2:D down to the last used row
    formula_range2 = ws3.Range(f"N{last_row3}:P{last_row_Agent}")  # Expanding N2:P down to the last used row
    formula_range1.FillDown()
    formula_range2.FillDown()
    
    print(f"🎯 Formulas successfully applied to range A{last_row3}:D{last_row_Agent} & N{last_row3}:P{last_row_Agent}")
    
    # -------------------------------------------------------------------------------------------
    # CDR Raw
    
    ws4.Activate()
    
    last_row_CDR_Raw = ws4.Cells(ws4.Rows.Count, 5).End(-4162).Row 
    
    formula_range1 = ws4.Range(f"A{last_row4}:AB{last_row_CDR_Raw}")  # Expanding A{last_row4}:AB down to the last used row
    formula_range1.FillDown()
    print(f"🎯 Formulas successfully applied to range A{last_row4}:AB{last_row_CDR_Raw}")
    
    # -------------------------------------------------------------------------------------------
    # Agent Details
    
    ws5.Activate()
    
    last_row = ws5.Cells(ws5.Rows.Count, 5).End(-4162).Row 
    
    formula_range1 = ws5.Range(f"A{last_row5}:B{last_row}")  # Expanding A{last_row5}:B down to the last used row
    formula_range2 = ws5.Range(f"F{last_row5}:BG{last_row}")  # Expanding F{last_row5}:BG down to the last used row
    formula_range1.FillDown()
    formula_range2.FillDown()
    print(f"🎯 Formulas successfully applied to range F{last_row5}:BG{last_row}")
    
    # -------------------------------------------------------------------------------------------
    # Responses Data
    
    ws6.Activate()
    
    last_row = ws6.Cells(ws6.Rows.Count, 5).End(-4162).Row 
    
    formula_range1 = ws6.Range(f"A{last_row6}:B{last_row}")  # Expanding A{last_row6}:B down to the last used row
    formula_range1.FillDown()
    print(f"🎯 Formulas successfully applied to range A{last_row6}:C{last_row}")
    
    # -------------------------------------------------------------------------------------------
    # G sheet Responses
    
    ws7.Activate()
    
    last_row = ws7.Cells(ws7.Rows.Count, 5).End(-4162).Row 
    
    formula_range1 = ws7.Range(f"A{last_row7}:B{last_row}")  # Expanding A{last_row7}:B down to the last used row
    formula_range1.FillDown()
    print(f"🎯 Formulas successfully applied to range A{last_row7}:B{last_row}")
    
    # -------------------------------------------------------------------------------------------
    # Login Logout Dump
    
    ws8.Activate()
    
    last_row = ws8.Cells(ws8.Rows.Count, 5).End(-4162).Row 
    
    formula_range1 = ws8.Range(f"A{last_row8}:D{last_row}")  # Expanding A{last_row8}:D down to the last used row
    formula_range1.FillDown()
    print(f"🎯 Formulas successfully applied to range A{last_row8}:D{last_row}")
    
    # -------------------------------------------------------------------------------------------
    
    wb.Save()
    wb.Close()
    excel.Quit()
    
    print("🚀💥 FORMULAS SUCCESSFULLY APPLIED TO ALL SHEETS USING FILLDOWN! MASTERED BY JASKIRAT! 🎯✅")

except Exception as e:
    print(f"❌ ERROR: {e}")
    
