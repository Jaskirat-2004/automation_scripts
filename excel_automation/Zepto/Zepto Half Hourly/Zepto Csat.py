# -*- coding: utf-8 -*-
"""
Created on Mon Oct 27 16:34:55 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw
import os
from datetime import date,timedelta
# from win32com.client import Dispatch

# DATE
today = date.today()
today_date = today.strftime("%d %b")
yesterday = today-timedelta(days=3)
yesterday_date = yesterday.strftime("%Y-%m-%d")
yesterday_date2 = yesterday.strftime("%d %b")

####################################################################################
#DUMPS
####################################################################################

folder_path1 = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Zepto\Internal Dashboard\2025\Oct'25\Raw Dump\Raw Dumps\Queue Level Overall"
queue_level_path = [f for f in os.listdir(folder_path1) if f.startswith((yesterday_date2))]

if queue_level_path:
    full_path1 = os.path.join(folder_path1, queue_level_path[0])
    
queue_level_dump = pd.read_csv(full_path1)

folder_path2 = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Zepto\Internal Dashboard\2025\Oct'25\Raw Dump\Raw Dumps\Ticket Dump Overall"
ticket_path = [f for f in os.listdir(folder_path2) if f.startswith((yesterday_date2))]

if ticket_path:
    full_path2 = os.path.join(folder_path2, ticket_path[0])
    
ticket_dump = pd.read_csv(full_path2)

####################################################################################
#TEMPLATE
####################################################################################

# template = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Zepto\Internal Dashboard\2025\Oct'25\Raw Dump\Raw Dumps\C-Sat Template\C-Sat Preparing Template - 24-Oct.xlsb"
# template_output = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Zepto\Internal Dashboard\2025\Oct'25\Raw Dump\Raw Dumps\C-Sat Template\C-Sat Preparing Template - "+today_date+".xlsb"
# CSAT_template = r"C:\Users\8242K\Desktop\WFM\Zepto Half Hourly\Chat Raw Dump.xlsb"
# CSAT_template_output = r"C:\Users\8242K\Desktop\WFM\Zepto Half Hourly\Chat Raw Dump_output.xlsb"

template=r"C:\Users\8242K\Desktop\WFM\Zepto\Zepto_Csat\Template\C-Sat Preparing Template - 24-Oct.xlsb"
template_output=r"C:\Users\8242K\Desktop\WFM\Zepto\Zepto_Csat\Template\C-Sat_OOOOOOOUUUUUUTTTTTT.xlsb"
CSAT_template = r"C:\Users\8242K\Desktop\WFM\Zepto Half Hourly\Chat Raw Dump.xlsb"
CSAT_template_output = r"C:\Users\8242K\Desktop\WFM\Zepto Half Hourly\Chat Raw Dump_output.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

# QUEUE LEVEL DUMP
queue_level_dump['queue_assigned_at'] = pd.to_datetime(queue_level_dump['queue_assigned_at'])
queue_level_dump.sort_values(by = 'queue_assigned_at', ascending = True , inplace = True)

# TICKET DUMP 
ticket_dump['tkt_create_time'] = pd.to_datetime(queue_level_dump['tkt_create_time'])
ticket_dump.sort_values(by = 'tkt_create_time', ascending = True , inplace = True)

ticket_dump = ticket_dump.dropna(subset = ["rating"])
ticket_dump = ticket_dump[ticket_dump['agent_ai_split'] == "Agent"]

####################################################################################
#WRITING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(template)
    
    sheet1 = wb.sheets["Queue Level"]    
    sheet2 = wb.sheets["Ticket Report"]
    
    has_error = False  
    
    # TICKET REPORT SHEET ------------------------------------------------------------
    try:
        sheet2.clear()
        sheet2["A1"].options(pd.DataFrame,header=True,index=False, expand="table").value = ticket_dump
        
        print("✅ SUCCESS: Data successfully written to 'TICKET REPORT SHEET'")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'TICKET REPORT SHEET': {e} \n")
    
    #QUEUE LEVEL SHEET ------------------------------------------------------------
    try:
        last_row1 = sheet1.range('C' + str(sheet1.cells.last_cell.row)).end('up').row
        sheet1.range(f"A3:AC{last_row1}").clear_contents()
        length = len(queue_level_dump)+1
        sheet1["C2"].options(pd.DataFrame,header=False,index=False, expand="table").value = queue_level_dump
        print("✅ SUCCESS: Data successfully written to 'QUEUE LEVEL SHEET'")
        
        sheet1.range("A2:B2").api.AutoFill(Destination=sheet1.range(f"A2:B{length}").api)
        sheet1.range("X2:AC2").api.AutoFill(Destination=sheet1.range(f"X2:AC2{length}").api)
        print(f"🎯 Formulas successfully applied to range A2:B{length} & X2:AC{length}")
        
        df =  sheet1["A1"].options(pd.DataFrame,header=True,index=False, expand="table").value
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'QUEUE LEVEL SHEET': {e} \n")
        

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
#MODIFICATIONS
####################################################################################

df = df[df['Partner By'] == 'Maxicus']
df = df[df['Capture By'] == 'Consider']
df = df[df['rating'] != '-']
df = df.drop(['Partner By','Capture By'],axis = 1)

####################################################################################
#WRITING DATA
####################################################################################
try:
    
    app = xw.App()
    wb = app.books.open(CSAT_template)
    
    sheet1 = wb.sheets["C-Sat Raw"]    
    
    has_error = False  
    
    #CSAT RAW SHEET ------------------------------------------------------------
    try:
        last_row1 = sheet1.range('C' + str(sheet1.cells.last_cell.row)).end('up').row
        next_row = last_row1+1
        length = last_row1 + len(df) + 1
        
        sheet1[f"C{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = df
        print("✅ SUCCESS: Data successfully written to 'CSAT RAW SHEET'")
        
        sheet1.range(f"A{last_row1}:B{last_row1}").api.AutoFill(Destination=sheet1.range(f"A{last_row1}:B{length}").api)
        sheet1.range(f"AD{last_row1}:AG{last_row1}").api.AutoFill(Destination=sheet1.range(f"AD{last_row1}:AG{length}").api)
        print(f"🎯 Formulas successfully applied to range A{last_row1}:B{length} & AD{last_row1}:AG{length}")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'CSAT RAW SHEET': {e} \n")
      
    wb.save(CSAT_template_output)
    wb.close()
    
    if has_error:
        print(f"⚠️ COMPLETED WITH ERRORS: Some sheets failed to update in {template_output}\n")
    else:
        print(f"😄 ALL GOOD: Excel update completed without errors at {template_output}\n")

except Exception as e:
    print(f"🔥 FAILURE: Could not save workbook {CSAT_template_output}: {e}")
finally:
    app.quit()

