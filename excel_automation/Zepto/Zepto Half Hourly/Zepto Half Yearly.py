# -*- coding: utf-8 -*-
"""
Created on Tue Oct 28 15:03:43 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw
from datetime import date,timedelta

####################################################################################
#DUMPS
####################################################################################

chat_raw = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\Zepto Half Hourly\Chat Raw Dump.xlsb",sheet_name="Chat Raw Dump")
# chat_raw = pd.read_excel(r"C:\Users\8242K\Desktop\Chat Raw Dump.xlsb",sheet_name="Chat Raw Dump")

login_hours = pd.read_excel(r"C:\Users\8242K\Desktop\WFM\Zepto Half Hourly\Half Hourly Login Dump - Oct'25.xlsb")
# login_hours = pd.read_excel(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Zepto\Internal Dashboard\2025\Oct'25\Raw Dump\Raw Dumps\Hoop Login Hours Overall\Half Hourly Login Dump - Oct'2025 (MTD).xlsb")

####################################################################################
#TEMPLATE
####################################################################################

half_hourly = r"C:\Users\8242K\Desktop\WFM\Zepto Half Hourly\Half Hourly Login Dump - Oct'25.xlsb"
half_hourly_output = r"C:\Users\8242K\Desktop\WFM\Zepto Half Hourly\Half Hourly Login Dump - Oct'25_OUTPUT.xlsb"
template = r"C:\Users\8242K\Desktop\WFM\Zepto Half Hourly\Zepto Performance Tracker - Oct'25.xlsb"
template_output = r"C:\Users\8242K\Desktop\WFM\Zepto Half Hourly\Zepto Performance Tracker - Oct'25_OUTPUT.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

# DATE
today = date.today()
today_date = today.strftime("%d %b %y")
yesterday = today-timedelta(days=11)
yesterday_date = yesterday.strftime("%d %b %y")

# LOGIN HOURS ----------------------------------------------------------------------------

login_hours = login_hours.iloc[:,3:-1]

# CHAT RAW -------------------------------------------------------------------------------

# REQUIRED COLUMNS
chat_raw = chat_raw[['agent_assigned_at','queue_name','agent_email','wait_time','frt_time','handling_time']]

# TIME COLUMNS
chat_raw['agent_assigned_at'] = pd.to_datetime(chat_raw['agent_assigned_at'], origin='1899-12-30', unit='D')
chat_raw['Interval Wise'] = chat_raw['agent_assigned_at'].dt.floor('30min').dt.time
chat_raw['agent_assigned_at'] = chat_raw['agent_assigned_at'].dt.strftime('%d %b %y')

cols = ['wait_time', 'frt_time', 'handling_time']
for c in cols:
    chat_raw[c] = pd.to_timedelta(chat_raw[c], unit='s')

# FILTER
chat_raw = chat_raw[chat_raw['agent_assigned_at'] != yesterday_date ]
chat_raw = chat_raw[~chat_raw["queue_name"].isin(["RNR Queue","OneSupport Emails Queue"])]
  
#  PIVOT TABLE 
pivot = chat_raw.groupby(['agent_assigned_at','Interval Wise','queue_name','agent_email']).agg(
    chat_count=('agent_email', 'count'),
    handling_time=('handling_time','mean')
).reset_index()

pivot['handling_time'] = pivot['handling_time'].dt.round('1s').apply(lambda x:str(x).split()[-1])
pivot['Interval Wise'] = pivot['Interval Wise'].astype(str)

# INTERVAL WISE
summary = chat_raw.groupby(['agent_assigned_at','queue_name','Interval Wise'])[
    ['handling_time','wait_time', 'frt_time']
    ].mean().apply(lambda x: x.dt.round('1s')).reset_index()

percentile_summary = chat_raw.groupby(['agent_assigned_at','queue_name','Interval Wise'])[
    ['wait_time', 'frt_time']
    ].quantile(0.9).apply(lambda x: x.dt.round('1s')).reset_index()

final_summary = summary.merge(
    percentile_summary,
    on=['agent_assigned_at','queue_name','Interval Wise'],
    how = "left",
    suffixes=('', '_90th')
    )

# DAY WISE
day_summary = chat_raw.groupby(['agent_assigned_at','queue_name'])[
    ['handling_time','wait_time', 'frt_time']
    ].mean().apply(lambda x: x.dt.round('1s')).reset_index()

day_percentile_summary = chat_raw.groupby(['agent_assigned_at','queue_name'])[
    ['wait_time', 'frt_time']
    ].quantile(0.9).apply(lambda x: x.dt.round('1s')).reset_index()

day_summary = day_summary.merge(
    day_percentile_summary,
    on=['agent_assigned_at','queue_name',],
    how = "left",
    suffixes=('', '_90th')
    )

day_summary = day_summary[['queue_name','agent_assigned_at','handling_time', 'wait_time',  'wait_time_90th', 'frt_time', 'frt_time_90th']]
day_summary.iloc[:,:2]


time_cols = ['wait_time', 'frt_time', 'handling_time', 'wait_time_90th', 'frt_time_90th']
final_summary[time_cols] = final_summary[time_cols].apply(lambda col: col.map(lambda x : str(x).split()[-1]))
day_summary[time_cols] = day_summary[time_cols].apply(lambda col: col.map(lambda x : str(x).split()[-1]))

final_summary['Interval Wise'] = final_summary['Interval Wise'].astype(str)

wimo_backup = final_summary[final_summary["queue_name"] == "WIMO Backup"]
d_karma = final_summary[final_summary["queue_name"] == "D Karma Queue"]
bad_quality = final_summary[final_summary["queue_name"] == "Bad Quality Desk"]
missing_item = final_summary[final_summary["queue_name"] == "Missing Item Queue"]
cds_tts = final_summary[final_summary["queue_name"] == "CDS/TTS Desk"]


####################################################################################
#WRITING DATA
####################################################################################

try:
    app = xw.App()
    wb = app.books.open(half_hourly)
        
    has_error = False  
    
    # QUEUE DATA SHEET ------------------------------------------------------------
    
    sheet1 = wb.sheets["Queue Data"] 
    sheet2 = wb.sheets["Hoop Login Hours"] 
    
    try:
        length = len(pivot)+1
        
        sheet1["B2"].options(pd.DataFrame,header=False,index=False, expand="table").value = pivot
        
        print("✅ SUCCESS: Data successfully written to ' QUEUE DATA SHEET'")
        
        sheet1.range("A2").api.AutoFill(Destination=sheet1.range(f"A2:A{length}").api)

        print(f"🎯 Formulas successfully applied to range A2:A{length}")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write ' QUEUE DATA SHEET': {e} \n")
        
     # HOOP LOGIN HOURS SHEET ------------------------------------------------------------

    try:
        length = len(login_hours)+1
        
        sheet2["N2"].options(pd.DataFrame,header=False,index=False, expand="table").value = login_hours
        
        print("✅ SUCCESS: Data successfully written to 'HOOP LOGIN HOURS SHEET'")
        
        sheet2.range("A2:M2").api.AutoFill(Destination=sheet2.range(f"A2:M{length}").api)

        print(f"🎯 Formulas successfully applied to range A2:M{length}")
        
        src = sheet2.range("N2:W2")            # Source format row
        dest = sheet2.range(f"N3:W{length}")   # Destination rows
        
        src.api.Copy()
        dest.api.PasteSpecial(Paste=-4122)     # -4122 → xlPasteFormats
        print(f"🎨 Formats successfully copied to range A3:M{length}")
        
    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'HOOP LOGIN HOURS SHEET': {e} \n")

    wb.save(half_hourly_output)
    wb.close()
    
    if has_error:
        print(f"⚠️ COMPLETED WITH ERRORS: Some sheets failed to update in {template_output}\n")
    else:
        print(f"😄 ALL GOOD: Excel update completed without errors at {template_output}\n")

except Exception as e:
    print(f"🔥 FAILURE: Could not save workbook {template_output}: {e}")
finally:
    app.quit()

print("======================= HOOP LOGIN COMPLETE =======================/n")
####################################################################################
#WRITING DATA
####################################################################################
try:
    app = xw.App()
    wb = app.books.open(template)
        
    has_error = False  
    
    # CHAT AND AHT SHEET ------------------------------------------------------------
    
    sheet1 = wb.sheets["Chat & AHT Status"] 
    
    try:
        length = len(pivot)+1
        
        sheet1["A2"].options(pd.DataFrame,header=False,index=False, expand="table").value = pivot.iloc[:,:-1]
        sheet1["N2"].options(pd.DataFrame,header=False,index=False, expand="table").value = pivot.drop(pivot.columns[4],axis = 1)
        
        print("✅ SUCCESS: Data successfully written to 'CHAT AND AHT SHEET'")
        
        sheet1.range("F2:K2").api.AutoFill(Destination=sheet1.range(f"F2:K{length}").api)
        sheet1.range("S2:U2").api.AutoFill(Destination=sheet1.range(f"S2:U{length}").api)
        print(f"🎯 Formulas successfully applied to range F2:K{length} & F3:P{length} & S3:U{length}")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'CHAT AND AHT SHEET': {e} \n")
    
    print("======================= CHAT AND AHT COMPLETE =======================/n")

    # CAMPAIGN SHEETS ------------------------------------------------------------
    
    def write_sheet(sheet_name,df):
        sheet = wb.sheets[sheet_name]
        try:
            length = len(df)+1+1
            sheet["D3"].options(pd.DataFrame,header=False,index=False, expand="table").value = df[['agent_assigned_at','Interval Wise']]
            sheet["Q3"].options(pd.DataFrame,header=False,index=False, expand="table").value = df[['handling_time', 'wait_time',  'wait_time_90th', 'frt_time', 'frt_time_90th']]
            print(f"✅ SUCCESS: Data successfully written to '{sheet_name}'")
            
            sheet.range("A3:C3").api.AutoFill(Destination=sheet.range(f"A3:C{length}").api)
            sheet.range("F3:P3").api.AutoFill(Destination=sheet.range(f"F3:P{length}").api)
            sheet.range("V3:AT3").api.AutoFill(Destination=sheet.range(f"V3:AT{length}").api)
            print(f"🎯 Formulas successfully applied to range A3:C{length} & F3:P{length} & V3:AT{length}")

        except Exception as e:
            has_error = True
            print(f"❌ ERROR: Failed to write '{sheet_name}': {e} \n")
    
    
    sheet2 = "Wimo Backup - Interval Wise"    
    sheet3 = "Missing Item - Interval Wise"
    sheet4 = "D-Karma - Interval Wise"
    sheet5 = "CDS TTS Desk - Interval Wise"
    sheet6 = "Bad Quality - Interval Wise"
    
    write_sheet(sheet2, wimo_backup)
    write_sheet(sheet3, missing_item)
    write_sheet(sheet4, d_karma)
    write_sheet(sheet5, cds_tts)
    write_sheet(sheet6, bad_quality)
    
    print("======================= CAMPAIGN SHEETS COMPLETE =======================")
    
    # DAY WISE SHEET ------------------------------------------------------------
    
    sheet7 = wb.sheets["Day Wise Performance"] 
    
    def write_sheet2(val,campaign):
        try:
            df = day_summary[day_summary['queue_name'] == campaign]
            length = val + len(df) - 1
            
            sheet7[f"A{val}"].options(pd.DataFrame,header=False,index=False, expand="table").value = df.iloc[:,:2]
            sheet7[f"I{val}"].options(pd.DataFrame,header=False,index=False, expand="table").value = df.iloc[:,2:]
            
            print(f"✅ SUCCESS: Data successfully written to 'Day Wise Performnce' - {campaign}")
            
            if length!=val:
                
                sheet7.range(f"C{val}:H{val}").api.AutoFill(Destination=sheet7.range(f"C{val}:H{length}").api)
                sheet7.range(f"N{val}:S{val}").api.AutoFill(Destination=sheet7.range(f"N{val}:S{length}").api)
                print(f"🎯 Formulas successfully applied to range C{val}:H{length} & N{val}:S{length}")

        except Exception as e:
            has_error = True
            print(f"❌ ERROR: Failed to write 'Day Wise Performnce' - {campaign} : {e} \n")
    
    write_sheet2(2, "WIMO Backup")
    write_sheet2(6, "Missing Item Queue")
    write_sheet2(10, "D Karma Queue")
    write_sheet2(14, "CDS/TTS Desk")
    write_sheet2(18, "Bad Quality Desk")
    
    print("======================= DAY WISE COMPLETE =======================")
        
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

