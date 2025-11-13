# -*- coding: utf-8 -*-
"""
Created on Fri Oct 31 16:54:02 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw
from datetime import date,timedelta

####################################################################################
#DUMPS
####################################################################################

HC  = pd.read_csv(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\PAN INDIA\Naps & Nats Report\head_count_job.csv")

####################################################################################
#TEMPLATE
####################################################################################

template = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\PAN INDIA\Naps & Nats Report\Naps & NATS Report Oct'25.xlsb"
# template_output = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\PAN INDIA\Naps & Nats Report\Naps & NATS Report Oct'25_OUTPUT.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

# DATE
today = date.today()
yesterday = today-timedelta(days=1)
month = yesterday.strftime("%Y-%m")
yesterday_date = yesterday.strftime("%Y-%m-%d")

# HC DUMP ----------------------------------------------------------------------------

HC = HC.drop(['Employee Type','Reporting To ID','Functional Reporting To ID','row_number','company'],axis = 1)

HC = HC[HC['Employee ID'].astype(str).apply(lambda x: x.startswith('N'))]

yesterday_HC =  HC[HC['Date Of Joining'] == yesterday_date]

month_HC = HC[HC['Date Of Joining'].astype(str).str.startswith(month)]

####################################################################################
#WRITING DATA
####################################################################################

if len(yesterday_HC) > 0:
    try:
        app = xw.App()
        wb = app.books.open(template)
            
        has_error = False  
        
        # JOINING RAW SHEET ------------------------------------------------------------
        
        sheet1 = wb.sheets["Joining Raw"] 
        
        try:
            last_row = sheet1.range('A' + str(sheet1.cells.last_cell.row)).end('up').row
            next_row = last_row + 1
            
            sheet1[f"D{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = yesterday_HC
            
            df = sheet1["D1"].options(pd.DataFrame,header=True,index=False, expand="table").value
            
            missing = month_HC[~month_HC['Employee ID'].isin(df['Employee ID'])]
            
            next_row = next_row + len(yesterday_HC)
            
            sheet1[f"D{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = missing
            
            print("✅ SUCCESS: Data successfully written to ' JOINING RAW SHEET'")
            
            drag = sheet1.range('D' + str(sheet1.cells.last_cell.row)).end('up').row
    
            sheet1.range(f"A{last_row}:C{last_row}").api.AutoFill(Destination=sheet1.range(f"A{last_row}:C{drag}").api)
            sheet1.range(f"AE{last_row}:AF{last_row}").api.AutoFill(Destination=sheet1.range(f"AE{last_row}:AF{drag}").api)
    
            print(f"🎯 Formulas successfully applied to range A{last_row}:C{drag} & AE{last_row}:AF{drag}")
    
        except Exception as e:
            has_error = True
            print(f"❌ ERROR: Failed to write 'JOINING RAW SHEET': {e} \n")
    
        wb.save(template)
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
else:
    print("<<<<< NO DATA TO BE PASTED >>>>>")

      
    
    
    
