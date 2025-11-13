# -*- coding: utf-8 -*-
"""
Created on Tue Nov  4 11:02:00 2025

@author: JASKIRAT
"""

import duckdb
import pandas as pd
import xlwings as xw

####################################################################################
#DUMPS
####################################################################################

sheet = "1-Nov"

performance  = pd.read_excel(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Meesho\SS Chat Perofrmance Dashboard\Dumps\Performance Sheet.xlsx",sheet_name=sheet)

####################################################################################
#TEMPLATE
####################################################################################

template = r"C:\Users\8242K\Desktop\WFM\Meesho\SS Chat\Meesho SS Chat Performance Report.xlsb"
template_output = r"C:\Users\8242K\Desktop\WFM\Meesho\SS Chat\Meesho SS Chat Performance Report_OUTPUT.xlsb"

# template = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Meesho\SS Chat Perofrmance Dashboard\Template\Meesho SS Chat Performance Report.xlsb"
# template_output = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Meesho\SS Chat Perofrmance Dashboard\Template\Meesho SS Chat Performance Report_OUTPUT.xlsb"

####################################################################################
#MODIFICATIONS
####################################################################################

performance = performance[
    ['Date', 'Emp ID', 'Agent Name', 'Attendance', 'Today Target', 'TL Name', 'Queue']
    ]

####################################################################################
#WRITING DATA
####################################################################################

try:
    app = xw.App()
    wb = app.books.open(template)
        
    has_error = False  
    
    sheet1 = wb.sheets["Date Wise raw Dump"] 
    
    # RAW SHEET ------------------------------------------------------------
    
    try:
        last_row = sheet1.range('B' + str(sheet1.cells.last_cell.row)).end('up').row
        next_row = last_row + 1
        
        sheet1[f"B{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = performance.iloc[:,:5]
        sheet1[f"N{next_row}"].options(pd.DataFrame,header=False,index=False, expand="table").value = performance.iloc[:,5:]

        print("✅ SUCCESS: Data successfully written to 'RAW SHEET'")
        
        drag = sheet1.range('B' + str(sheet1.cells.last_cell.row)).end('up').row

        sheet1.range(f"A{last_row}:A{last_row}").api.AutoFill(Destination=sheet1.range(f"A{last_row}:A{drag}").api)
        sheet1.range(f"G{last_row}:K{last_row}").api.AutoFill(Destination=sheet1.range(f"G{last_row}:K{drag}").api)

        print(f"🎯 Formulas successfully applied to range A{last_row}:H{drag} & W{last_row}:Y{drag}")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'RAW SHEET': {e} \n")
        
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

####################################################################################
#DUMPS
####################################################################################

con = duckdb.connect()

df = con.execute("""
SELECT "Ticket ID","Center","SSAT","Reopen","Ticket ID_1","Subject","Status","3PL L2","Group","Resolved Date Time","Last Conversation Time","Emp Code"
FROM read_xlsx('C:/Users/8242K/Desktop/WFM/Meesho/SS Chat\Kochar Nov25 SSAT & Reopen.xlsx'
               )
""").df()


# df = con.execute("""
# SELECT "Ticket ID","Center","SSAT","Reopen","Ticket ID_1","Subject","Status","3PL L2","Group","Resolved Date Time","Last Conversation Time","Emp Code"
# FROM read_xlsx('//172.17.52.16/172.17.3.195-data/KocharWFM/Data Science/WFM Automation/Output/Meesho/SS Chat Perofrmance Dashboard/Dumps/Kochar Nov25 SSAT & Reopen.xlsx')
# """).df()

####################################################################################
#MODIFICATIONS
####################################################################################

df["Resolved Date Time"] = pd.to_datetime(df["Resolved Date Time"], errors='coerce', format="%d/%m/%Y %H:%M:%S")
df["Last Conversation Time"] = pd.to_datetime(df["Last Conversation Time"], format="%H:%M:%S", errors='coerce').dt.time
df["Last Conversation Time"] = df["Last Conversation Time"].astype(str)

df["Emp Code"] = df["Emp Code"].astype(str).str.replace(r"^MKOC", "", regex=True)

####################################################################################
#WRITING DATA
####################################################################################

try:
    app = xw.App()
    wb = app.books.open(template)
        
    has_error = False  
    
    sheet1 = wb.sheets["Raw"] 

    # RAW SHEET ------------------------------------------------------------
    
    try:
        
        sheet1["J3"].options(pd.DataFrame,header=False,index=False, expand="table").value = df.iloc[:,:7]
        sheet1["R3"].options(pd.DataFrame,header=False,index=False, expand="table").value = df.iloc[:,7:]

        print("✅ SUCCESS: Data successfully written to 'RAW SHEET'")
        
        drag = len(df) + 1 + 1

        sheet1.range("A3:I3").api.AutoFill(Destination=sheet1.range(f"A3:I{drag}").api)
        sheet1.range("Q3").api.AutoFill(Destination=sheet1.range(f"Q3:Q{drag}").api)

        print(f"🎯 Formulas successfully applied to range A3:I{drag} & Q3:Q{drag}")

    except Exception as e:
        has_error = True
        print(f"❌ ERROR: Failed to write 'RAW SHEET': {e} \n")

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
    


