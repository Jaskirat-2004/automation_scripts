# -*- coding: utf-8 -*-
"""
Created on Wed Sep 17 14:47:29 2025

@author: JASKIRAT
"""

import pandas as pd
import xlwings as xw     
from win32com.client import Dispatch
import datetime

# ==========================================================================================
# DUMPS

Agent_productivity_dump = pd.read_csv(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Vodafone APR\Dump\custom_agent_productivity_interval_summary.csv")
shrinkage_dump = pd.read_excel(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Vodafone APR\Dump\Shrinkage.xlsx", sheet_name = "For Central",header=2)

template = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Vodafone APR\Template\Neo APR Data Sept'25.xlsb"
output_file = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\WFM Automation\Output\Vodafone APR\Template\Neo APR Data Sept'25_Output.xlsb"

# ==========================================================================================
# MODIFICATIONS Agent dump

Agent_productivity_dump = Agent_productivity_dump.dropna(subset=['Interval Start'])
Agent_productivity_dump["User ID"] = Agent_productivity_dump['User ID'].str.split("@").str[0]

# Agent_productivity_dump["Interval Start"] = pd.to_datetime(Agent_productivity_dump['Interval Start']).dt.to_pydatetime()
# Agent_productivity_dump["Interval End"] = pd.to_datetime(Agent_productivity_dump['Interval End']).dt.to_pydatetime()

# ==========================================================================================
# MODIFICATIONS shrinkage dump

shrinkage_dump = shrinkage_dump[shrinkage_dump["Attendance"].isin(["Present","HD"])]
shrinkage_dump = shrinkage_dump.drop('ss',axis=1)
shrinkage_dump = shrinkage_dump.drop('Unnamed: 19',axis=1)
shrinkage_dump = shrinkage_dump.drop('Unnamed: 20',axis=1)


# changing time delta to string
# cols = shrinkage_dump.select_dtypes(include=['object'])

cols = ["Sum of Login Hrs. With Exception","Login Hrs.","Sum of THT"]
for col in cols:
    shrinkage_dump[col] = shrinkage_dump[col].apply(lambda x: x.strftime("%H:%M:%S") if isinstance(x,datetime.time) else x)

# ==========================================================================================
# WRITE IN TEMPLATE

try:
    wb = xw.Book(template, password ='3296')
    
    APR_sheet = wb.sheets['APR']
    shrinkage_sheet = wb.sheets['custom_agent_productivity_inter']
    
    APR_sheet["B2"].options(pd.DataFrame, header=False, index=False, expand='table').value = shrinkage_dump
    shrinkage_sheet["A2"].options(pd.DataFrame, header=False, index=False, expand='table').value = Agent_productivity_dump
    
    wb.save(output_file)
    wb.close()
 
    print(f"Data successfully written to {output_file}")
 
except Exception as e:
    print(f"Error while saving data to Excel using xlwings: {e}")

# ==========================================================================================
# FORMULA DRAG
 
excel = Dispatch("Excel.Application")
excel.Visible = True
source = excel.Workbooks.Open(output_file,Password = "3296")

ws1 = source.Worksheets("APR")
ws2 = source.Worksheets("custom_agent_productivity_inter")

try:
    ws1.Activate()
    ws2.Activate()
 
    le1 = str(len(shrinkage_dump)+1)
    le2 = str(len(Agent_productivity_dump)+1)
 
    destination_range1 = ws1.Range(f"A2:A{le1}")
    destination_range2 = ws1.Range(f"T2:T{le1}")
    destination_range3 = ws2.Range(f"BJ2:BP{le2}")
 
    destination_range1.FillDown()
    destination_range2.FillDown()
    destination_range3.FillDown()
 
    print(f"Formulas successfully applied to range A2:A{le1}, T2:T{le1} & BJ2:BP{le2}")

except Exception as e:
    print(f"Error while applying formulas on tickets dump sheet using pywin32: {e}")    
 
print("Formulas successfully applied to all sheets using FillDown.")
 
source.Save()
source.Close()
excel.Quit()