# -*- coding: utf-8 -*-
"""
Created on Wed Sep 17 10:16:44 2025

@author: JASKIRAT (8242k)
"""
import pandas as pd 
import xlwings as xw
from win32com.client import Dispatch

##########################################################
# IMPPORT PATHS

agent_productivity_dump = pd.read_csv(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\charu\Vodafone Neo\Dump\custom_agent_productivity_interval_summary_2025-09-16_16_54_28(runnableReportId1758021867653).csv")
shrinkage_dump = pd.read_excel(r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\charu\Vodafone Neo\Dump\Shrinkage Sept'25.xlsx", sheet_name = "For Central",header=2)

template = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\charu\Vodafone Neo\Template\Neo APR Data Sept'25.xlsb"
output_file = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Data Science\charu\Vodafone Neo\Template\Neo APR Data Sept'25_output.xlsb"

##########################################################
# MODIFICATOINS in Agent

agent_productivity_dump = agent_productivity_dump.dropna(subset=['Interval Start'])
agent_productivity_dump["User ID"] = agent_productivity_dump["User ID"].str.split("@").str[0]


agent_productivity_dump['Interval Start'] = pd.to_datetime(agent_productivity_dump['Interval Start'])
#df["col"] = pd.to_datetime(df["col"]).dt.to_pydatetime()


##########################################################
# MODIFICATOINS in SHRINCAGE

shrinkage_dump = shrinkage_dump[shrinkage_dump["Attendance"].isin(["Present","HD"])]

##########################################################
# WRITTING DATA

try:
    wb = xw.Book(template,password='3296')
    
    APR_sheet = wb.sheets['APR']
    shrinkage_sheet = wb.sheets['custom_agent_productivity_inter']
    
    APR_sheet["A2"].options(pd.DataFrame,header=False,index=False, expand="table").value = shrinkage_dump
    shrinkage_sheet["B2"].options(pd.DataFrame,header=False,index=False, expand="table").value = agent_productivity_dump
    
    wb.save(output_file)
    wb.close()

    print(f"Data successfully written to {output_file}")

except Exception as e:
    print(f"Error while saving data to Excel using xlwings: {e}")

##########################################################
# FORMULA FILLDOWN
    
excel = Dispatch("Excel.Application")
excel.Visible = True
source = excel.Workbooks.Open(output_file, Password = '3296')

ws1 = source.Worksheets("APR")
ws2 = source.Worksheets("custom_agent_productivity_inter")

try:
    ws1.Activate()
    ws2.Activate()
    
    le1 = str(len(agent_productivity_dump)+1)
    le2 = str(len(shrinkage_dump)+1)

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
source.SaveAs(output_file)
source.Close()
excel.Quit()
