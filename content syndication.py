# -*- coding: utf-8 -*-
"""
Created on Mon Oct 03 12:53:00 2016

@author: Qi Yi
"""

import pandas as pd

#two ways to read excel file in python
xl = pd.ExcelFile("C:\\Users\\Qi Yi\\Desktop\inSegment - Arbor Networks - Weekly Lead Report 7.7.2016.xlsx")
leads = xl.parse("All Leads") #sheet name
#leads = pd.read_excel('C:\\Users\\Qi Yi\\Desktop\inSegment - Arbor Networks - Weekly Lead Report 7.7.2016.xlsx', sheetname="All Leads")

unique = leads.drop_duplicates()
len(unique) #16153

flag = unique[unique.duplicated(["Email Address"], keep=False)]
flag = flag.sort_values(["Email Address"], ascending=True)
flag.to_excel("C:\\Users\\Qi Yi\\Desktop\\flag.xlsx")     

assets = ["DDoS_Security_Myths_Busted", "SecCurrent_The_Hunted_Becomes_the_Hunter",\
          "ITHarvest_Report_Security_Analytics__A_Required_Escalation_In_Cyber_Defense",\
          "Security_Beyond_the_SIEM", "OVUM_Dispelling_the_Myths_Around_DDoS"]
def filterReport(unique):
    unique[unique["Country"].isin(["US", "UK"])] &\
    unique[unique["Most Recent Asset Download"].isin(assets)] &\
    unique[unique["Role"].isin(["IT", "Secutiry"])]