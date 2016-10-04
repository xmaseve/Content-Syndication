# -*- coding: utf-8 -*-
"""
Created on Mon Oct 03 12:53:00 2016

@author: Qi Yi
"""

import pandas as pd

#two ways to read excel file in python
#xl = pd.ExcelFile("C:\\Users\\Qi Yi\\Desktop\inSegment - Arbor Networks - Weekly Lead Report 7.7.2016.xlsx")
#leads = xl.parse("All Leads") #sheet name
leads = pd.read_excel('C:\\Users\\Qi Yi\\Desktop\inSegment - Arbor Networks - Weekly Lead Report 7.7.2016.xlsx', sheetname="All Leads")

unique = leads.drop_duplicates()
len(unique) #16153

flag = unique[unique.duplicated(["Email Address"], keep=False)]
flag = flag.sort_values(["Email Address"], ascending=True)
flag.to_excel("C:\\Users\\Qi Yi\\Desktop\\flag.xlsx")     

assets = ["DDoS_Security_Myths_Busted", "SecCurrent_The_Hunted_Becomes_the_Hunter",\
          "ITHarvest_Report_Security_Analytics__A_Required_Escalation_In_Cyber_Defense",\
          "Security_Beyond_the_SIEM", "OVUM_Dispelling_the_Myths_Around_DDoS"]

industry = ["Banking/Financial Services", "Cloud Hosting Provider", "Computer Hardware/Software",\
            "Entertainment", "Federal/National Government", "Healthcare/Hospitals",\
            "Hospitality/Leisure", "Insurance", "Internet Service Provider", "Logistics/Transportation",\
            "Managed Security Service Provider", "Media/Publishing", "Mobile Provider",\
            "Online Gaming", "Pharmaceutical/Medical Devices", "Retail/Online Services", "Utilities"]          

            
def filterReport(unique):
    a = unique[unique["Country"].isin(["US", "US of America", "UK"])] 
    b = a[a["Most Recent Asset Download"].isin(assets)] 
    c = b.dropna(subset=["Role"]) #Role columns has NAs.
    d = c[c["Role"].str.contains('IT|Security')]
    e = d[d["Industry"].isin(industry)]
    f = e[e["Job Title"].str.contains('VP|Director|Manager|CTO|Chief|CIO|Vice|COO|CRO|CEO|\
         CCO|CFO|CISO|C Level|V.P')]
    g = f[f["Country"].isin(["US", "US of America"])]
    us = g[g["Number Of Employees"].isin(["1000-4999", "5000-10000", "10001 or More"])]
    i = f[f["Country"].isin(["UK"])]
    uk = i[i["Number Of Employees"].isin(["500-999", "1000-4999", "5000-10000", "10001 or More"])]
    result = pd.concat([us, uk])
    return result
              
result = filterReport(unique)    
    
result.to_excel("C:\\Users\\Qi Yi\\Desktop\\result.xlsx") 
