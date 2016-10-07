# -*- coding: utf-8 -*-
"""
Created on Mon Oct 03 12:53:00 2016

@author: Qi Yi
"""

import pandas as pd

#two ways to read excel file in python
#xl = pd.ExcelFile("C:\\Users\\Qi Yi\\Desktop\inSegment - Arbor Networks - Weekly Lead Report 7.7.2016.xlsx")
#leads = xl.parse("All Leads") #sheet name
leads = pd.read_excel('C:\\Users\\Qi Yi\\Desktop\inSegment - Arbor Networks - Weekly Lead Report 7.7.2016 test.xlsx', sheetname="This Batch")   

assets = ["DDoS_Security_Myths_Busted", "SecCurrent_The_Hunted_Becomes_the_Hunter",\
          "ITHarvest_Report_Security_Analytics__A_Required_Escalation_In_Cyber_Defense",\
          "Security_Beyond_the_SIEM", "OVUM_Dispelling_the_Myths_Around_DDoS"]

industry = ["Banking/Financial Services", "Cloud Hosting Provider", "Computer Hardware/Software",\
            "Entertainment", "Federal/National Government", "Healthcare/Hospitals",\
            "Hospitality/Leisure", "Insurance", "Internet Service Provider", "Logistics/Transportation",\
            "Managed Security Service Provider", "Media/Publishing", "Mobile Provider",\
            "Online Gaming", "Pharmaceutical/Medical Devices", "Retail/Online Services", "Utilities"]          

            
def filterReport(leads):
    a = leads[leads["Country"].isin(["US", "US of America", "UK"])] 
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
              
result = filterReport(leads)        
#result.to_excel("C:\\Users\\Qi Yi\\Desktop\\result.xlsx") 

'''
Locate missing values
'''
newleads = result[result.columns[[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18]]]
mvalues = result[newleads.isnull().any(axis=1)]
mvalues.to_excel("C:\\Users\\Qi Yi\\Desktop\\missing values.xlsx")


'''
Find unique values
'''
unique_email = result.drop_duplicates(["Email Address"], keep=False)
len(unique_email) 
unique_email.to_excel("C:\\Users\\Qi Yi\\Desktop\\unique_email.xlsx")

'''
Find duplicates
'''
duplicate_email = result[result.duplicated(["Email Address"], keep=False)]
len(duplicate_email) 
duplicate_email = duplicate_email.sort_values(["Email Address"])
duplicate_email.to_excel("C:\\Users\\Qi Yi\\Desktop\\duplicate_email.xlsx") 

group_email = duplicate_email.groupby(["Email Address"]).size()
group_email[group_email > 2]
a = duplicate_email.groupby(["Email Address", "Most Recent Asset Download"]).size()
duplicate_email[a == 1]

#flag = flag.sort_values(["Email Address"], ascending=True)
same = duplicate_email[duplicate_email.duplicated(["Lead Source Third Party", "Most Recent Asset Download"], keep=False)]
len(same)
same["Last Interesting Moment Date"] = pd.to_datetime(same["Last Interesting Moment Date"])                       
same = same.sort_values(["Email Address", "Last Interesting Moment Date"])     
b = same.groupby("Email Address")  
c = b["Last Interesting Moment Date"].diff()
same["Days"] = c
same["Days"] = same["Days"].fillna(method='bfill')
same_g90 = same[same["Days"] > "90 days"]
same_l90 = same[same["Days"] < "90 days"]
len(same_g90)
len(same_l90)                          
same_g90.to_excel("C:\\Users\\Qi Yi\\Desktop\\same_G90.xlsx") 
same_l90.to_excel("C:\\Users\\Qi Yi\\Desktop\\same_L90.xlsx")                          
          
                       
different = duplicate_email.drop_duplicates(["Lead Source Third Party","Most Recent Asset Download"], keep=False)
len(different)
different.to_excel("C:\\Users\\Qi Yi\\Desktop\\different venders and assets.xlsx")     
