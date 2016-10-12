# -*- coding: utf-8 -*-
"""
Created on Wed Oct 12 09:15:37 2016

@author: Qi Yi
"""

import os
import pandas as pd

def concat_files():
    file_list = os.listdir(r"C:\Users\Qi Yi\Desktop\New folder")
    saved_path = os.getcwd()
    os.chdir(r"C:\Users\Qi Yi\Desktop\New folder")
    
    df = []
    for file_name in file_list:
        data = pd.read_csv(file_name)
        df.append(data)
        
    os.chdir(saved_path)
    df = pd.concat(df)
    return df
    
df = concat_files()

df.to_csv(r"C:\Users\Qi Yi\Desktop\New folder\combined data.csv")