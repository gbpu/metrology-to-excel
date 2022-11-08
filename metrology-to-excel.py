

"""
Created on Thu Aug 11 16:58:05 2022

@author: Guilherme Beppu (TdB)

Extrai dados de multiplos relatórios gerados pela CMM Zeiss 
e exporta para uma planilha de Excel 
"""

from pdfminer.high_level import extract_text
import pandas as pd
#import numpy as np
import os
import time
import natsort 


start = time.process_time()


sheet = r"get_data.xlsx"
param = pd.read_excel(sheet);
dirname = param.iloc[0,1]

file_prefix = param.iloc[1,1]
Measures = param.iloc[2,1:]
len_values = param.iloc[3,1:]


col_names = ["Name", "Value", "Nominal Value", "Upper Allowance", "Lower Allowance", "Deviation"]

def get_values(file, directory = dirname):
    """
    Extrai texto do PDF, seleciona valores de cada medição. 
    """
    
    Values = []
    pdf = extract_text(dirname+file)
    split_ = pdf.split('\n')
    split = []    
    
    for line in split_:
        if len(line)>3:
            split.append(line)
            
    i=999           
    for meas in Measures:
        for line in split:
            if line == meas:
                i = 0
            if i <= len_values[0]:
                if i == 1 and line[-1].isnumeric() == False:
                    #print(line)
                    i = i+4
                else:   
                    if len(Values) == 0:
                        Values.append(line)
                        print(line)
                    else:
                        if line != Values[-1]:
                            Values.append(line)
                            print(line)                        
                i = i+1
      
    return Values



All_values = []
All_files = []
for file in os.listdir(dirname):
    if file.startswith(file_prefix) and file.endswith(".pdf"):
        #print(file)
        All_files.append(file)
        

#for file in sorted(All_files):
#    All_values.append(get_values(file))

All_files = natsort.natsorted(All_files)

for file in All_files:
    All_values.append(get_values(file))
    


#path = r"get_data.ods"

df = pd.DataFrame(All_values, columns = (col_names*len(len_values))[:sum(len_values)+len(len_values)], index=All_files)

with pd.ExcelWriter(sheet, mode='a', if_sheet_exists='replace') as writer:  
    df.to_excel(writer, sheet_name=f'{file_prefix}')

try:
    os.system(f"start EXCEL.EXE {sheet}")
except:
    print("Can't open excel")
        

elapsed_time = time.process_time() - start
print('\nElapsed time: %.2f seconds' % elapsed_time)


