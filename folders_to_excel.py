import os
import sys
import pandas as pd
dir_path = os.path.dirname(os.path.realpath('__file__'))
datadir  = os.path.dirname(sys.executable)
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

filepath = r'C:\Users\#\Desktop\AMM - Tech Deals by region'
current_dir = os.chdir(filepath)
templatefilename = os.path.join(current_dir,"Master.xlsx")

def running_function (templatefilename=templatefilename, current_dir=current_dir):   
    temp_dir = os.path.join(current_dir,'AMM - Tech Deals by region')
    folder_list = [x[0] for x in os.walk(temp_dir)]
    folder_list = folder_list[1:len(folder_list)]
    # run the code for every folder in the directory
    exceptions = 0
    row_idx = 1
    df_0 = pd.read_excel(templatefilename,sheet_name=0)
    for item in folder_list:
        print (item)    
        dirname = os.path.basename(item) ###TEMPORARY
        print (dirname)
        # This part will be used to find all the folders
        path = item
        full_path_list = [os.path.join(path,f) for\
                         f in os.listdir(path) if os.path.isfile(os.path.join(path,f)) ]        
        project_idx = 1
        row_idx = row_idx + 1 
        for filename in full_path_list:
            row_idx+= 1
            try:
                df_xlsx = pd.read_excel(filename,sheet_name=0) # sep=";"
                df_xlsx['Filename'] = filename
                df_xlsx['Dirname']  = dirname
                #qui dovrei aggiungere le altre info --> nome file
                df_0 = df_0.append(df_xlsx.copy(),sort=False)
            except:
                exceptions = exceptions + 1
                print("didn't work as wished")
    print("exceptions: "+str(exceptions))
    print("total files: "+str(row_idx))
    return df_0

df = running_function (templatefilename=templatefilename, current_dir=current_dir)
writer = pd.ExcelWriter('FinalOutput.xlsx')
df.to_excel(writer,'Sheet1')
