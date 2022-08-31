# Import Module
import os
import csv
import pandas as pd
import openpyxl
import numpy as np

filedir = os.path.dirname(__file__)  # File dir of .py file

main_folder = os.path.join(filedir, 'inventory')  # floor´s main folder

paths_in_folder = [os.path.join(main_folder, i) for i in os.listdir(main_folder)]  # subfolders

full_inf = []  # Array for Informations

for folder in paths_in_folder:
    print("Current Folder: " + folder)
    for file in os.listdir(folder):
        #print("Reading: ", file)
        file_direction = os.path.join(folder, file)
        with open(file_direction, encoding="utf-16-le") as f: #You can change encoding
            text_get = f.read()
            #print(text_get)
            informations = text_get.split("\n")
            folder_name_split = str(folder).split("\\")
            name_num = 0
            while folder_name_split[name_num] != "inventory":
                name_num +=1
            folder_name = folder_name_split[name_num + 1] #name_of_folder
            process = informations[7]
            processes = process.split(' ')
            ip_full = informations[9]
            ip=ip_full.split('"')
            #print("Test Protokol 1: "+str(ip[1]))
            if(ip[0] != "{" or ip[0]== " "):
                ipadress = "empty"
            else:
                ipadress = ip[1]
            #print("Test Protokol 2: " + ipadress)
            inf_temp = [folder_name, informations[1], informations[3], informations[5], ipadress, processes[2], processes[5]]
            print(informations)
            full_inf.append(inf_temp)
    #print(full_inf)

#header = ['FLOOR', 'COMPNAME', 'COMPID', 'SERIALNUMBER','IPADRESS', 'ISLEMCI', 'HIZ']
fileName = 'envanter.csv'
df = pd.DataFrame(full_inf, columns=['FLOOR', 'COMPNAME', 'COMPID', 'SERIALNUMBER','IPADRESS', 'PROCESS', 'GHZ'])
print(df)
df.to_excel('inventory.xlsx', sheet_name='inventory',index=False)