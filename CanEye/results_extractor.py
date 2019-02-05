"""
The purpose of the script is to the following:
--> CanEye softeware was used to generate excel data for multiple image sets, each excel file was placed in a directory (time) and many subdirectories (person, and other junk ones).
--> The Post-Doc researcher I was working with told me to go through the files one by one to extract 9 key data points and put them all in one excel file with pages for each time.
--> The 9 data points are: 4 FAPAR values, 3 PAI - ALA values, the directory name (time), the subdirectory name (person).
    --> Note that directory and subdirectory names are not from the excel file itsself.
--> There were over 70 files so this work would be extremely tedious to do by hand.
--> to overcome this I decided to automate the processes.
--> This script also had to be run on OTHER people's CANEYE data so it had to be generalizable to over 500 different excel files from different people.
--> This script had to be given to a person with no computer science training, so it had to be simple for them to run.
--> Lots of 'print' statements are included so non-computer science users can know whats happening just from reading the terminal.
--> To get this to run on Windows simply: SHIFT + RIGHT CLICK inside the main containing the .py file, PRESS 'open PowerShell window here', TYPE & ENTER: python results_extractor.py
--> For the sake of keeping our lab data confidential, I will make excel files with fake values.


This is what the script is doing: 
--> Go through each excel files in the directory tree, extract 4 key values from each excel file, put all data in one main excel file where each sheet is a different time period


This is how the directories are structured in the main CANEYE folder (Junk is just other files/folders that you don't need to know about, it can represent one or multiple files/folders | [] represent many files/folders of that category):
--> | CANEYE \\ Fisheye, Junk \\ [Time folders] \\ [Person Folders] \\  Junk  \\ data.xlsx, Junk
"""

from openpyxl import load_workbook, Workbook
import pandas as pd
import os


#The BASE_DIR is the directory of where-ever the .py file is saved
BASE_DIR= os.path.dirname(os.path.abspath(__file__))
image_dir= os.path.join(BASE_DIR, 'fisheye')

#2 main variables of this script
excel_files=[]
excel_data=[]

#Excel file directory extraction: Os.walk goes through each subfolder in the main subfolder and returns the root, directories and files. I then take the excel file name (file) and join it with its directory name (root)
for root, dirs, files in os.walk(image_dir):
    for file in files:
        if file.lower().endswith('xlsx'):
            excel_files.append(os.path.join(root,file))
print(f'Found over {len(excel_files)} excel files, now collecting data, this may take a while...')

#Excel file data extraction: For each excel file path, get the time value (dirs), person value (subdirs). Then enter each excel file using the openpyxl library to extract 7 data points. Append all the 9 values to excel_data list.
for excel_file in excel_files:
    for dirs in os.listdir(image_dir):
        if dirs in excel_file:
            for subdirs in (os.listdir(os.path.join(image_dir,dirs))):
                if subdirs in excel_file:
                    wb= load_workbook(filename=excel_file)
                    data = [dirs,subdirs,
                            wb['PAI, ALA']['B6'].value, wb['PAI, ALA']['B7'].value, wb['PAI, ALA']['B8'].value,
                            wb['FAPAR']['B2'].value, wb['FAPAR']['B3'].value, wb['FAPAR']['B4'].value, wb['FAPAR']['B5'].value]
                    excel_data.append(data)
print('Data collection complete, now saving...')

#Saving data to an excel file (.xlsx): For each unique time value, make a page with the 9 key values.
df= pd.DataFrame(excel_data, columns=['Time', 'Person',
                                      'LAI2000, 3 rings', 'LAI2000, 4 rings', 'LAI2000, 5 rings',
                                      'Measured Direct FAPAR', 'Modeled Direct FAPAR', 'Measured Diffuse FAPAR', 'Modeled Direct FAPAR'])
writer = pd.ExcelWriter('results.xlsx', engine='xlsxwriter')
for times in list(set(df['Time'])):
    (df.loc[df['Time'] == times]).to_excel(writer, sheet_name=times, index= False)
writer.save()
print('Done saving file as "results.xlsx", all processes complete.')
print('If you are to repeat the script all previous data will be overwritten.')
print('You ran into an error and an excel file was made, delete it. Read the README.txt file, to see what went wrong. If its still not working contact me.')
                   
        

