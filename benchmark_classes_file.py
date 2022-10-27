import openpyxl
import pandas as pd
from openpyxl import Workbook,load_workbook
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import filedialog
root = tk.Tk()
root.withdraw()
print("Select class file")
class_file = filedialog.askopenfilename()
teacher_file = 'Classes.xlsx'

#print(ws)
#print(ws['A1'].value) # Read value at A1

#ws['M2'].value = "TEST" # Change value at A2
#wb.save('Classes.xlsx')

# for row in range (1,900): #stops at row 10
#     for col in range (1,7): #columns 1 -4
#         char = get_column_letter(col)
#         print(ws[char + str(row)].value)


# Change the Kindergarten Grade code from KF or 25 to K.
def fixKindergartenGradeLevel(file):
    wb = load_workbook(file) #Load Workbook
    ws = wb.active #Worksheet
    for row in range (1,900): #stops at row 900
        for col in range (6,7): # column J
            char = get_column_letter(col)
            if ws[char + str(row)].value == "KF" or ws[char + str(row)].value == "25":
                ws[char + str(row)].value = 'K'
    wb.save(file)
            

# Format the columns of class sheet to Benchmark format.
def formatSheet(file):
    wb = load_workbook(file) #Load Workbook
    ws = wb.active #Worksheet
    #print(ws)
    
    #Concat Class Name with Last Name
    for row in range (1,900): #stops at row 10
     for col in range (3,4): #columns 1 -4
        char = get_column_letter(col)
        try:
            ws[char + str(row)].value = ws[get_column_letter(2) + str(row)].value + "-" +ws[get_column_letter(7) + str(row)].value
            print(ws[char + str(row)].value)
        except:
            print("error")
    
    for row in range (2,900): #stops at row 10
     for col in range (5,6): #columns 1 -4
        char = get_column_letter(col)
        char2 = get_column_letter(6)
        
       
        if(ws[char2 + str(row)].value):
            ws[char + str(row)].value = "STUDENT"
        else:
             print("ROW",row)
             print(ws[char2 + str(row)].value)

    #change Column c
    ws['C1'].value = "Class's SIS Id"
    wb.save(file)

# Attach the teachers to the end of the class file.
def pastInTeachers(file,file2):
    wb = load_workbook(file) #Load Workbook
    ws = wb.active #Worksheet
    last_empty_row = len(list(ws.rows))
    print(last_empty_row) #809 migh tneed to plus one this
    # Append or just hard code it in. 
    #Concat to data frames that share common columns
    # load in the various tables from an excel document
    file = pd.read_excel(file,sheet_name='QRY801')
    file2 = pd.read_excel('Classes.xlsx',sheet_name='Teachers')
    merge = pd.concat([file,file2],axis=0) 
    print(merge)
    merge.to_excel('classes_appened.xlsx',index=False)

# Delete Extra columns in Newly appended sheet.
def deleteExtraColumns(file):
    wb = load_workbook(file) #Load Workbook
    ws = wb.active #Worksheet
    ws.delete_cols(7)
    wb.save(file)

#Main Program 
if __name__ == "__main__":
    print("Starting...")
    #fixKindergartenGradeLevel(class_file)
    formatSheet(class_file)
    pastInTeachers(class_file,teacher_file)
    deleteExtraColumns('classes_appened.xlsx')
    print("done")
    #mergeSheets(student_file,clever_file)
    #formatSheet('merge.xlsx')
    #cleverData = pd.read_excel('merge.xlsx',sheet_name='Sheet1')
    #cleverData.to_csv("students_IMPORT.csv",index=False)
    #print("finished")
