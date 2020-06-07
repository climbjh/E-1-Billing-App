import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
import time
import openpyxl

root= tk.Tk()

canvas1 = tk.Canvas(root, width = 300, height = 450, bg = 'lightsteelblue2', relief = 'raised')
canvas1.pack()

label1 = tk.Label(root, text='Billing Application', bg = 'lightsteelblue2')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)

def getCSV1 ():
    global file1

    import_file_path = filedialog.askopenfilename()
    file1 = pd.read_csv (import_file_path)

browseButton_CSV = tk.Button(text="      Import LABOR CSV File     ", command=getCSV1, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_CSV)
#browseButton_CSV.grid(row = 1)

def getCSV2 ():
    global file2

    import_file_path = filedialog.askopenfilename()
    file2 = pd.read_csv (import_file_path)

browseButton_CSV2 = tk.Button(text="      Import MATERIALS CSV File     ", command=getCSV2, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 180, window=browseButton_CSV2)
#browseButton_CSV2.grid(row = 2)

def getCSV3 ():
    global file3

    import_file_path = filedialog.askopenfilename()
    file3 = pd.read_csv (import_file_path)

browseButton_CSV3 = tk.Button(text="      Import PAY RATES CSV File     ", command=getCSV3, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 230, window=browseButton_CSV3)

def getCSV4 ():
    global file4

    import_file_path = filedialog.askopenfilename()
    file4 = pd.read_csv (import_file_path)

browseButton_CSV4 = tk.Button(text="      Import COST CODES CSV File     ", command=getCSV4, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 280, window=browseButton_CSV4)

def convertToExcel ():
    global read_file

    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    read_file.to_excel (export_file_path, index = None, header=True)


def createApplication():
    MsgBox = tk.messagebox.askquestion ('Create New Billing Folder',"This will create a new folder with today's date.",icon = 'warning')
    if MsgBox == 'yes':
       # Open Labor file
       df_lab = file1
       df_mat = file2
       df_pay = file3
       df_cost = file4

       # Drop unwanted columns in df1
       df_lab = df_lab.drop(columns=['payroll_id','fname','lname','number','group','local_day','local_end_time','tz','location'])

       # Split job code and cost code columns into new column sets
       new = df_lab["jobcode"].str.split("-", n = 1, expand = True)
       new2 = df_lab['cost code'].str.split('-', n = 1, expand = True)

       df_lab['Job No'] = new[1]
       df_lab['Job Description'] = new[0]
       df_lab['Cost Code'] = new2[0]
       df_lab['Cost Code Description'] = new2[1]

       # rename columns/create new columns
       df_lab = df_lab.rename(columns={'local_date': 'Date','hours':'Cost/Hours','username':'Vendor/Employee'})
       df_lab['Class'] = "LAB"
       df_lab['Cost/Hours'] = pd.to_numeric(df_lab['Cost/Hours'],errors='coerce')
       df_lab['Billable'] = ""
       df_lab['Billable'] = pd.to_numeric(df_lab['Billable'],errors='coerce')
    
       # drop residual 'jobcode' column
       df_lab = df_lab.drop(columns=['jobcode'])

       df_comp1 = pd.merge(df_lab,df_pay, how = 'left')

       # column schema
       df_comp1 = df_comp1[['Job No','Job Description','Cost Code','Cost Code Description','Date','Class','Cost/Hours','Rate','Billable','Vendor/Employee','notes']]
       
       # multiply data
       df_comp1['Billable']=df_comp1['Rate']*df_comp1['Cost/Hours']

       # Create new path w/ date
       TodaysDate = time.strftime("%m-%d-%Y")
       outname = 'LABOR.csv'

       outdir = r'C:\Users\evanj\Desktop\E1 Project\ '+TodaysDate+' Billing Files'
       if not os.path.exists(outdir):
           os.mkdir(outdir)

       fullname = os.path.join(outdir, outname)

       df_comp1.to_csv(fullname, index=False)

       # drop unnecessary 'notes' column
       df_comp1.drop(columns='notes')

      # Drop unwanted columns in df2
       df_mat = df_mat.drop(columns=['Geographic Area','Phase No','Phase Description','Source','Category','Hours/Units','Quantity','Type'])

       # rename columns/create new columns
       df_mat = df_mat.rename(columns={'Dollars': 'Cost/Hours','Comment':'Vendor/Employee'})
       df_mat['Cost/Hours'] = pd.to_numeric(df_mat['Cost/Hours'],errors='coerce')
       df_mat['Billable'] = ""
       df_mat['Billable'] = pd.to_numeric(df_mat['Billable'],errors='coerce')

       df_comp2 = pd.merge(df_mat,df_cost, how = 'left')
       
       # column schema
       df_comp2 = df_comp2[['Job No','Job Description','Cost Code','Cost Code Description','Date','Class','Cost/Hours','Rate','Billable','Vendor/Employee']]

       # multiply data
       df_comp2['Billable']=df_comp2['Rate']*df_comp2['Cost/Hours']

       outname = 'MATERIALS.csv'

       outdir = r'C:\Users\evanj\Desktop\E1 Project\ '+TodaysDate+' Billing Files'
       if not os.path.exists(outdir):
           os.mkdir(outdir)

       fullname = os.path.join(outdir, outname)

       df_comp2.to_csv(fullname, index=False)

       # Append the two and make a Master file
       master = df_comp1.append(df_comp2)

       # 'notes' no longer needed
       master = master.drop(columns = ['notes'])

       master = master[['Job No','Job Description','Cost Code','Cost Code Description','Date','Class','Cost/Hours','Rate','Billable','Vendor/Employee']]

       outdir = r'C:\Users\evanj\Desktop\E1 Project\ '+TodaysDate+' Billing Files'
       if not os.path.exists(outdir):
           os.mkdir(outdir)

       TodaysDate = time.strftime("%m-%d-%Y")
       outname = TodaysDate +" MASTER Billing"+".xlsx"
       sheetname = " MasterSheet.csv"

       fullname = os.path.join(outdir, outname)
       sheethand = os.path.join(outdir, sheetname)


       master.to_excel(fullname, sheet_name='Billing', index=False)
       master.to_csv(sheethand, index=False)

       print('New folder and spreadsheets generated!')

createButton = tk.Button (root, text='       Create New Billing Folder     ',command=createApplication, bg='blue', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 330, window=createButton)
#createButton.grid(row = 3)

def exitApplication():
    root.destroy()

exitButton = tk.Button (root, text='       Exit Application     ',command=exitApplication, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 380, window=exitButton)
#exitButton.grid(row = 4)

root.mainloop()
