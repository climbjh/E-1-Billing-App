import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
import time
import openpyxl

root= tk.Tk()

canvas1 = tk.Canvas(root, width = 300, height = 350, bg = 'lightsteelblue2', relief = 'raised')
canvas1.pack()

label1 = tk.Label(root, text='Billing Application', bg = 'lightsteelblue2')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)

def getCSV1 ():
    global file1

    import_file_path = filedialog.askopenfilename()
    file1 = pd.read_csv (import_file_path)
    #MsgBox = tk.messagebox.askquestion ('File Selection','Is this the labor file?',(text=tk.path.basename(file1), fg="blue"))
    #if MsgBox == 'yes':
    #    file1 = file1
    #else:
    #    import_file_path = filedialog.askopenfilename()
    #    file1 = pd.read_csv (import_file_path)

browseButton_CSV = tk.Button(text="      Import Labor CSV File     ", command=getCSV1, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_CSV)
#browseButton_CSV.grid(row = 1)

def getCSV2 ():
    global file2

    import_file_path = filedialog.askopenfilename()
    file2 = pd.read_csv (import_file_path)

browseButton_CSV2 = tk.Button(text="      Import Materials CSV File     ", command=getCSV2, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 180, window=browseButton_CSV2)
#browseButton_CSV2.grid(row = 2)

def convertToExcel ():
    global read_file

    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    read_file.to_excel (export_file_path, index = None, header=True)

#saveAsButton_Excel = tk.Button(text='Convert CSV to Excel', command=convertToExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
#canvas1.create_window(150, 180, window=saveAsButton_Excel)

def createApplication():
    MsgBox = tk.messagebox.askquestion ('Create New Billing Folder',"This will create a new folder with today's date.",icon = 'warning')
    if MsgBox == 'yes':
       # Open Labor file
       df = file1

       # Drop unwanted columns
       df = df.drop(columns=['payroll_id','fname','lname','number','group','local_day','local_end_time','tz','location'])

       # Split job code and cost code columns into new column sets
       new = df["jobcode"].str.split("-", n = 1, expand = True)
       new2 = df['cost code'].str.split('-', n = 1, expand = True)

       df['Job No'] = new[1]
       df['Job Description'] = new[0]
       df['Cost Code'] = new2[0]
       df['Cost Code Description'] = new2[1]

       # rename columns/create new columns
       df = df.rename(columns={'local_date': 'Date','hours':'Cost/Hours','username':'Vendor/Employee'})
       df['Class'] = "LAB"
       df['Rate'] = ""
       #df['rate'] = df['rate'].astype(float)
       df['Cost/Hours'] = df['Cost/Hours'].astype(float)
       df['Billable'] = "" # df['rate']*df['cost/hours']
       df['Type'] = ""

       # drop residual jobcode column
       df = df.drop(columns=['jobcode'])

       # column schema
       df = df[['Job No','Job Description','Cost Code','Cost Code Description','Date','Class','Cost/Hours','Rate','Billable','Vendor/Employee','notes']]

       # Create new path w/ date
       TodaysDate = time.strftime("%m-%d-%Y")
       outname = 'LABOR.csv'

       outdir = r'C:\Users\evanj\Desktop\E1 Project\ '+TodaysDate+' Billing Files'
       if not os.path.exists(outdir):
           os.mkdir(outdir)

       fullname = os.path.join(outdir, outname)

       df.to_csv(fullname, index=False)

       # drop unnecessary 'notes' column
       df.drop(columns='notes')



       # Open and convert Materials File
       df2 = file2

       # Drop unwanted columns
       df2 = df2.drop(columns=['Geographic Area','Phase No','Phase Description','Source','Category','Hours/Units','Quantity','Type'])

       # rename columns/create new columns
       df2 = df2.rename(columns={'Dollars': 'Cost/Hours','Comment':'Vendor/Employee'})
       df2['Rate'] = ""
       #df['rate'] = df['rate'].astype(float)
       df['Cost/Hours'] = df['Cost/Hours'].astype(float)
       df2['Billable'] = "" # df2['rate']*df2['cost/hours']

       #column schema
       df2 = df2[['Job No','Job Description','Cost Code','Cost Code Description','Date','Class','Cost/Hours','Rate','Billable','Vendor/Employee']]

       outname = 'MATERIALS.csv'

       outdir = r'C:\Users\evanj\Desktop\E1 Project\ '+TodaysDate+' Billing Files'
       if not os.path.exists(outdir):
           os.mkdir(outdir)

       fullname = os.path.join(outdir, outname)

       df2.to_csv(fullname, index=False)

       # Append the two and make a Master file
       compiled = df.append(df2)

       # 'notes' no longer needed
       compiled = compiled.drop(columns = ['notes'])

       outdir = r'C:\Users\evanj\Desktop\E1 Project\ '+TodaysDate+' Billing Files'
       if not os.path.exists(outdir):
           os.mkdir(outdir)

       TodaysDate = time.strftime("%m-%d-%Y")
       outname = TodaysDate +" MASTER Billing"+".xlsx"
       sheetname = " MasterSheet.csv"

       fullname = os.path.join(outdir, outname)
       sheethand = os.path.join(outdir, sheetname)


       compiled.to_excel(fullname, sheet_name='Billing', index=False)
       compiled.to_csv(sheethand, index=False)

       print('New folder and spreadsheets generated!')

createButton = tk.Button (root, text='       Create New Billing Folder     ',command=createApplication, bg='blue', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 230, window=createButton)
#createButton.grid(row = 3)

def exitApplication():
    MsgBox = tk.messagebox.askquestion ('Exit Application','Are you sure you want to exit the application',icon = 'warning')
    if MsgBox == 'yes':
       root.destroy()

exitButton = tk.Button (root, text='       Exit Application     ',command=exitApplication, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 280, window=exitButton)
#exitButton.grid(row = 4)

root.mainloop()
