import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
import time
import openpyxl

# For testing - set these variables so you won't get prompted
global labor_path
global materials_path
global pay_path
global cost_path
labor_path = 'D:/Development/Evan/E-1-Billing-App/Labor 05 03 20 - 06 05 20.csv'
materials_path = 'D:/Development/Evan/E-1-Billing-App/Materials 05 03 20 - 06 05 20.csv'
pay_path = 'D:/Development/Evan/E-1-Billing-App/E1 Employee Pay Rates - Sheet1.csv'
cost_path = 'D:/Development/Evan/E-1-Billing-App/E1 Cost Codes - Sheet1.csv'

def getLabor ():
    global df_l

    if 'labor_path' in globals():
        import_file_path = labor_path
    else: 
        import_file_path = filedialog.askopenfilename()
    
    df_l = pd.read_csv (import_file_path)


def getMaterials ():
    global df_m

    if 'materials_path' in globals():
        import_file_path = materials_path
    else: 
        import_file_path = filedialog.askopenfilename()
    df_m = pd.read_csv (import_file_path)


def getEmployees():
    global df_p 
    
    if 'pay_path' in globals():
        import_file_path = pay_path
    else: 
        import_file_path = filedialog.askopenfilename()
    df_p = pd.read_csv (import_file_path)

def getCost():
    global df_c 
    
    if 'cost_path' in globals():
        import_file_path = cost_path
    else: 
        import_file_path = filedialog.askopenfilename()
    df_c = pd.read_csv (import_file_path)

def convertToExcel ():
    global read_file

    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    read_file.to_excel (export_file_path, index = None, header=True)

#saveAsButton_Excel = tk.Button(text='Convert CSV to Excel', command=convertToExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
#canvas1.create_window(150, 180, window=saveAsButton_Excel)

def createApplication():
    getLabor()
    getMaterials()
    getCost()
    getEmployees()

    global df_l
    global df_m
    global df_c
    global df_p

    MsgBox = tk.messagebox.askquestion ('Create New Billing Folder',"This will create a new folder with today's date.",icon = 'warning')
    if MsgBox == 'yes':
        TodaysDate = time.strftime("%m-%d-%Y")
        outdir = filedialog.askdirectory() + '\\' + TodaysDate +' Billing Files'
        if not os.path.exists(outdir):
            os.mkdir(outdir)

        # Drop unwanted columns
        df_l = df_l.drop(columns=['payroll_id','fname','lname','number','group','local_day','local_end_time','tz','location'])

        # Split job code and cost code columns into new column sets
        new = df_l["jobcode"].str.split("-", n = 1, expand = True)
        new2 = df_l['cost code'].str.split('-', n = 1, expand = True)

        df_l['Job No'] = new[1]
        df_l['Job Description'] = new[0]
        df_l['Cost Code'] = new2[0]
        df_l['Cost Code Description'] = new2[1]

        # rename columns/create new columns
        df_l = df_l.rename(columns={'local_date': 'Date','hours':'Cost/Hours','username':'Vendor/Employee'})
        df_l['Class'] = "LAB"
        df_l['Cost/Hours'] = pd.to_numeric(df_l['Cost/Hours'],errors='coerce')
        df_l['Type'] = ""
        df_l['Billable'] = ""
        df_l['Billable'] = pd.to_numeric(df_l['Billable'],errors='coerce')
    
        # drop residual 'jobcode' column
        df_l = df_l.drop(columns=['jobcode'])

        df_lp = pd.merge(df_l,df_p, how = 'left')

        # column schema
        df_lp = df_lp[['Job No','Job Description','Cost Code','Cost Code Description','Date','Class','Cost/Hours','Rate','Billable','Vendor/Employee','notes']]
       
       # multiply data
        df_lp['Billable']=df_lp['Rate']*df_lp['Cost/Hours']

        # Create new path w/ date
        outname = 'LABOR.csv'

        fullname = os.path.join(outdir, outname)

        df_lp.to_csv(fullname, index=False)
        # drop unnecessary 'notes' column
        df_lp.drop(columns='notes')

        # Drop unwanted columns
        df_m = df_m.drop(columns=['Geographic Area','Phase No','Phase Description','Source','Category','Hours/Units','Quantity','Type'])

        # rename columns/create new columns
        df_m = df_m.rename(columns={'Dollars': 'Cost/Hours','Comment':'Vendor/Employee'})
        df_m['Cost/Hours'] = pd.to_numeric(df_m['Cost/Hours'],errors='coerce')
        df_m['Billable'] = ""
        df_m['Billable'] = pd.to_numeric(df_m['Billable'],errors='coerce')

        df_mc = pd.merge(df_m,df_c, how = 'left')
       
        #column schema
        df_mc = df_mc[['Job No','Job Description','Cost Code','Cost Code Description','Date','Class','Cost/Hours','Rate','Billable','Vendor/Employee']]

        # multiply data
        df_mc['Billable']=df_mc['Rate']*df_mc['Cost/Hours']

        outname = 'MATERIALS.csv'

        fullname = os.path.join(outdir, outname)

        df_mc.to_csv(fullname, index=False)

        # Append the two and make a Master file
        compiled = df_lp.append(df_mc)

        # 'notes' no longer needed
        compiled = compiled.drop(columns = ['notes'])

        compiled = compiled[['Job No','Job Description','Cost Code','Cost Code Description','Date','Class','Cost/Hours','Rate','Billable','Vendor/Employee']]

        outname = TodaysDate +" MASTER Billing"+".xlsx"
        sheetname = " MasterSheet.csv"

        fullname = os.path.join(outdir, outname)
        sheethand = os.path.join(outdir, sheetname)


        compiled.to_excel(fullname, sheet_name='Billing', index=False)
        compiled.to_csv(sheethand, index=False)

        print('New folder and spreadsheets generated!')


def exitApplication():
    MsgBox = tk.messagebox.askquestion ('Exit Application','Are you sure you want to exit the application',icon = 'warning')
    if MsgBox == 'yes':
        root.destroy()


root= tk.Tk()
root.title('Billing Application')
root.geometry('300x600')
root.configure(bg='lightsteelblue2')


label1 = tk.Label(root, text='Billing Application', bg = 'lightsteelblue2', anchor='center')
label1.config(font=('helvetica', 20))
label1.grid(row=0)
browseButton_Labor = tk.Button(root, text="      Import Labor CSV File     ", command=getLabor, bg='green', fg='white', font=('helvetica', 12, 'bold'))
browseButton_Labor.grid(row=1)

browseButton_Materials = tk.Button(root, text="      Import Materials CSV File     ", command=getMaterials, bg='green', fg='white', font=('helvetica', 12, 'bold'))
browseButton_Materials.grid(row=2)

browseButton_Employee = tk.Button(root, text="      Import Pay Rates CSV File     ", command=getEmployees, bg='green', fg='white', font=('helvetica', 12, 'bold'))
browseButton_Employee.grid(row=3)

browseButton_Cost = tk.Button(root, text="      Import Cost Codes CSV File     ", command=getCost, bg='green', fg='white', font=('helvetica', 12, 'bold'))
browseButton_Cost.grid(row=4)

createButton = tk.Button (root, text='       Create New Billing Folder     ',command=createApplication, bg='blue', fg='white', font=('helvetica', 12, 'bold'))
createButton.grid(row=5)

exitButton = tk.Button (root, text='       Exit Application     ',command=exitApplication, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
exitButton.grid(row=6)
root.mainloop()
