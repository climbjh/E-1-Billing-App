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
global employees_path
global cost_path
labor_path = 'D:/Development/Evan/E-1-Billing-App/LABOR DEMO - Sheet1.csv'
materials_path = 'D:/Development/Evan/E-1-Billing-App/MATERIALS DEMO - Sheet1.csv'
employees_path = 'D:/Development/Evan/E-1-Billing-App/EMPLOYEES.csv'
cost_path = 'D:/Development/Evan/E-1-Billing-App/COST.csv'

def getLabor ():
    global file1

    if 'labor_path' in globals():
        import_file_path = labor_path
    else: 
        import_file_path = filedialog.askopenfilename()
    
    file1 = pd.read_csv (import_file_path)


def getMaterials ():
    global file2

    if 'materials_path' in globals():
        import_file_path = materials_path
    else: 
        import_file_path = filedialog.askopenfilename()
    file2 = pd.read_csv (import_file_path)


def getEmployees():
    global employees 
    employees = dict()
    if 'employees_path' in globals():
        import_file_path = employees_path
    else: 
        import_file_path = filedialog.askopenfilename()
    df_e = pd.read_csv (import_file_path)

    for index, row in df_e.iterrows():
        employees[row['employee']] = row['rate']

def getCost():
    global cost 
    cost = dict()

    if 'cost_path' in globals():
        import_file_path = cost_path
    else: 
        import_file_path = filedialog.askopenfilename()
    df_c = pd.read_csv (import_file_path)

    for index, row in df_c.iterrows():
        cost[str(row['cost code'].astype(int))] = row['multiplier']

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

    MsgBox = tk.messagebox.askquestion ('Create New Billing Folder',"This will create a new folder with today's date.",icon = 'warning')
    if MsgBox == 'yes':
        TodaysDate = time.strftime("%m-%d-%Y")
        outdir = filedialog.askdirectory() + '\\' + TodaysDate +' Billing Files'
        if not os.path.exists(outdir):
            os.mkdir(outdir)

        # Open Labor file
        df = file1

        # Drop unwanted columns
        df = df.drop(columns=['other1','other2','other3','other4'])

        # Split job code and cost code columns into new column sets
        new = df["jobcode_1"].str.split("-", n = 1, expand = True)
        new2 = df['cost code'].str.split('-', n = 1, expand = True)

        df['job #'] = new[0]
        df['job description'] = new[1]
        df['cost code'] = new2[0]
        df['cost code description'] = new2[1]

        # rename columns/create new columns
        df = df.rename(columns={'local_date': 'date','hours':'cost/hours','username':'vendor/employee'})
        df['class'] = "LAB"
        df['rate'] = ""
        #df['rate'] = df['rate'].astype(float)
        df['cost/hours'] = df['cost/hours'].astype(float)
        df['billable'] = "" # df['rate']*df['cost/hours']
        df['type'] = ""

        # drop residual jobcode_1 column
        df = df.drop(columns=['jobcode_1'])

        # column schema
        df = df[['job #','job description','cost code','cost code description','date','class','cost/hours','rate','billable','vendor/employee','notes']]

        # Create new path w/ date
        outname = 'LABOR.csv'

        #outdir = r'C:\Users\evanj\Desktop\E1 Project\ '+TodaysDate+' Billing Files'
        
        fullname = os.path.join(outdir, outname)

        df.to_csv(fullname, index=False)




        # Open and convert Materials File
        df2 = file2

        # Drop unwanted columns
        df2 = df2.drop(columns=['other1','other2','other3','other4'])

        # rename columns/create new columns
        df2 = df2.rename(columns={'dollars': 'cost/hours','comments':'vendor/employee'})
        df2['rate'] = ""
        #df['rate'] = df['rate'].astype(float)
        df['cost/hours'] = df['cost/hours'].astype(float)
        df2['billable'] = "" # df2['rate']*df2['cost/hours']

        #column schema
        df2 = df2[['job #','job description','cost code','cost code description','date','class','cost/hours','rate','billable','vendor/employee']]

        for i, row in df.iterrows():
            rate = employees[row['vendor/employee']] * cost[row['cost code']]
            billable = row['cost/hours'] * rate

            df.at[i, 'rate'] = '%.2f'%(rate)
            df.at[i, 'billable'] = '%.2f'%(billable)

        outname = 'MATERIALS.csv'

        fullname = os.path.join(outdir, outname)

        df2.to_csv(fullname, index=False)

        # Append the two and make a Master file
        compiled = df.append(df2)

        # 'notes' no longer needed
        compiled = compiled.drop(columns = ['notes'])

        
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

browseButton_Employee = tk.Button(root, text="      Import Employee CSV File     ", command=getEmployees, bg='green', fg='white', font=('helvetica', 12, 'bold'))
browseButton_Employee.grid(row=3)

browseButton_Cost = tk.Button(root, text="      Import Cost CSV File     ", command=getCost, bg='green', fg='white', font=('helvetica', 12, 'bold'))
browseButton_Cost.grid(row=4)

createButton = tk.Button (root, text='       Create New Billing Folder     ',command=createApplication, bg='blue', fg='white', font=('helvetica', 12, 'bold'))
createButton.grid(row=5)

exitButton = tk.Button (root, text='       Exit Application     ',command=exitApplication, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
exitButton.grid(row=6)
root.mainloop()
