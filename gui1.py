import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
import time
import openpyxl


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



def getCSV2 ():
    global file2

    import_file_path = filedialog.askopenfilename()
    file2 = pd.read_csv (import_file_path)


def getEmployees():
    global employees 
    employees = dict()
    import_file_path = filedialog.askopenfilename()
    df_e = pd.read_csv (import_file_path)


    for index, row in df_e.iterrows():
        employees[row['employee']] = row['rate']

def getCost():
    global cost 
    cost = dict()

    import_file_path = filedialog.askopenfilename()
    df_c = pd.read_csv (import_file_path)

    for x in df_c:
        print (x);
    for index, row in df_c.iterrows():
        cost[row['cost code']] = row['multiplier']

def convertToExcel ():
    global read_file

    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    read_file.to_excel (export_file_path, index = None, header=True)

#saveAsButton_Excel = tk.Button(text='Convert CSV to Excel', command=convertToExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
#canvas1.create_window(150, 180, window=saveAsButton_Excel)

def createApplication():
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

       for index, row in df2.iterrows():
          if row['vendor/employee'] in employees and row['cost code'] in cost:
            row['rate'] = employees[row['vendor/employee']] * cost[row['cost code']]


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
browseButton_CSV = tk.Button(root, text="      Import Labor CSV File     ", command=getCSV1, bg='green', fg='white', font=('helvetica', 12, 'bold'))
browseButton_CSV.grid(row=1)

browseButton_CSV2 = tk.Button(root, text="      Import Materials CSV File     ", command=getCSV2, bg='green', fg='white', font=('helvetica', 12, 'bold'))
browseButton_CSV2.grid(row=2)

browseButton_CSV = tk.Button(root, text="      Import Employee CSV File     ", command=getEmployees, bg='green', fg='white', font=('helvetica', 12, 'bold'))
browseButton_CSV.grid(row=3)

browseButton_CSV2 = tk.Button(root, text="      Import Cost CSV File     ", command=getCost, bg='green', fg='white', font=('helvetica', 12, 'bold'))
browseButton_CSV2.grid(row=4)

createButton = tk.Button (root, text='       Create New Billing Folder     ',command=createApplication, bg='blue', fg='white', font=('helvetica', 12, 'bold'))
createButton.grid(row=5)

exitButton = tk.Button (root, text='       Exit Application     ',command=exitApplication, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
exitButton.grid(row=6)
root.mainloop()
