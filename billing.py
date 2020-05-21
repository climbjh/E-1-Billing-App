import pandas as pd

# Open Labor file
df = pd.read_csv('LABOR DEMO - Sheet1.csv')

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

df.to_csv('Labor NEW.csv')




# Open and convert Materials File
df2 = pd.read_csv('MATERIALS DEMO - Sheet1.csv')

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

df2.to_csv('Materials NEW.csv')

# Append the two and make a Master file
compiled = df.append(df2)

# 'notes' no longer needed
compiled = compiled.drop(columns = ['notes'])

compiled.to_csv('Master.csv', index = False)
print(compiled)
