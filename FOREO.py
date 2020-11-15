import pandas as pd
import tkinter as tk
from tkinter import filedialog
import psycopg2
import sys

root= tk.Tk()
canvas1 = tk.Canvas(root, height=500, width=500, bg="#263D42")
canvas1.pack()

def getExcel(): #function to choose file
    global df

    import_file_path = filedialog.askopenfilename()
    df = pd.read_excel(import_file_path, sheet_name=0, skiprows=5, usecols=('C:D, F:O'), dtype=object)
browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'), padx=30,pady=10)
exit_button = tk.Button(text='Done', command=canvas1.quit, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'), padx=30,pady=10)

canvas1.create_window(200, 100, window=browseButton_Excel)
canvas1.create_window(200, 200, window=exit_button)

root.mainloop()

df = df.fillna('0')
print (df)

percentage_cols = df[["LW Var %", "LY Var %", "LYTD Var %"]]
for i in percentage_cols:
    df[i] = df[i].astype(float).mul(100)
    df[i] = df[i].astype(float).map('{:,.0f}'.format)
    df[i] = df[i].astype(str).add('%')


df.to_csv(r'C:\Users\Public\Documents\Generated_CSV_file.csv',encoding='utf-8', index=False, float_format='%.f')

def pg_load_table(file_path, table_name, dbname, host, port, user, pwd):

    #uploading csv file to a target table
    try:
        conn = psycopg2.connect(dbname=dbname, host=host, port=port,
         user=user, password=pwd)
        print("Connected to Database.\n")
        cur = conn.cursor()
        cur.execute("DROP TABLE IF EXISTS foreo")
        sql = '''CREATE TABLE foreo(
           "STORE NO"  TEXT,
           "STORE" text,
           "TY Units" FLOAT,
           "LY Units" FLOAT,
           "TW Sales" FLOAT,
           "LW Sales" FLOAT,
           "LW Var %" TEXT,
           "LY Sales" FLOAT,
           "LY Var %" TEXT,
           "YTD Sales" FLOAT,
           "LYTD Sales" FLOAT,
           "LYTD Var %" TEXT
        )'''
        cur.execute(sql)
        print("Table created successfully!\n")

        f = open(file_path, "r")
        # brisanje iz tablice
        cur.execute("Truncate {} Cascade;".format(table_name))
        print("Deleted data from {}.\n".format(table_name))
        # učitavanje tablice iz CSV datoteke
        cur.copy_expert("copy {} from STDIN CSV HEADER QUOTE '\"'".format(table_name), f)
        cur.execute("commit;")
        print("Loaded data into {}.\n".format(table_name))
        conn.close()
        print("DB connection closed.\n")

    except Exception as e:
        print("Error: {}".format(str(e)))
        sys.exit(1)

file_path = r'C:\Users\Public\Documents\Generated_CSV_file.csv'
table_name = 'foreo'
dbname = 'postgres'
host = 'localhost'
port = '3333'
user = 'postgres'
pwd = 'databasepassword'
pg_load_table(file_path, table_name, dbname, host, port, user, pwd)