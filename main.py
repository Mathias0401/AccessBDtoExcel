import pyodbc, xlsxwriter,openpyxl, time, os, traceback, tabulate
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from tabulate import tabulate

pathBD = input("Ingrese la ruta a la Base de Datos Access: ")

connStr = (r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"fr"DBQ={pathBD};")
cnxn = pyodbc.connect(connStr)

query = "SELECT * FROM Tickets"
dataf = pd.read_sql(query, cnxn)
cnxn.close()

try:
    if(os.path.isfile("./Stock.xlsx")):
        workbook = openpyxl.load_workbook('Stock.xlsx')
    else:
        workbook = openpyxl.Workbook()
    if("datos" not in workbook.sheetnames):
        workbook.create_sheet('datos')
        worksheet = workbook["datos"]
    else:
        del workbook['datos']
        workbook.create_sheet('datos')
        worksheet = workbook['datos']
        worksheet.delete_rows(1, worksheet.max_row)
except:
    traceback.print_exc()
    input("Enter pa seguir...")

try:
    fecha = pd.Timestamp(day=19,month=3,year=2024,hour=00,minute=00,second=00)
    maxRow = 0
    for r in dataframe_to_rows(dataf, header=True, index=False):
        if(r[0] == fecha or maxRow == 0):
            maxRow+=1
            worksheet.append(r)
    tab = Table(displayName="Table1", ref=f"A1:R{maxRow}")
    # Añado un estilo por defecto con filas a rayas
    style = TableStyleInfo(name="TableStyleMedium1", showFirstColumn=False,
                        showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    worksheet.add_table(tab)
    for row in range(2, worksheet.max_row+1):
        worksheet["{}{}".format("F", row)].number_format = 'General'
except:
    traceback.print_exc()

workbook.save('Stock.xlsx')

input("Enter pa seguir...")