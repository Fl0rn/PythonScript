import plotly
import pandas as pd
from openpyxl.styles import PatternFill
import openpyxl
from CreareTabelMediu import apply_cell_colors
def ConcatTable(excel_file1, excel_file2):
    df = pd.read_excel(excel_file1, engine='openpyxl')
    df2 = pd.read_excel(excel_file2, engine='openpyxl')
    df = pd.concat([df, df2], axis=1)
    print(df)
    df.to_excel(excel_file1, index=False, sheet_name='Test Overview')
    book = openpyxl.load_workbook('Book1.xlsx')
    ws = book['tete']
    apply_cell_colors(ws)
    book.save(excel_file1)
    book.close()

