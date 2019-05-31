# File Combining Script
# Alexander Gille 2019

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

in_path = filedialog.askdirectory()
listing = os.listdir(in_path)

files_xlsx = [f for f in listing if f[-4:] == 'xlsx']

df = pd.DataFrame()

for infile in listing:
    file_data = pd.read_excel(in_path + '/' + infile)
    df = df.append(file_data, sort=False)

out_path = in_path + ' Combined.xlsx'

writer = pd.ExcelWriter(out_path, engine='xlsxwriter')

df.to_excel(writer,
            sheet_name='Combined',
            index=False,
            header=None)

writer.save()
