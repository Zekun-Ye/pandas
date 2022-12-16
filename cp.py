import pandas as pd
import os
from openpyxl import load_workbook
import numpy as np


def append_df_to_excel(filename, df, sheet_name='Sheet1', startrow=None, startcol=None,
                       truncate_sheet=False, 
                       **to_excel_kwargs):
    """
    https://stackoverflow.com/questions/20219254/how-to-write-to-an-existing-excel-file-without-overwriting-data-using-pandas
    """
    # Excel file doesn't exist - saving and exiting
    if not os.path.isfile(filename):
        df.to_excel(
            filename,
            sheet_name=sheet_name, 
            startrow=startrow if startrow is not None else 0, 
            **to_excel_kwargs)
        return
    
    # ignore [engine] parameter if it was passed
    if 'engine' in to_excel_kwargs:
        to_excel_kwargs.pop('engine')

    writer = pd.ExcelWriter(filename, engine='openpyxl', mode='a')

    # try to open an existing workbook
    writer.book = load_workbook(filename)
    
    # get the last row in the existing Excel sheet
    # if it was not specified explicitly
    if startrow is None and sheet_name in writer.book.sheetnames:
        startrow = writer.book[sheet_name].max_row

    # truncate sheet
    if truncate_sheet and sheet_name in writer.book.sheetnames:
        # index of [sheet_name] sheet
        idx = writer.book.sheetnames.index(sheet_name)
        # remove [sheet_name]
        writer.book.remove(writer.book.worksheets[idx])
        # create an empty sheet [sheet_name] using old index
        writer.book.create_sheet(sheet_name, idx)
    
    # copy existing sheets
    writer.sheets = {ws.title:ws for ws in writer.book.worksheets}

    if startrow is None:
        startrow = 0
    if startcol is None:
        startcol = 0

    # write out the new sheet
    df.to_excel(writer, sheet_name, startrow=startrow, startcol=startcol, **to_excel_kwargs)

    # save the workbook
    writer.save()

def get_df_from_file(filename, sheetname):
    return pd.read_excel(filename, sheet_name=sheetname, engine='openpyxl', header=None)

def get_val_from_df(df, row, col):
    return df.iloc[row, col]

def create_df_from_val(val):
    return pd.DataFrame(np.array([val]))
    
def update_df(df, row, col, new_val):
    df.iloc[row, col] = new_val

def write_df_to_file(df, filename, sheetname, start_row, start_col):
    append_df_to_excel(filename=filename, df=df, sheet_name=sheetname, startrow=start_row, startcol=start_col, 
                       truncate_sheet=False, header=None, index=False)
    
    
def start():
    source = get_df_from_file("2.xlsx", "dest")
    print(get_val_from_df(source, 9, 2))
    update_df(source, 9, 2, "overwrite")
    write_df_to_file(source, "2.xlsx", "dest", 9, 3)
    
if __name__ == "__main__":
    start()
