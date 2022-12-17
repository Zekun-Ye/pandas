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

def get_df_from_file(filename, sheetname=0):
    if "csv" in filename:
        return pd.read_csv(filename, header=None)
    else:
        return pd.read_excel(filename, sheet_name=sheetname, engine='openpyxl', header=None)
        # return pd.read_excel(filename, sheet_name=sheetname, engine='openpyxl_wo_formatting', header=None)

def get_val_from_df(df, row, col):
    return df.iloc[row, col]

def create_df_from_val(val):
    return pd.DataFrame(np.array([val]))
    
def update_df(df, row, col, new_val):
    df.iloc[row, col] = new_val

def write_df_to_file(df, filename, sheetname, start_row, start_col):
    append_df_to_excel(filename=filename, df=df, sheet_name=sheetname, startrow=start_row, startcol=start_col, 
                       truncate_sheet=False, header=None, index=False)

def convert2csv(input_file_name, sheet_name=0):
    outputfile_name = input_file_name.replace("xlsx", "csv")
    if os.path.isfile(outputfile_name):
        print(outputfile_name, "already exist. Skipping converting")
    data_xls = get_df_from_file(input_file_name, sheet_name)
    data_xls.to_csv(outputfile_name, encoding='utf-8', header=None, index=False)

def convertfiles(input_file_list):
    for fname in input_file_list:
        convert2csv(fname)

def start():
    source = get_df_from_file("2.xlsx", "dest")
    print(get_val_from_df(source, 9, 2))
    update_df(source, 9, 2, "overwrite")
    write_df_to_file(source, "2.xlsx", "dest", 9, 3)

def step(data_csv, report_csv):
    data = [] # list of sth
    with open(data_csv, "r") as fin:
        lines = fin.readlines()[1:] # list of strings
        for i in range(len(lines)):
            line = lines[i].strip()
            splitted = line.split(",")
            first, last = splitted[0], splitted[7]
            if first == "":
                continue
            first, last = int(first), float(last)
            data.append([first, last])
            # print(first, last)

    report = []
    with open(report_csv, "r", errors='ignore') as fin:
        lines = fin.readlines()
        for i in range(len(lines)):
            line = lines[i].strip()
            if "Time - pressurize (s)" in line:
               time_1 = line.split(",")[1]
               time_1 = int(float(time_1))
               assert "Time - end (s)" in lines[i+1]
               time_2 = lines[i+1].strip().split(",")[1]
               time_2 = int(float(time_2))
               report.append([time_1, time_2])
    # print(report)
    # print(len(report))
    for i in range(len(report)):
        t1, t2 = report[i]
        filter_data = list(filter(lambda x: x[0] >= t1 and x[0] <= t2, data))
        # print(t1, t2, filter_data)
        to_write = []
        for first_last in filter_data:
            first_last = list(map(str, first_last))
            to_write.append(",".join(first_last) + "\n")
        with open("out" + str(i) + ".csv", "w") as fout:
            fout.writelines(to_write)

def copy(inputf, outputf, sheet_name):
    source = get_df_from_file(inputf)
    write_df_to_file(source, outputf, sheet_name, 19, 0)
    

def copy_files(output):
    filelist=["out" + str(i) + ".csv" for i in range(0, 6)]
    sheet_name = {0: "Ar_35_15",
                  1: "H2_35_15",
                  2: "CH4_35_15",
                  3: "N2_35_15",
                  4: "O2_35_15",
                  5: "CO2_35_15"}
    for i in range(len(filelist)):
        copy(filelist[i], output, sheet_name[i])

if __name__ == "__main__":
    convertfiles(["data.xlsx", "report.xlsx"])
    step("data.csv", "report.csv")
    copy_files(output="Permeation Test Template (Short).xlsm")
