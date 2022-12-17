import pandas as pd
import os
# https://stackoverflow.com/questions/50236928/openpyxl-valueerror-max-value-is-14-when-using-load-workbook
# IMPORTANT, you must do this before importing openpyxl
from unittest import mock
# Set max font family value to 100
p = mock.patch('openpyxl.styles.fonts.Font.family.max', new=100)
p.start()
from openpyxl import load_workbook
import numpy as np


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

def convert2csv(input_file_name, sheet_name=0):
    outputfile_name = input_file_name.replace("xlsx", "csv")
    if os.path.isfile(outputfile_name):
        print(outputfile_name, "already exist. Skipping converting")
    data_xls = get_df_from_file(input_file_name, sheet_name)
    data_xls.to_csv(outputfile_name, encoding='utf-8', header=None, index=False)

def convertfiles(input_file_list):
    for fname in input_file_list:
        convert2csv(fname)
    print("Conversion to csv done")

def start():
    source = get_df_from_file("2.xlsx", "dest")
    print(get_val_from_df(source, 9, 2))
    update_df(source, 9, 2, "overwrite")

def extract(data_csv, report_csv):
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
    print("Data extraction done")

def copy(inputf, writer, sheet_name):
    source = get_df_from_file(inputf)
    """
    https://stackoverflow.com/questions/20219254/how-to-write-to-an-existing-excel-file-without-overwriting-data-using-pandas
    """  
    source.to_excel(writer, sheet_name, startrow=19, startcol=0, header=None, index=False)

def copy_files(output):
    writer = pd.ExcelWriter(output, engine='openpyxl', mode='a')
    writer.book = load_workbook(output)
    writer.sheets = {ws.title:ws for ws in writer.book.worksheets}
    print("Loading Template done, start copying")
    
    filelist=["out" + str(i) + ".csv" for i in range(0, 6)]
    sheet_name = {0: "Ar_35_15",
                  1: "H2_35_15",
                  2: "CH4_35_15",
                  3: "N2_35_15",
                  4: "O2_35_15",
                  5: "CO2_35_15"}
    for i in range(len(filelist)):
        copy(filelist[i], writer, sheet_name[i])
        print("Copy", i, "has finished")
    print("Start saving")
    writer.save()

if __name__ == "__main__":
    convertfiles(["data.xlsx", "report.xlsx"])
    extract("data.csv", "report.csv")
    # copy_files(output="Permeation Test Template (Short).xlsm")
    copy_files(output="Permeation Test Template (Short).xlsx")
    print("Done")
