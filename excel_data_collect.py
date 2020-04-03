import os.path
import os
import re
import win32com.client as win32
import csv

# - - - - - - - - - - - - - - - file handling. - - - - - - - - - - - - - - - 
# source file path for excel files.
master_path = 'C:\\file_location\\'

f = open ("list.txt", "r")                                          # open source excel file folder list.
f1 = f.readlines()

g = open ("data.csv","w+")

for line in f1:
    line = line.rstrip()                                            # rstrip - remove the new line.

    excel_folder_path = os.path.join(master_path, line)             # print folder path.
    #print (excel_folder_path)

    excel_name = os.listdir(excel_folder_path)                      # print files in folder.
    #print (excel_name)

    excel_file = os.path.join(excel_folder_path, *excel_name)      # print full file path.
    #print (excel_file)

 

    # - - - - - reading data from excel file. - - - - -

    xlApp = win32.Dispatch('Excel.Application')                  
    wb = xlApp.Workbooks.Open(r"%s" % excel_file)                  # wb --> workbook
    ws1 = wb.Worksheets('Home')                                     # ws --> worksheet

    # column, row
    # D14 (4,14)     E14 (5,14)     F14 (6,14)     C32 (3,32)     C33 (3,33)

                    # row,column
    cellData1 = ws1.cells(14,4).value
    cellData2 = ws1.cells(14,5).value
    cellData3 = ws1.cells(14,6).value
    cellData4 = ws1.cells(32,3).value
    cellData5 = ws1.cells(33,3).value

    wb.Close(True) 
    # - - - - - - - - - - - - - 

    data1 = int(cellData1)
    data2 = int(cellData2)
    data3 = int(cellData3)
    data4 = int(cellData4)
    data5 = int(cellData5)


    #print (cellData1, cellData2, cellData3, cellData4, cellData5)
    #print (data1, data2, data3, data4, data5)

    g.write("%s,%s,%s,%s,%s \n" % (data1, data2, data3, data4, data5))
