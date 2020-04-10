import os.path
import os
import re
import win32com.client as win32
import csv

# - - - - - - - - - - - - - - - file handling. - - - - - - - - - - - - - - - 
# source file path for excel files.
#master_path = 'C:\\VM_Automation 2019-12-16\\VM_Reports\\CTI_WEEKLY_REPORTS\\scheduled\\'
master_path = 'C:\\scheduled\\'

f = open ("list.txt", "r")                                          # open source excel file folder list.
f1 = f.readlines()

g = open ("data.csv","w+")

for line in f1:
    line = line.rstrip()                                            # rstrip - remove the new line.

    excel_folder_path = os.path.join(master_path, line)             # print folder path.
    #print (excel_folder_path)

    excel_name = os.listdir(excel_folder_path) 						# print files in folder.
    print (excel_name)

    excel_file = os.path.join(excel_folder_path, *excel_name) 		# print full file path. put a * on excel_name 
    #print (excel_file)

    # - - - - - reading data from excel file. - - - - -

    xlApp = win32.Dispatch('Excel.Application')                  
    wb = xlApp.Workbooks.Open(r"%s" % excel_file)                  # wb --> workbook
    ws1 = wb.Worksheets('Home')                                    # ws --> worksheet

                        # row,column
    cellData1 = ws1.cells(14,4).value   						# D14   medium (>cvss7)
    cellData2 = ws1.cells(14,5).value 							# E14   high (>cvss7)
    cellData3 = ws1.cells(14,6).value   						# F14   critical (>cvss7)
    
    cellData4 = ws1.cells(34,4).value   						# D34   medium (>cvss7) (>90days)
    cellData5 = ws1.cells(34,5).value   						# E34   high (>cvss7) (>90days)
    cellData6 = ws1.cells(34,6).value   						# F34   critical (>cvss7) (>90days)
    
    cellData7 = ws1.cells(32,3).value   						# c32   number of hosts

    wb.Close(True) 
    
    # - - - - - convert data to integer values. - - - - - - - - 

    data1 = int(cellData1)
    data2 = int(cellData2)
    data3 = int(cellData3)
    data4 = int(cellData4)
    data5 = int(cellData5)
    data6 = int(cellData6)
    data7 = int(cellData7)

    #print (cellData1, cellData2, cellData3, cellData4, cellData5, cellData6, cellData7)
    #print (data1, data2, data3, data4, data5, data6, data7)

    # - - - - - write data to csv file. - - - - - - - - 

    g.write("%s,%s,%s,%s,%s,%s,%s \n" % (data1, data2, data3, data4, data5, data6, data7))
