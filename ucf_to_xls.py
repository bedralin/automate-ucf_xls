#!/usr/bin/env python
#Following command will print documentation of ucf_to_xls.py:
#pydoc ucf_to_xls  

"""
OVERVIEW:
.ucf to .xls (Read Description for more info)

AUTHOR:
Bronson Edralin <bedralin@hawaii.edu>
University of Hawaii at Manoa
Instrumentation Development Lab (IDLab), WAT214

DESCRIPTION:
This script is used to extract I/O_Port_Name and Pin_Location_Name 
from the .ucf file and directly input it into the excel file (.xls). 

INSTRUCTIONS:
In order to make this work, you need to install 3 important packages:
1) xlutils (http://pypi.python.org/pypi/xlutils)
   a) Go to link and Download xlutils-VERS#.tar.gz
2) xlrd (http://pypi.python.org/pypi/xlrd)
   a) Need this to read worksheets
   b) Go to link and Download xlrd-VERS#.tar.gz
3) xlwt
   a) Need this to write worksheets
   b) Go to link and Download xlwt-VERS#.tar.gz
4) Install the packages.
   a) To unzip using terminal: 
      i) tar zxvf NameOfFile.tar.gz
   b) Install
      i) sudo python setup.py install 
"""

from collections import *
import xlwt
import string
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook # http://pypi.python.org/pypi/xlrd
from xlwt import easyxf # http://pypi.python.org/pypi/xlwt
import os


def extractUCF(ifile):
    fIn=open(ifile)           #ifile is results
    fields_found = 0  #To map or keep track of values found
    result = {}
    for line in fIn:          #for loop going through each line in input file
	line = line.replace('=',' ')  #Replace each delimiter
        line = line.replace(',',' ')
	line = line.replace('|',' ')
	line = line.replace(';',' ')
	line = line.replace(':',' ')
	line = line.replace('"',' ')
	value=line.split()	      #with a blank space (meaning splitting)
        #if value!=[]:   #if a string or value occur
        #Ensuring right position
	if value!=[]:   #if a string or value occur
	    if len(value) >= 4 and value[0] == 'NET' and value[2]=='LOC':
		fields_found += 1
		result[value[3].upper()] = value[1].upper()
	    if fields_found==7:#We found all 8!
		fields_found=0 #Reset mapper
    return result
    fIn.close()


def changeCell(worksheet, row, col, text):
    """ Changes a worksheet cell text while preserving formatting """
    # Adapted from http://stackoverflow.com/a/7686555/1545769
    previousCell = worksheet._Worksheet__rows.get(row)._Row__cells.get(col)
    worksheet.write(row, col, text)
    newCell = worksheet._Worksheet__rows.get(row)._Row__cells.get(col)
    newCell.xf_idx = previousCell.xf_idx


#EDIT THIS IF YOUR DIMENSION CHANGES ON EXCEL
#Create a dictionary for route on Excel Sheet
def routeExcel():
    loc = string.ascii_uppercase + " AA" + " AB" " AC" + " AD" + " AE" + " AF"
    #xcelMapUCF = dict( (key, 0) for key in (string.ascii_uppercase + "AA" + "AB" \
    #"AC" + "AD" + "AE" + "AF" )
    prekeys = list(string.ascii_uppercase)
    prekeys.remove("I")
    prekeys.remove("O")
    prekeys.remove("Q")
    prekeys.remove("S")
    prekeys.remove("X")
    prekeys.remove("Z")
    prekeys.extend(["AA","AB","AC","AD","AE","AF"])
    # keys represent actual Pin #s on FPGA
    keys = [] # Initialize Key List
    i,j = 0,0
    for i in range(0,26):
	for j in range(1,27):
	    keys.append(str(prekeys[i]) + str(j)) 
    # values represent cell location on excel file
    values=[]
    x,y = 0,0
    for x in range(1,52,2):
	for y in range(2,53,2):
	    values.append((x,y))

    excelMapCoords = {}
    i = 0
    for element in keys:
	excelMapCoords[element] = values[i]
	i += 1
    #print "exelMapCoords is: ",excelMapCoords
    return excelMapCoords # Dictionary used for Mapping Location


# Read from .ucf file
# Create an excel file with Pin_Location_Names (LOC) and I/O_Port_Names (UCF)
# ucfMapList(input,output)
def ucfMapList(ifile_ucf,ofile_xls):
    book = xlwt.Workbook(encoding="utf-8")  # Create a workbook
    sheet1 = book.add_sheet("Sheet 1")      # Create Sheet 1

    ucfMap = extractUCF(ifile_ucf)    # Name of file to 
    keylist = ucfMap.keys()
    keylist.sort()
    i = 0
    keylist = [element.upper() for element in keylist]
    for key in keylist:
	print "%s: %s" % (key, ucfMap[key])
	sheet1.write(i,0,key)
	sheet1.write(i,1,ucfMap[key])
	i += 1
    #print keylist
    book.save(ofile_xls)


# Edit Excel File and put extracted I/O_port_names from UCF file in there
def ucf_to_xls(ifile_ucf,ofile_xls,sheet_name):
    file_path = ofile_xls

    book = open_workbook(file_path,formatting_info=True)
    for index in range(book.nsheets):
	worksheet_name = book.sheet_by_index(index)
	if worksheet_name.name == sheet_name:
	    index_sheet_numb = index

    # use r_sheet if you want to make conditional writing to sheets
    #r_sheet = book.sheet_by_index(index_sheet_numb) # read only copy to introspect the file
    wb = copy(book) # a writable copy (can't read values out of this, only write to it)
    w_sheet = wb.get_sheet(index_sheet_numb) # sheet write within writable copy

    ucfMap = extractUCF(ifile_ucf)    # Name of file to 
    keylist = ucfMap.keys()
    keylist.sort()
    excelMapCoords = routeExcel()     

    i = 0
    for key in keylist:
	x,y = excelMapCoords[str(key)]
	changeCell(w_sheet,x,y,ucfMap[key]) # Write to Sheet without changing format
	#w_sheet.write(x,y,ucfMap[key])  # Write to sheet but changes format
	i += 1
    
    # Save .xls file into a diff file name
    wb.save(os.path.splitext(file_path)[-2]+"_rv"+os.path.splitext(file_path)[-1])

# Run the automation scripts
ucfMapList("spartan3.ucf","ucfMapList.xls")
ucf_to_xls("spartan3.ucf","test.xls","Spartan-6 FGG")


