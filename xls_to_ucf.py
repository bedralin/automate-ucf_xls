#!/usr/bin/env python
#Following command will print documentation of xls_to_ucf.py:
#pydoc xls_to_ucf  

"""
OVERVIEW:
.xls to .ucf (text file format) (Read Description for more info)

AUTHOR:
Bronson Edralin <bedralin@hawaii.edu>
University of Hawaii at Manoa
Instrumentation Development Lab (IDLab), WAT214

DESCRIPTION:
This script is used to extract I/O_Port_Name and Pin_Location_Name 
from the .xls file and directly input it into a text file (.txt). 

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


#EDIT THIS IF YOUR DIMENSION CHANGES ON EXCEL
#Create a dictionary for route on Excel Sheet
def routeExcel2():
    loc = string.ascii_uppercase + " AA" + " AB" " AC" + " AD" + " AE" + " AF"
    #xcelMapUCF = dict( (key, 0) for key in (string.ascii_uppercase + "AA" + "AB" \
    #"AC" + "AD" + "AE" + "AF" )
    prevalues = list(string.ascii_uppercase)
    prevalues.remove("I")
    prevalues.remove("O")
    prevalues.remove("Q")
    prevalues.remove("S")
    prevalues.remove("X")
    prevalues.remove("Z")
    prevalues.extend(["AA","AB","AC","AD","AE","AF"])
    # keys represent actual Pin #s on FPGA
    values = [] # Initialize Key List
    i,j = 0,0
    for i in range(0,26):
	for j in range(1,27):
	    values.append(str(prevalues[i]) + str(j)) 
    # values represent cell location on excel file
    keys=[]
    x,y = 0,0
    for x in range(1,52,2):
	for y in range(2,53,2):
	    keys.append((x,y))

    locMapCoords = {}
    i = 0
    for element in keys:
	locMapCoords[element] = values[i]
	i += 1
    # keys: excel cel block, values: pin location (LOC)
    return locMapCoords # Dictionary used for Mapping Location


# Edit Excel File and put extracted I/O_port_names from UCF file in there
# xls_to_ucf(input,output)
def xls_to_ucf(ifile_xls,ofile_txt,sheet_name):
    file_path = ifile_xls
    outFile = open(ofile_txt,'w')

    book = open_workbook(file_path,formatting_info=True)
    for index in range(book.nsheets):
        worksheet_name = book.sheet_by_index(index)
        if worksheet_name.name == sheet_name:
            index_sheet_numb = index

    r_sheet = book.sheet_by_index(index_sheet_numb) # read only copy to introspect the file
    
    locMapCoords = routeExcel2()
    loc = []
    io_port_name = []
    fields_found = 0
    # Iterate over excel for io_port_name values
    for rows in range (1,52,2):
	for cols in range (2,53,2):
	    cell = r_sheet.cell_value(rows,cols)
	    if cell != "":
		loc.append(locMapCoords[(rows,cols)])
		io_port_name.append(cell)
		fields_found += 1

    for index in range(fields_found):
	outFile.write('NET {} LOC= {};\n'.format(io_port_name[index],loc[index]))
    outFile.close()	
    	
# Run the automation scripts
xls_to_ucf("test_rv.xls","ucfOUT.txt","Spartan-6 FGG")


