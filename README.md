automate-ucf_xls
================

In this repository, I am open-sourcing my python code used to automate the process of extracting I/O_Port_Name and Pin_Location_Name from the .ucf file and directly input it into the excel file (.xls). I also automated the other direction as well.


Attached is two scripts:

ucf_to_xls.py (.ucf to .xls)
- This script (ucf_to_xls.py) will extract the I/O_Port_Name and the Pin_Location_Names in the .ucf file and output them into the appropriate cell block in the "Spartan-6 FGG" sheet from workbook, test.xls.

.xls to .ucf   (xls_to_ucf.py)
- This script (xls_to_ucf.py) will extract the I/O_Port_Name and the Pin_Location_Names in the "Spartan-6 FGG" sheet from workbook, test_rv.xls and output them into a .txt file.


You will also need to Download and install 3 Packages:
- xlutils (http://pypi.python.org/pypi/xlutils)
- xlrd (http://pypi.python.org/pypi/xlrd)
- xlwt (http://pypi.python.org/pypi/xlwt)
 
Unzip packages in terminal:
- tar zxvf NameOfFile.tar.gz

Install packages in terminal:
- sudo python setup.py install

Feel free to email me for any help:
- Bronson Edralin <bedralin@hawaii.edu>
- BS Electrical Engineering, May 2014
- MS Computer Engineering Graduate Student and Research Assistant
- Instrumentation Development Lab (IDL), Physics and Astronomy Department, University of Hawaii at Manoa

