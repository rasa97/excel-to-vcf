# excel-to-vcf
Converting Excel format Contacts to vcf format

Reading of xlsx file done by xlrd package (https://pypi.python.org/pypi/xlrd)

The code here is not general; it is customised for my specific file. Before using it, make sure to change the code according to your Excel file. 


Structure of my xlsx file:

1st row contains heading

1st column is for serial number

2nd column contains name

3rd column contains phone number

4th column may or may not contain the second phone number
