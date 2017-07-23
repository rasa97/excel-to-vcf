import xlrd
book = xlrd.open_workbook("NC.xlsx")
sh = book.sheet_by_index(0)

s = ""
ph = "BEGIN:VCARD\nVERSION:2.1\nN:;"

for i in range(1,sh.nrows+1):
    if(sh.cell_value(i,sh.ncols-1)):
        s+=ph+str(sh.cell_value(i,1))+";;;\nFN:"+str(sh.cell_value(i,1))+"\nTEL;CELL;PREF:"+str(sh.cell_value(i,2))+"\nTEL;HOME:"+str(sh.cell_value(i,3))+"\nEND:VCARD\n"
        
    else:
        s+=ph+str(sh.cell_value(i,1))+";;;\nFN:"+str(sh.cell_value(i,1))+"\nTEL;CELL:"+str(sh.cell_value(i,2))+"\nEND:VCARD\n"
        
text_file = open("newvcf.vcf", "w")
text_file.write(s)
text_file.close()
print "Completed!"