import xlrd
fname =r'test.xls'
bk = xlrd.open_workbook(fname)
shxrange = range(bk.nsheets)

try:
    sh = bk.sheet_by_name("Sheet1")
except:
    print "no sheet in %s named Sheet1" % fname
    
nrows = sh.nrows
ncols = sh.ncols

print nrows


"""
row_list = []
for rownum in range(nrows):      
    row_data = sh.row_values(rownum)                         
    row_list.append(row_data) 
    #print row_list[rownum][1]
    #print "insert into ms_student(name,age,city) values('%s',%s,'%s')" %(row_list[rownum][0],row_list[rownum][1],row_list[rownum][2])
    cx.execute("insert into ms_student(name,age,city) values('%s',%s,'%s')" %(row_list[rownum][0],row_list[rownum][1],row_list[rownum][2])) 
    cx.commit()
cu.close()
cx.close()
"""
