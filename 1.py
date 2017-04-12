# -*- coding: utf-8 -*-

import xlrd
import xlwt
import sys
from xlrd import open_workbook
from xlutils.copy import copy


#release note, different from DSYBG007 is as follow:
#1, add running paramter for source file name and dest file name.


def open_excel(file = 'wenjun.xlsm'):
	data = xlrd.open_workbook(file)
	return data

def excel_table_byindex(inputfile = 'wenjun.xlsm',colnameindex=0,by_index=0):
	data = open_excel(inputfile)
	table = data.sheets()[by_index]
	nrows = table.nrows #行数
	ncols = table.ncols #列数
	#print nrows
	colnames = table.row_values(colnameindex) #某一行数据
	list = []
	for rownum in range(1, nrows):

		#app = {}
		for number in range(0, ncols):
			if table.Cells(1,1).Comment.Text():
				print (table.Cells(1,1).Comment.Text())
			#app[colnames[number]] = row[number]
		#list.append(app)
	return list



#22 is reason, 2 is PR id
#W is reason, C is PR id
def process(inputfile, outputfile):
	tables = excel_table_byindex(inputfile, 0, 0)
	for item in tables:
		print(item)
	index = 0
	f = xlwt.Workbook()
	sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
	row_index = 1
	col_index = 0

	for key in tables[0]:
		sheet1.write(0, col_index, key)
		col_index += 1

	for row in tables:
		col_index = 0
		for key in row:
			sheet1.write(row_index, col_index, row[key])
			col_index += 1


		row_index += 1
	f.save(outputfile)


if __name__ == "__main__":
    if len(sys.argv) == 3:
        #print str(sys.argv)

        inputfile = str(sys.argv[1])
        outputfile = str(sys.argv[2])
        #print inputfile
        #print outputfile
        process(inputfile, outputfile)
    else:
        print ("parameter is wrong!! right format should be \"DSYBG008.py sourcefilename destfilename\"")


