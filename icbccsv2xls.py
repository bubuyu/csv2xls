#!/usr/bin/env python
# -*- coding: utf-8 -*- 

def convcsv(cfile):
	xlsname = cfile[0:-3]+'xls' #生成转换Excel后的文件名
	csvarr = [] #声明csv数组
	#打开csv文件，去除tab并按照逗号分隔写入二维数组
	with open(cfile, 'r', encoding='gb18030') as csvfile:
		csvreader = csv.reader(csvfile, dialect='excel')
		for line in csvreader:
			csvstr = (','.join(line)).replace ('\t', '')
			csvarr.append(csvstr.split(','))

	workbook = xlwt.Workbook(encoding='gb18030') #创建Excel工作簿，编码为GB18030
	sheet = workbook.add_sheet('Sheet1') #创建Excel工作表
	
	#设置Excel字体样式
	font = xlwt.Font()
	font.name = '新宋体'
	style = xlwt.XFStyle()
	style.font = font

	#将csv数组内容写入工作表
	colnum = 0
	rownum = 0
	for row in csvarr:
		for cell in row:
			sheet.write(rownum, colnum, cell, style)
			colnum = colnum + 1
		colnum = 0
		rownum = rownum + 1

	#保存工作表
	workbook.save (xlsname)
	print (cfile+' 已转换！')

def main():
	fpath = os.path.abspath('.') #确定当前工作路径
	files = os.listdir(fpath) #列出当前工作路径全部文件
	convfiles = [] #声明需转换文件数组

	#找到全部扩展名为csv的文件写入convfiles数组，如无csv文件则退出
	for fname in files:
		extname = os.path.splitext(fname)
		if extname[1].lower() == '.csv':
			convfiles.append(fname)
	if convfiles == []:
		print ('无csv文件!')
		exit()

	#转换文件
	for file in convfiles:
		convcsv (file)

if __name__ == '__main__':
	import csv
	import xlwt
	import os
	main();