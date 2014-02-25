# -*- coding: utf-8 -*- 
import tkFileDialog
import xlrd
import sys
 
def getFilename():
	filename = tkFileDialog.askopenfilename(initialdir = '', defaultextension='.xlsx', filetypes=[('excel',"*.xlsx")])
	return filename

def openExcel(file):
	try:
		data = xlrd.open_workbook(file)
		return data
	except Exception,e:
		print str(e)

def main():
	openExcel(getFilename())

if __name__=="__main__":
	main()