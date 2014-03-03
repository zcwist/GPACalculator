# -*- coding: utf-8 -*- 
import tkFileDialog
import Tkinter
import xlrd
import xlwt
import tkMessageBox

class Student:
	idNum = ''
	name = ''
	classNum = ''
	scoreSum = 0
	creditSum = 0
	gpa = 0
	def __init__(self, studnetId, studentName, studentClass):
		self.idNum = studnetId
		self.name = studentName
		self.classNum = studentClass
	def addCourse(self, score, credit):
		self.scoreSum = self.scoreSum + score * credit
		self.creditSum = self.creditSum + credit
	def calculate(self):
		self.gpa = self.scoreSum / self.creditSum


def getFilename():
	master = Tkinter.Tk()
	master.withdraw()

	filename = tkFileDialog.askopenfilename(initialdir = '', defaultextension='.xls', filetypes=[('excel 97-03',"*.xls")])
	return filename

def openExcel(file):
	try:
		data = xlrd.open_workbook(file)
		return data
	except Exception,e:
		print str(e)
		tkMessageBox.showinfo("Error", "Choose a excel file")


def calculate(data):
	table = data.sheets()[0]
	nrows = table.nrows
	ncols = table.ncols

	studentList = {}
	for row in range(1,nrows):
		rowData = table.row_values(row)
		studentId = rowData[0]
		if not studentList.has_key(studentId):
			newStudent = Student(studentId, rowData[1], rowData[13])
			studentList[studentId] = newStudent
		studentList[studentId].addCourse(rowData[5], rowData[7])

	for studentId in studentList:
		studentList[studentId].calculate()

	return studentList

def writeResult(fileName,data):
	file = xlwt.Workbook()
	table = file.add_sheet('GPA')
	table.write(0,0,u'学号')
	table.write(0,1,u'姓名')
	table.write(0,2,u'班级')
	table.write(0,3,u'学分绩')

	i = 1
	order = sorted(data.keys())
	for idIndex in order:
		table.write(i,0,data[idIndex].idNum)
		table.write(i,1,data[idIndex].name)
		table.write(i,2,data[idIndex].classNum)
		table.write(i,3,float(data[idIndex].gpa))
		i = i+1
	file.save(fileName)



def main():
	fileName = getFilename()
	originData = openExcel(fileName)
	studentList = calculate(originData)
	writeResult(u'学分绩结果（尽快修改文件名）.xls', studentList)



if __name__=="__main__":
	main()