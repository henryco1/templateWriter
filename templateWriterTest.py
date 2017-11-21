import sys
import unittest
import os 

from lxml import etree as ET
from lxml import html as HT

from PyQt5.QtWidgets import QApplication
from PyQt5 import QtWidgets
from PyQt5.QtTest import QTest
from PyQt5.QtCore import QPoint, Qt
import templateWriter

app = QApplication(sys.argv)

class TemplateWriterTest(unittest.TestCase):

	dir_path = os.path.dirname(os.path.abspath(__file__))
	testGui = templateWriter.TemplateWriter()
	testGui.templateDict = {}

	def test_fileOpen(self):
		# Test opening file
		with open('testData.xml', 'r') as testFile:
			self.data=testFile.read()

		# read from a test file and add the data to the program
		tempDict = self.testGui.parseXML(self.dir_path+'/testData.xml')
		self.testGui.setTemplateDict(tempDict)
		self.testGui.setListData(tempDict)
		self.testGui.templateTextList.sortItems()

		testFile.close()

	def test_setListWidgetText(self):
		self.testGui.templateTextList.item(1).setSelected(True)

		# Check TextEdit widget
		self.testGui.setTemplateTextEdit()
		templateTextEditText = self.testGui.finalEdit.toPlainText()
		print("The value of the string is " + str(templateTextEditText))

		# Check TemplateDesc widget
		self.testGui.setTemplateDesc
		templateDescText = self.testGui.templateDisplay.toPlainText()
		print("The value of the templateDesc string is " + str(templateDescText))

		for i in range(self.testGui.templateTextList.count()):
			item = self.testGui.templateTextList.item(i)
			print(item.text())

	def test_defaults(self):
		# constructor test (not needed, this is only an example)
		self.testGui.newAction.trigger()

if __name__ == "__main__":
	unittest.main()