import sys
from PyQt4 import QtGui, QtCore
from PyQt4.QtCore import Qt
# import xml.etree.ElementTree as ET
from lxml import etree as ET
from lxml import html as HT
from io import StringIO


import os 
import imp
import re

from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

class Main(QtGui.QMainWindow):

  def __init__(self, parent = None):
    QtGui.QMainWindow.__init__(self,parent)
    self.dir_path = os.path.dirname(os.path.realpath(__file__))
    self.config_path = str(self.dir_path) + "/template.txt"

    self.filename = ""

    self.initUI()

    self.document = Document()

  def initToolbar(self):

    self.newAction = QtGui.QAction(QtGui.QIcon("icons/new.png"),"New",self)
    self.newAction.setShortcut("Ctrl+N")
    self.newAction.setStatusTip("Create a new document from scratch.")
    self.newAction.triggered.connect(self.new)

    self.openAction = QtGui.QAction(QtGui.QIcon("icons/open.png"),"Open file",self)
    self.openAction.setStatusTip("Open existing document")
    self.openAction.setShortcut("Ctrl+O")
    self.openAction.triggered.connect(self.open)

    self.saveAction = QtGui.QAction(QtGui.QIcon("icons/save.png"),"Save",self)
    self.saveAction.setStatusTip("Save document")
    self.saveAction.setShortcut("Ctrl+S")
    self.saveAction.triggered.connect(self.save)

    self.printAction = QtGui.QAction(QtGui.QIcon("icons/print.png"),"Print document",self)
    self.printAction.setStatusTip("Print document")
    self.printAction.setShortcut("Ctrl+P")
    self.printAction.triggered.connect(self.printHandler)

    self.previewAction = QtGui.QAction(QtGui.QIcon("icons/preview.png"),"Page view",self)
    self.previewAction.setStatusTip("Preview page before printing")
    self.previewAction.setShortcut("Ctrl+Shift+P")
    self.previewAction.triggered.connect(self.preview)

    self.quitAction = QtGui.QAction(QtGui.QIcon("icons/quit.png"),"Quit application",self)
    self.quitAction.setStatusTip("Close this application")
    self.quitAction.setShortcut("Ctrl+Q")
    self.quitAction.triggered.connect(QtGui.qApp.quit)

    self.cutAction = QtGui.QAction(QtGui.QIcon("icons/cut.png"),"Cut to clipboard",self)
    self.cutAction.setStatusTip("Delete and copy text to clipboard")
    self.cutAction.setShortcut("Ctrl+X")
    self.cutAction.triggered.connect(self.finalEdit.cut)

    self.copyAction = QtGui.QAction(QtGui.QIcon("icons/copy.png"),"Copy to clipboard",self)
    self.copyAction.setStatusTip("Copy text to clipboard")
    self.copyAction.setShortcut("Ctrl+C")
    self.copyAction.triggered.connect(self.finalEdit.copy)

    self.pasteAction = QtGui.QAction(QtGui.QIcon("icons/paste.png"),"Paste from clipboard",self)
    self.pasteAction.setStatusTip("Paste text from clipboard")
    self.pasteAction.setShortcut("Ctrl+V")
    self.pasteAction.triggered.connect(self.finalEdit.paste)

    self.undoAction = QtGui.QAction(QtGui.QIcon("icons/undo.png"),"Undo last action",self)
    self.undoAction.setStatusTip("Undo last action")
    self.undoAction.setShortcut("Ctrl+Z")
    self.undoAction.triggered.connect(self.finalEdit.undo)

    self.redoAction = QtGui.QAction(QtGui.QIcon("icons/redo.png"),"Redo last undone thing",self)
    self.redoAction.setStatusTip("Redo last undone thing")
    self.redoAction.setShortcut("Ctrl+Y")
    self.redoAction.triggered.connect(self.finalEdit.redo)

    self.getSelectedTxtAction = QtGui.QAction(QtGui.QIcon("icons/reorder.png"),"Get templates",self)
    self.getSelectedTxtAction.setStatusTip("Toggle drag and drop for reordering the template list")
    self.getSelectedTxtAction.setShortcut("Ctrl+T")
    self.getSelectedTxtAction.triggered.connect(self.getSelectedTxt)

    self.templateOptionsAction = QtGui.QAction(QtGui.QIcon("templateOptions.png"),"Template options",self)
    self.templateOptionsAction.setStatusTip("Adjust options for user defined template variables")
    self.templateOptionsAction.setShortcut("Ctrl+E")
    self.templateOptionsAction.triggered.connect(self.templateSettings)

    bulletAction = QtGui.QAction(QtGui.QIcon("icons/bullet.png"),"Insert bullet List",self)
    bulletAction.setStatusTip("Insert bullet list")
    bulletAction.setShortcut("Ctrl+Shift+B")
    bulletAction.triggered.connect(self.bulletList)

    numberedAction = QtGui.QAction(QtGui.QIcon("icons/number.png"),"Insert numbered List",self)
    numberedAction.setStatusTip("Insert numbered list")
    numberedAction.setShortcut("Ctrl+Shift+L")
    numberedAction.triggered.connect(self.numberList)

    self.toolbar = self.addToolBar("Options")

    self.toolbar.addAction(self.newAction)
    self.toolbar.addAction(self.openAction)
    self.toolbar.addAction(self.saveAction)

    self.toolbar.addSeparator()

    self.toolbar.addAction(self.printAction)
    self.toolbar.addAction(self.previewAction)

    self.toolbar.addSeparator()

    self.toolbar.addAction(self.cutAction)
    self.toolbar.addAction(self.copyAction)
    self.toolbar.addAction(self.pasteAction)
    self.toolbar.addAction(self.undoAction)
    self.toolbar.addAction(self.redoAction)
    self.toolbar.addAction(self.getSelectedTxtAction)
    # self.toolbar.addAction(self.templateOptionsAction)

    self.toolbar.addSeparator()

    self.toolbar.addAction(bulletAction)
    self.toolbar.addAction(numberedAction)

    # Makes the next toolbar appear underneath this one
    self.addToolBarBreak()

  def initFormatbar(self):
    
    fontBox = QtGui.QFontComboBox(self)

    # sets the default font in the fontbox and the textedit widget
    # self.templateDisplay.setCurrentFont(QtGui.QFont('Times New Roman'))
    # fontBox.setCurrentFont(QtGui.QFont('Times New Roman'))
    fontBox.currentFontChanged.connect(lambda font: self.finalEdit.setCurrentFont(font))

    fontSize = QtGui.QSpinBox(self)

    # Will display " pt" after each value
    fontSize.setSuffix(" pt")

    fontSize.valueChanged.connect(lambda size: self.finalEdit.setFontPointSize(size))

    fontSize.setValue(12)

    fontColor = QtGui.QAction(QtGui.QIcon("icons/font-color.png"),"Change font color",self)
    fontColor.triggered.connect(self.fontColorChanged)

    backColor = QtGui.QAction(QtGui.QIcon("icons/highlight.png"),"Change background color",self)
    backColor.triggered.connect(self.highlight)

    boldAction = QtGui.QAction(QtGui.QIcon("icons/bold.png"),"Bold",self)
    boldAction.triggered.connect(self.bold)

    italicAction = QtGui.QAction(QtGui.QIcon("icons/italic.png"),"Italic",self)
    italicAction.triggered.connect(self.italic)

    underlAction = QtGui.QAction(QtGui.QIcon("icons/underline.png"),"Underline",self)
    underlAction.triggered.connect(self.underline)

    alignLeft = QtGui.QAction(QtGui.QIcon("icons/align-left.png"),"Align left",self)
    alignLeft.triggered.connect(self.alignLeft)

    alignCenter = QtGui.QAction(QtGui.QIcon("icons/align-center.png"),"Align center",self)
    alignCenter.triggered.connect(self.alignCenter)

    alignRight = QtGui.QAction(QtGui.QIcon("icons/align-right.png"),"Align right",self)
    alignRight.triggered.connect(self.alignRight)

    alignJustify = QtGui.QAction(QtGui.QIcon("icons/align-justify.png"),"Align justify",self)
    alignJustify.triggered.connect(self.alignJustify)

    indentAction = QtGui.QAction(QtGui.QIcon("icons/indent.png"),"Indent Area",self)
    indentAction.setShortcut("Ctrl+Tab")
    indentAction.triggered.connect(self.indent)

    dedentAction = QtGui.QAction(QtGui.QIcon("icons/dedent.png"),"Dedent Area",self)
    dedentAction.setShortcut("Shift+Tab")
    dedentAction.triggered.connect(self.dedent)

    self.formatbar = self.addToolBar("Format")

    self.formatbar.addWidget(fontBox)
    self.formatbar.addWidget(fontSize)

    self.formatbar.addSeparator()

    self.formatbar.addAction(fontColor)
    self.formatbar.addAction(backColor)

    self.formatbar.addSeparator()

    self.formatbar.addAction(boldAction)
    self.formatbar.addAction(italicAction)
    self.formatbar.addAction(underlAction)

    self.formatbar.addSeparator()

    self.formatbar.addAction(alignLeft)
    self.formatbar.addAction(alignCenter)
    self.formatbar.addAction(alignRight)
    self.formatbar.addAction(alignJustify)

    self.formatbar.addSeparator()

    self.formatbar.addAction(indentAction)
    self.formatbar.addAction(dedentAction)


  def initMenubar(self):

    menubar = self.menuBar()

    file = menubar.addMenu("&File")
    edit = menubar.addMenu("&Edit")
    view = menubar.addMenu("&View")
    helpmenu = menubar.addMenu("&Help")

    file.addAction(self.newAction)
    file.addAction(self.openAction)
    file.addAction(self.saveAction)
    file.addAction(self.printAction)
    file.addAction(self.previewAction)
    file.addAction(self.quitAction)

    edit.addAction(self.undoAction)
    edit.addAction(self.redoAction)
    edit.addAction(self.cutAction)
    edit.addAction(self.copyAction)
    edit.addAction(self.pasteAction)
    edit.addAction(self.getSelectedTxtAction)
    edit.addAction(self.templateOptionsAction)

    # Toggling actions for the various bars
    toolbarAction = QtGui.QAction("Toggle Toolbar",self)
    toolbarAction.triggered.connect(self.toggleToolbar)

    formatbarAction = QtGui.QAction("Toggle Formatbar",self)
    formatbarAction.triggered.connect(self.toggleFormatbar)

    statusbarAction = QtGui.QAction("Toggle Statusbar",self)
    statusbarAction.triggered.connect(self.toggleStatusbar)

    view.addAction(toolbarAction)
    view.addAction(formatbarAction)
    view.addAction(statusbarAction)

  def initUI(self):

    # Contains the list of templates
    self.skilllist = QtGui.QListWidget(self)

    # Contains the final product
    self.templateDisplay = QtGui.QTextBrowser(self)

    # Contains text within template
    self.finalEdit = QtGui.QTextEdit(self)

    self.initToolbar()
    self.initFormatbar()
    self.initMenubar()

    self.messageSplitter = QtGui.QSplitter(Qt.Vertical)
    self.messageSplitter.addWidget(self.finalEdit)
    self.messageSplitter.addWidget(self.templateDisplay)
    self.messageSplitter.setSizes([700,100])


    self.mainSplitter = QtGui.QSplitter(Qt.Horizontal)
    self.mainSplitter.addWidget(self.skilllist)
    self.mainSplitter.addWidget(self.messageSplitter)
    self.mainSplitter.setSizes([100,500])

    self.setCentralWidget(self.mainSplitter)

    # Gets selected items on skill list to populate description text
    self.skilllist.itemClicked.connect(self.getDesc)

    # Set the tab stop width to around 33 pixels which is
    # about 8 spaces
    self.finalEdit.setTabStopWidth(33)

    # Initialize a statusbar for the window
    self.statusbar = self.statusBar()

    # If the cursor position changes, call the function that displays
    # the line and column number
    self.finalEdit.cursorPositionChanged.connect(self.cursorPosition)

    # x and y coordinates on the screen, width, height
    self.setGeometry(0,0,1030,800)

    self.setWindowTitle("Template Writer")

    self.setWindowIcon(QtGui.QIcon("icons/icon.png"))

  def toggleToolbar(self):

    state = self.toolbar.isVisible()

    # Set the visibility to its inverse
    self.toolbar.setVisible(not state)

  def toggleFormatbar(self):

    state = self.formatbar.isVisible()

    # Set the visibility to its inverse
    self.formatbar.setVisible(not state)

  def toggleStatusbar(self):

    state = self.statusbar.isVisible()

    # Set the visibility to its inverse
    self.statusbar.setVisible(not state)

  def new(self):

    spawn = Main(self)
    spawn.show()

  def open(self):

    self.templateDict = {}
    # Key = name of skill
    # Value = template text for skill

    self.listOfDesc = []
    # List of skill descriptions used to easily access skill templates

    # Get filename 
    self.filename = QtGui.QFileDialog.getOpenFileName(self, 'Open File',".","*.xml")

    if self.filename:
      fname = str(self.filename)
      tree = HT.parse(fname)
      root = tree.getroot()

      self.templateDict = self.getTemplateData(root)

      # Adds list of skills to the checkbox list
      for key in self.templateDict:
        self.skilllist.addItem(key)

      # Enables drag and drop mode and multiple selections
      self.skilllist.setDragDropMode(self.skilllist.InternalMove)
      self.skilllist.setSelectionMode(QtGui.QAbstractItemView.ExtendedSelection)


      for value in self.templateDict.items():
        # Adds the description to the dictionary and then adds the text to the textbox
        # note that value will be a tuple
        self.listOfDesc.append(value[1])
        self.templateDisplay.setText(value[1])

    # sets the font for the newly imported text
    # self.templateDisplay.setCurrentFont(QtGui.QFont('Times New Roman'))

  def getTemplateData(self, fileTxt): 

    outputDict = {}
    desc = ""
    name = ""

    for element in fileTxt.iter():
      if element.tag == 'description':
        desc = ET.tostring(element, pretty_print=True)
        # print(desc)

      # print(element.tag)

      if element.tag == 'templatedata':
        name = element.get("name")
        print(name)

      outputDict[name] = desc

    return outputDict

  def getDesc(self):

    # Gets descriptions of items on skill list

    # clear
    self.templateDisplay.setText("")

    value = self.skilllist.currentItem().text()
    skillDesc = self.templateDict[str(value)]
    self.templateDisplay.insertHtml(skillDesc)

  def getSelectedTxt(self):

    # Outputs the user's selected text on in the textBox
    outputStr = ""

    # clear
    self.finalEdit.setText("")

    for i in range(self.skilllist.count()):
      item = self.skilllist.item(i)
      if item.isSelected() == True:
        listKey = str(item.text())
        outputStr  += self.templateDict[listKey]

    outputStr = self.parser(outputStr)
    self.finalEdit.insertHtml(outputStr)

  def save(self):

    # Only open dialog if there is no filename yet


    # if not self.filename:
    self.filename = QtGui.QFileDialog.getSaveFileName(self, 'Save File')

    # We just store the contents of the text file along with the
    # format in plain text, which Qt does in a very nice way for us
    with open(self.filename+".docx","wt") as file:
      file.write(self.finalEdit.toHtml())
      # print(self.filename)


  def preview(self):

    # Open preview dialog
    preview = QtGui.QPrintPreviewDialog()

    # If a print is requested, open print dialog
    preview.paintRequested.connect(lambda p: self.finalEdit.print_(p))

    preview.exec_()

  def printHandler(self):

    # Open printing dialog
    dialog = QtGui.QPrintDialog()

    if dialog.exec_() == QtGui.QDialog.Accepted:
        self.finalEdit.document().print_(dialog.printer())

  def cursorPosition(self):

    cursor = self.finalEdit.textCursor()

    # Mortals like 1-indexed things
    line = cursor.blockNumber() + 1
    col = cursor.columnNumber()

    self.statusbar.showMessage("Line: {} | Column: {}".format(line,col))

  def bulletList(self):

    cursor = self.finalEdit.textCursor()

    # Insert bulleted list
    cursor.insertList(QtGui.QTextListFormat.ListDisc)

  def numberList(self):

    cursor = self.finalEdit.textCursor()

    # Insert list with numbers
    cursor.insertList(QtGui.QTextListFormat.ListDecimal)

  def fontColorChanged(self):

    # Get a color from the text dialog
    color = QtGui.QColorDialog.getColor()

    # Set it as the new text color
    self.finalEdit.setTextColor(color)

  def highlight(self):

    color = QtGui.QColorDialog.getColor()

    self.finalEdit.setTextBackgroundColor(color)

  def bold(self):

    if self.finalEdit.fontWeight() == QtGui.QFont.Bold:
        self.finalEdit.setFontWeight(QtGui.QFont.Normal)

    else:
        self.finalEdit.setFontWeight(QtGui.QFont.Bold)

  def italic(self):

    state = self.finalEdit.fontItalic()

    self.finalEdit.setFontItalic(not state)

  def underline(self):

    state = self.finalEdit.fontUnderline()

    self.finalEdit.setFontUnderline(not state)

  def alignLeft(self):
    self.finalEdit.setAlignment(Qt.AlignLeft)

  def alignRight(self):
    self.finalEdit.setAlignment(Qt.AlignRight)

  def alignCenter(self):
    self.finalEdit.setAlignment(Qt.AlignCenter)

  def alignJustify(self):
    self.finalEdit.setAlignment(Qt.AlignJustify)

  def indent(self):

    # Grab the cursor
    cursor = self.finalEdit.textCursor()

    if cursor.hasSelection():
      # Store the current line/block number
      temp = cursor.blockNumber()

      # Move to the selection's end
      cursor.setPosition(cursor.anchor())

      # Calculate range of selection
      diff = cursor.blockNumber() - temp

      direction = QtGui.QTextCursor.Up if diff > 0 else QtGui.QTextCursor.Down

      # Iterate over lines (diff absolute value)
      for n in range(abs(diff) + 1):
        # Move to start of each line
        cursor.movePosition(QtGui.QTextCursor.StartOfLine)

        # Insert tabbing
        cursor.insertText("\t")

        # And move back up
        cursor.movePosition(direction)

    # If there is no selection, just insert a tab
    else:
      cursor.insertText("\t")

  def handleDedent(self,cursor):

      cursor.movePosition(QtGui.QTextCursor.StartOfLine)

      # Grab the current line
      line = cursor.block().text()

      # If the line starts with a tab character, delete it
      if line.startswith("\t"):

          # Delete next character
          cursor.deleteChar()

      # Otherwise, delete all spaces until a non-space character is met
      else:
          for char in line[:8]:

              if char != " ":
                  break

              cursor.deleteChar()

  def dedent(self):

    cursor = self.finalEdit.textCursor()

    if cursor.hasSelection():
      # Store the current line/block number
      temp = cursor.blockNumber()

      # Move to the selection's last line
      cursor.setPosition(cursor.anchor())

      # Calculate range of selection
      diff = cursor.blockNumber() - temp

      direction = QtGui.QTextCursor.Up if diff > 0 else QtGui.QTextCursor.Down

      # Iterate over lines
      for n in range(abs(diff) + 1):
        self.handleDedent(cursor)

        # Move up
        cursor.movePosition(direction)

    else:
      self.handleDedent(cursor)

  def parser(self, inputString):

    # parses a string based on a predefined user input.

    configData = self.readConfig()

    inputString = inputString.replace("{newline}", "\n")
    inputString = inputString.replace("{tab}", "\t")
    for key in configData:
      inputString = inputString.replace("{" + str(key) + "}", str(configData[key]))

    return inputString

  def templateSettings(self):

    count = 0
    keylist = []
    keyLabels = ""

    # table for editing template variable inputs
    self.window = QtGui.QWidget()
    self.table = QtGui.QTableWidget()
    self.tableItem = QtGui.QTableWidgetItem()

    # initialize data
    templateData = self.readConfig()

    # initiate table
    self.table.setRowCount(len(templateData.keys()))
    self.table.setColumnCount(1)

    # set data
    for key in templateData:
      self.table.setItem(count, 0, QtGui.QTableWidgetItem(templateData[key]))
      keyLabels = keyLabels + str(key) + ';'
      keylist.append(key)
      count+=1

    # set label, each element must be separated by a ;
    self.table.setHorizontalHeaderLabels(QtCore.QString("Var Value;").split(";"))
    self.table.setVerticalHeaderLabels(QtCore.QString(keyLabels).split(";"))

    # resize col
    header = self.table.horizontalHeader()
    header.setResizeMode(0, QtGui.QHeaderView.Stretch)

    # add, accept, and cancel buttons
    okayButton = QtGui.QPushButton('Okay', self)
    cancelButton = QtGui.QPushButton('Cancel', self)
    
    # must use a lambda or else the program does not think that connect is given a function
    okayButton.clicked.connect(lambda: self.templateSettingsOkayButton(keylist))
    okayButton.clicked.connect(self.window.close)
    cancelButton.clicked.connect(self.window.close)

    # attach items to layout
    # addWidget's 3rd parameter is rowSpan and the 4th is colSpan 
    layout = QtGui.QGridLayout(self.window)
    layout.addWidget(self.table, 0, 0, 4, 3)
    layout.addWidget(okayButton, 5, 1)
    layout.addWidget(cancelButton, 5, 2)

    # show settings window
    self.window.setWindowTitle("Template Variable Settings")
    self.window.setGeometry(365, 400, 400, 250)
    self.window.show()

  def templateSettingsOkayButton(self, keys):

    # When the user presses okay, their input on the table is saved to the file so that it can
    # be used in the template

    configFile = open(self.config_path, 'w')

    for i in range(0, self.table.rowCount()):
      item = self.table.item(i, 0)
      outputStr = str(keys[i]) + ' = ' + item.text() + '\n'
      print(outputStr)
      configFile.write(outputStr)

    configFile.close()

  def readConfig(self):

    # open the template.txt file and uses regex to get the variable name and the variable value

    tableVals = {}

    configFile = open(self.config_path, 'r')
    lines = configFile.readlines()
    lines = [line.rstrip() for line in lines] # strip \n
    for items in lines:  
      # regex group matching
      strCheck = r'(\S+)\s*=\s*(.*)'
      templateVars = re.search(strCheck, items)
      tableVals[templateVars.group(1)] = str(templateVars.group(2))

    configFile.close()
    return tableVals



def main():

  app = QtGui.QApplication(sys.argv)
  # QtGui.QApplication.setFont(QtGui.QFont('Times New Roman'))

  main = Main()
  main.show()

  sys.exit(app.exec_())

if __name__ == "__main__":
  main()