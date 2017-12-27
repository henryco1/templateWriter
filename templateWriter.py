import sys

from lxml import etree as ET
from lxml import html as HT

from PyQt5 import QtWidgets

from PyQt5 import QtPrintSupport

from PyQt5 import QtGui, QtCore
from PyQt5.QtCore import Qt

import os 
import imp
import re
import string

class TemplateWriter(QtWidgets.QMainWindow):

    def __init__(self, parent = None):
        QtWidgets.QMainWindow.__init__(self,parent)
        self.dir_path = os.path.dirname(os.path.realpath(__file__))
        self.config_path = str(self.dir_path) + "/template.txt"

        self.filenameOpen = ""
        self.filenameSave = ""

        self.templateDict = {}

        self.initUI()

    def initToolbar(self):
        self.newAction = QtWidgets.QAction(QtGui.QIcon("icons/new.png"),"New",self)
        self.newAction.setShortcut("Ctrl+N")
        self.newAction.setStatusTip("Create a new document from scratch.")
        self.newAction.triggered.connect(self.new)

        self.openAction = QtWidgets.QAction(QtGui.QIcon("icons/open.png"),"Open file",self)
        self.openAction.setStatusTip("Open existing document")
        self.openAction.setShortcut("Ctrl+O")
        self.openAction.triggered.connect(self.open)

        self.saveAction = QtWidgets.QAction(QtGui.QIcon("icons/save.png"),"Save",self)
        self.saveAction.setStatusTip("Save document")
        self.saveAction.setShortcut("Ctrl+S")
        self.saveAction.triggered.connect(self.save)

        self.pdfAction = QtWidgets.QAction("Save to PDF", self)
        self.pdfAction.setStatusTip("Save to PDF")
        self.pdfAction.setShortcut("Ctrl+Shift+S")
        self.pdfAction.triggered.connect(self.saveToPDF)

        self.printAction = QtWidgets.QAction(QtGui.QIcon("icons/print.png"),"Print document",self)
        self.printAction.setStatusTip("Print document")
        self.printAction.setShortcut("Ctrl+P")
        self.printAction.triggered.connect(self.printHandler)

        self.previewAction = QtWidgets.QAction(QtGui.QIcon("icons/preview.png"),"Page view",self)
        self.previewAction.setStatusTip("Preview page before printing")
        self.previewAction.setShortcut("Ctrl+Shift+P")
        self.previewAction.triggered.connect(self.preview)

        self.quitAction = QtWidgets.QAction(QtGui.QIcon("icons/quit.png"),"Quit application",self)
        self.quitAction.setStatusTip("Close this application")
        self.quitAction.setShortcut("Ctrl+Q")
        self.quitAction.triggered.connect(QtWidgets.qApp.quit)

        self.cutAction = QtWidgets.QAction(QtGui.QIcon("icons/cut.png"),"Cut to clipboard",self)
        self.cutAction.setStatusTip("Delete and copy text to clipboard")
        self.cutAction.setShortcut("Ctrl+X")
        self.cutAction.triggered.connect(self.finalEdit.cut)

        self.copyAction = QtWidgets.QAction(QtGui.QIcon("icons/copy.png"),"Copy to clipboard",self)
        self.copyAction.setStatusTip("Copy text to clipboard")
        self.copyAction.setShortcut("Ctrl+C")
        self.copyAction.triggered.connect(self.finalEdit.copy)

        self.pasteAction = QtWidgets.QAction(QtGui.QIcon("icons/paste.png"),"Paste from clipboard",self)
        self.pasteAction.setStatusTip("Paste text from clipboard")
        self.pasteAction.setShortcut("Ctrl+V")
        self.pasteAction.triggered.connect(self.finalEdit.paste)

        self.undoAction = QtWidgets.QAction(QtGui.QIcon("icons/undo.png"),"Undo last action",self)
        self.undoAction.setStatusTip("Undo last action")
        self.undoAction.setShortcut("Ctrl+Z")
        self.undoAction.triggered.connect(self.finalEdit.undo)

        self.redoAction = QtWidgets.QAction(QtGui.QIcon("icons/redo.png"),"Redo last undone thing",self)
        self.redoAction.setStatusTip("Redo last undone thing")
        self.redoAction.setShortcut("Ctrl+Y")
        self.redoAction.triggered.connect(self.finalEdit.redo)

        self.getSelectedTxtAction = QtWidgets.QAction(QtGui.QIcon("icons/reorder.png"),"Compile templates",self)
        self.getSelectedTxtAction.setStatusTip("Compiles the selected items from the list onto the text editor and displays them in the order they are arranged")
        self.getSelectedTxtAction.setShortcut("Ctrl+T")
        self.getSelectedTxtAction.triggered.connect(self.setTemplateTextEdit)

        self.templateOptionsAction = QtWidgets.QAction(QtGui.QIcon("templateOptions.png"),"Template options",self)
        self.templateOptionsAction.setStatusTip("Adjust options for user defined template variables")
        self.templateOptionsAction.setShortcut("Ctrl+E")
        self.templateOptionsAction.triggered.connect(self.templateSettings)

        bulletAction = QtWidgets.QAction(QtGui.QIcon("icons/bullet.png"),"Insert bullet List",self)
        bulletAction.setStatusTip("Insert bullet list")
        bulletAction.setShortcut("Ctrl+Shift+B")
        bulletAction.triggered.connect(self.bulletList)

        numberedAction = QtWidgets.QAction(QtGui.QIcon("icons/number.png"),"Insert numbered List",self)
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
        fontBox = QtWidgets.QFontComboBox(self)

        # sets the default font in the fontbox and the textedit widget
        # self.templateDisplay.setCurrentFont(QtGui.QFont('Times New Roman'))
        fontBox.setCurrentFont(QtGui.QFont('Times New Roman'))
        # fontBox.currentFontChanged.connect(lambda font: self.finalEdit.setCurrentFont(font))

        fontSize = QtWidgets.QSpinBox(self)

        # Will display " pt" after each value
        fontSize.setSuffix(" pt")

        fontSize.valueChanged.connect(lambda size: self.finalEdit.setFontPointSize(size))

        fontSize.setValue(12)

        fontColor = QtWidgets.QAction(QtGui.QIcon("icons/font-color.png"),"Change font color",self)
        fontColor.triggered.connect(self.fontColorChanged)

        backColor = QtWidgets.QAction(QtGui.QIcon("icons/highlight.png"),"Change background color",self)
        backColor.triggered.connect(self.highlight)

        boldAction = QtWidgets.QAction(QtGui.QIcon("icons/bold.png"),"Bold",self)
        boldAction.triggered.connect(self.bold)

        italicAction = QtWidgets.QAction(QtGui.QIcon("icons/italic.png"),"Italic",self)
        italicAction.triggered.connect(self.italic)

        underlAction = QtWidgets.QAction(QtGui.QIcon("icons/underline.png"),"Underline",self)
        underlAction.triggered.connect(self.underline)

        alignLeft = QtWidgets.QAction(QtGui.QIcon("icons/align-left.png"),"Align left",self)
        alignLeft.triggered.connect(self.alignLeft)

        alignCenter = QtWidgets.QAction(QtGui.QIcon("icons/align-center.png"),"Align center",self)
        alignCenter.triggered.connect(self.alignCenter)

        alignRight = QtWidgets.QAction(QtGui.QIcon("icons/align-right.png"),"Align right",self)
        alignRight.triggered.connect(self.alignRight)

        alignJustify = QtWidgets.QAction(QtGui.QIcon("icons/align-justify.png"),"Align justify",self)
        alignJustify.triggered.connect(self.alignJustify)

        indentAction = QtWidgets.QAction(QtGui.QIcon("icons/indent.png"),"Indent Area",self)
        indentAction.setShortcut("Ctrl+Tab")
        indentAction.triggered.connect(self.indent)

        dedentAction = QtWidgets.QAction(QtGui.QIcon("icons/dedent.png"),"Dedent Area",self)
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
        file.addAction(self.pdfAction)
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
        toolbarAction = QtWidgets.QAction("Toggle Toolbar",self)
        toolbarAction.triggered.connect(self.toggleToolbar)

        formatbarAction = QtWidgets.QAction("Toggle Formatbar",self)
        formatbarAction.triggered.connect(self.toggleFormatbar)

        statusbarAction = QtWidgets.QAction("Toggle Statusbar",self)
        statusbarAction.triggered.connect(self.toggleStatusbar)

        view.addAction(toolbarAction)
        view.addAction(formatbarAction)
        view.addAction(statusbarAction)

    def initUI(self):
        # Contains the list of templates that the user can select
        self.templateTextList = QtWidgets.QListWidget(self)
        self.templateTextList.itemActivated.connect(self.setTemplateDesc)

        # When the user selects a template from the list, a description of the template appears on this display
        self.templateDisplay = QtWidgets.QTextBrowser(self)

        # A box that contains the templates compiled together
        self.finalEdit = QtWidgets.QTextEdit(self)

        # Default formatting
        self.templateDisplay.setCurrentFont(QtGui.QFont('Times New Roman'))
        self.finalEdit.setCurrentFont(QtGui.QFont('Times New Roman'))
        self.finalEdit.setWordWrapMode(1)

        self.initToolbar()
        self.initFormatbar()
        self.initMenubar()

        self.messageSplitter = QtWidgets.QSplitter(Qt.Vertical)
        self.messageSplitter.addWidget(self.finalEdit)
        self.messageSplitter.addWidget(self.templateDisplay)
        self.messageSplitter.setSizes([700,100])


        self.mainSplitter = QtWidgets.QSplitter(Qt.Horizontal)
        self.mainSplitter.addWidget(self.templateTextList)
        self.mainSplitter.addWidget(self.messageSplitter)
        self.mainSplitter.setSizes([100,500])

        self.setCentralWidget(self.mainSplitter)

        # Gets selected items on skill list to populate description text
        self.templateTextList.itemClicked.connect(self.setTemplateDesc)

        # Enables drag and drop mode and multiple selections
        # 4 = self.templateTextList.InternalMove
        self.templateTextList.setDragDropMode(4)
        self.templateTextList.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)

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
        spawn = TemplateWriter(self)
        spawn.show()

    def setTemplateDict(self, dictionaryData):
        self.templateDict = dictionaryData

    def getTemplateDict(self):
        return self.templateDict

    def open(self):
        # Handles opening XML files

        tempDict = {}

        # Gets filename and its path
        # PYQT5 Returns a tuple in PyQt5, we only need the filename
        self.filenameOpen = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File',".","(*.xml)")[0]

        if self.filenameOpen:
            tempDict = self.parseXML(self.filenameOpen)
            self.setTemplateDict(tempDict)

            self.setListData(tempDict)
            self.templateDisplay.clear()
            self.templateTextList.sortItems()

    def parseXML(self, filePath):
        # Using the text from the XML file, this function parses XML into a dictionary

        tempDict = {}
        # Key = name of skill
        # Value = template text for skill

        fname = str(filePath)
        tree = HT.parse(fname)
        root = tree.getroot()

        tempDict = self.getTemplateData(root)
        print(tempDict)
        return tempDict

    def getTemplateData(self, fileTxt): 
        # handler that extracts data from XML file into a dictionary

        outputDict = {}
        desc = ""
        name = ""

        for element in fileTxt.iter():
            if element.tag == 'description':
                # ported Python2 to Python3 by encoding in unicode
                desc = str(ET.tostring(element, pretty_print=True), 'utf-8')

                # remove xml tags so the string can be read as html correctly
                desc = desc.replace('<description>', '')
                desc = desc.replace('</description>', '')

            if element.tag == 'templatedata':
                name = element.get("name")

            outputDict[name] = desc

        return outputDict

    def setListData(self, listData):
        # Adds items to list of templates that the user can select
        for key in listData:
            self.templateTextList.addItem(key)

    def setTemplateDesc(self):
        # This function is called when the user selects an item on the template list
        # It populates the description window with the text that corresponds to the selected
        # item on the template list

        self.templateDisplay.clear()

        tempDict = self.getTemplateDict()

        value = self.templateTextList.currentItem().text()
        skillDesc = tempDict[str(value)]
        self.templateDisplay.insertHtml(skillDesc)

    def setTemplateTextEdit(self):
        # Outputs the user's selected list items in the text edit box
        outputStr = ""
        tempDict = self.getTemplateDict()

        self.finalEdit.clear()

        for i in range(self.templateTextList.count()):
            item = self.templateTextList.item(i)
            print("The selected item is " + item.text())
            if item.isSelected() == True:
                listKey = str(item.text())
                outputStr  += tempDict[listKey]

        outputStr = self.parser(outputStr)
        print(outputStr)
        self.finalEdit.insertHtml(outputStr)

    def save(self):
        # Only open dialog if there is no filenameSave yet
        if not self.filenameSave:
            self.filenameSave = QtWidgets.QFileDialog.getSaveFileName(self, 'Save File')[0]

        # Stores the contents of the saved file in plain text html.
        # By saving as HTML, the program is able to save some of the text formating
        with open(self.filenameSave+".lyn","wt") as file:
              file.write(self.finalEdit.toHtml())

    def saveToPDF(self):
        filename = QtWidgets.QFileDialog.getSaveFileName(self, 'Save to PDF')
        if filename:
            printer = QtPrintSupport.QPrinter(QtPrintSupport.QPrinter.HighResolution)
            printer.setPageSize(QtPrintSupport.QPrinter.A4)
            printer.setColorMode(QtPrintSupport.QPrinter.Color)
            printer.setOutputFormat(QtPrintSupport.QPrinter.PdfFormat)
            printer.setOutputFileName(filename[0])
            self.finalEdit.document().print_(printer)

    def preview(self):
        # Open preview dialog
        preview = QtPrintSupport.QPrintPreviewDialog()

        # If a print is requested, open print dialog
        preview.paintRequested.connect(lambda p: self.finalEdit.print_(p))

        preview.exec_()

    def printHandler(self):
        # Open printing dialog
        dialog = QtPrintSupport.QPrintDialog()

        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.finalEdit.document().print_(dialog.printer())

    def cursorPosition(self):
        cursor = self.finalEdit.textCursor()

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
        color = QtWidgets.QColorDialog.getColor()

        # Set it as the new text color
        self.finalEdit.setTextColor(color)

    def highlight(self):
        color = QtWidgets.QColorDialog.getColor()

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
        # This function does two things:
        # Parse the string for template variables to change into the user's input
        # Parse the string for certain hardcoded strings that must be changed to something else

        configData = self.readConfig()

        # inputString = inputString.replace("{newline}", "\n")
        # inputString = inputString.replace("{tab}", "\t")
        # inputString = inputString.replace("&amp;nbsp", "&nbsp;")

        for key in configData:
            inputString = inputString.replace("{" + str(key) + "}", str(configData[key]))

        return inputString

    def templateSettings(self):
        count = 0
        keylist = []
        keyLabels = ""

        # table for editing template variable inputs
        self.window = QtWidgets.QWidget()
        self.table = QtWidgets.QTableWidget()
        self.tableItem = QtWidgets.QTableWidgetItem()

        # initialize data
        templateData = self.readConfig()

        # initiate table
        self.table.setRowCount(len(templateData.keys()))
        self.table.setColumnCount(1)

        # set data for template var editing table
        for key in templateData:
            self.table.setItem(count, 0, QtWidgets.QTableWidgetItem(templateData[key]))
            keyLabels = keyLabels + str(key) + ';'
            keylist.append(key)
            count+=1

        # set label, each element must be separated by a ;
        self.table.setHorizontalHeaderLabels("Var Value;".split(";"))
        self.table.setVerticalHeaderLabels(keyLabels.split(";"))

        # resize col
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)

        okayButton = QtWidgets.QPushButton('Okay', self)
        cancelButton = QtWidgets.QPushButton('Cancel', self)

        # must use a lambda or else the program does not think that connect is given a function
        okayButton.clicked.connect(lambda: self.templateSettingsOkayButton(keylist))
        okayButton.clicked.connect(self.window.close)
        cancelButton.clicked.connect(self.window.close)

        # attach items to layout
        # addWidget's 3rd parameter is rowSpan and the 4th is colSpan 
        layout = QtWidgets.QGridLayout(self.window)
        layout.addWidget(self.table, 0, 0, 4, 3)
        layout.addWidget(okayButton, 5, 1)
        layout.addWidget(cancelButton, 5, 2)

        self.window.setWindowTitle("Template Variable Settings")
        self.window.setGeometry(365, 400, 400, 250)
        self.window.show()

    def templateSettingsOkayButton(self, keys):
        # When the user presses okay, their input on the table is saved to the file so that it can
        # be used to substitute the user's input for the template variables

        configFile = open(self.config_path, 'w')

        for i in range(0, self.table.rowCount()):
            item = self.table.item(i, 0)
            outputStr = str(keys[i]) + ' = ' + item.text() + '\n'
            configFile.write(outputStr)

        configFile.close()

    def readConfig(self):
        # open the template.txt file and uses regex to get the variable name and the variable value

        tableVals = {}

        configFile = open(self.config_path, 'r')
        lines = configFile.readlines()
        lines = [line.rstrip() for line in lines] 
        for items in lines:  
            # regex group matching
            strCheck = r'(\S+)\s*=\s*(.*)'
            templateVars = re.search(strCheck, items)
            tableVals[templateVars.group(1)] = str(templateVars.group(2))

        configFile.close()
        return tableVals



def main():

    app = QtWidgets.QApplication(sys.argv)
    # QtGui.QApplication.setFont(QtGui.QFont('Times New Roman'))

    main = TemplateWriter()
    main.show()

    sys.exit(app.exec_())

if __name__ == "__main__":
  main()