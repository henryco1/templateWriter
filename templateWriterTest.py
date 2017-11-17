import sys
import unittest
from PyQt5.QtWidgets import QApplication
from PyQt5.QtTest import QTest
from PyQt5.QtCore import Qt
import templateWriter


class TemplateWriterTest(unittest.TestCase):
    def test_defaults(self):
        # newTest = self.form.ui.newAction
        # openTest = self.form.ui.openAction

        QTest.mouseClick(templateWriter.newAction)
        # QTest.mouseClick(openTest)

if __name__ == "__main__":
    unittest.main()