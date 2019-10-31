#==========================================================================================================================
# Authors: Dr. Eaglin, Andrew Polemeni
# Organization: Daytona State College
# Date Written: October 24, 2019
# Program Purpose: To automate course check downs for student progress.
#==========================================================================================================================
# IMPORT STATEMENTS
import sys, os
import generateReports

import PyQt5
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QAction, QLineEdit, QMessageBox, QFileDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import YELLOW, BLACK

# UI STARTS HERE

class App(QMainWindow):

    def __init__(self):
        # Here we initialize the UI attributes such as hight, width, title, ect.
        super().__init__()
        self.title = 'Daytona State College - Student Check Down Generator'
        self.left = 600
        self.top = 300
        self.width = 500
        self.height = 300
        self.initUI()
    
    def initUI(self):
        # Here is were we call the attributes to actually display them.
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        #self.setWindowIcon(QIcon('images/icon.png'))

        # CREATE A BUTTON TO OPEN THE FILE
        self.button = QPushButton('1) Open the query file: ', self)
        self.button.move(100, 50) # the placement of the button in the application
        self.button.resize(200, 50) # the size of the button
        self.button.clicked.connect(self.openFileNameDialog) # connect the button to the function

        # CREATE A BUTTON TO GENERATE THE REPORT
        self.button = QPushButton('2) Genereate the reports', self)
        self.button.move(100, 150)
        self.button.resize(200, 50)
        self.button.clicked.connect(generateReports.GenerateAllSpreadsheets)

        # Call functions here to execute them in the UI
        #self.openFileNameDialog()
        
        self.show() # Here is how we actually show the UI
    
      
    def openFileNameDialog(self, filename1):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        self.filename1, _ = QFileDialog.getOpenFileName(self, "Open the Query file to generate student checkdowns", "", "Excel Files (*.xlsx)", options=options)
        if self.filename1:
            print(self.filename1)
        return self.filename1

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())

#==========================================================================================================================
#UI ENDS HERE
def resource_path(relative_path):
  if hasattr(sys, '_MEIPASS'):
      return os.path.join(sys._MEIPASS, relative_path)
  return os.path.join(os.path.abspath('.'), relative_path)
#==========================================================================================================================
