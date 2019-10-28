#import grabrows
import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QAction, QLineEdit, QMessageBox, QFileDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot

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
        self.setWindowIcon(QIcon('images/icon.png'))

        # CREATE A BUTTON TO OPEN THE FILE
        self.button = QPushButton('1) Open the query file: ', self)
        self.button.move(100, 50) # the placement of the button in the application
        self.button.resize(200, 50) # the size of the button
        self.button.clicked.connect(self.openFileNameDialog) # connect the button to the function

        # CREATE A BUTTON TO GENERATE THE REPORT
        self.button = QPushButton('2) Genereate the reports', self)
        self.button.move(100, 150)
        self.button.resize(200, 50)

        # Call functions here to execute them in the UI
        #self.openFileNameDialog()
        #self.openFileNamesDialog()
        #self.saveFileDialog()
        
        self.show() # Here is how we actually show the UI
    
    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Open the Query file to generate student checkdowns", "","Excel Files (*.xlsx)", options=options)
        if fileName:
            print(fileName)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
