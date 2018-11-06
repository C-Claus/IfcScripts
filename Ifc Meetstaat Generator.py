import sys
import os
import ctypes 

app_icon = "category_filter_32x32_dpi_90.png"
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_icon)
 

from PyQt4 import QtGui
from PyQt4 import QtCore 
from PyQt4 import QtWebKit

from PyQt4.QtCore import QThread, pyqtSignal
from ifc2excel import ifcfile, write_to_file

width = 400
height = 600

class External(QThread):
    """
    Runs a counter thread.
    """
    countChanged = pyqtSignal(int)
    
 
    def run(self):
        
        filename = QtGui.QFileDialog.getOpenFileName(None, 'Open IFC bestand','', 'IFC File (*.ifc)')
            
          
        if len(filename) > 0:  
            self.countChanged.emit(0)
           
          
            if len(ifcfile(filename=filename)) > 0:
                self.countChanged.emit(1)
                
                write_to_file(export_file=filename)
                
                #sys.exit()
                

class Actions(QtGui.QMainWindow):
    """
    Simple dialog that consists of a Progress Bar and a Button.
    Clicking on the button results in the start of a timer and
    updates the progress bar.
    """
    def __init__(self):
        super(Actions, self).__init__()
    
        self.initUI()

    def initUI(self):
        
        self.setWindowIcon(QtGui.QIcon(app_icon)) 


        #chb_red = 'rgb(232, 52, 38)'
        #chb_grey = 'rgb(237,237,237)'
        #white = 'rgb(255, 255, 255)'
        #black = 'rgb(0, 0, 0)'

        
        self.setWindowTitle('Coen Hagedoorn Bouwgroep')
        self.setFixedHeight(height)
        self.setFixedWidth(width)
        
       
        self.view  = QtWebKit.QWebView(self)
        self.view.setHtml(open('gui.html').read())
        self.view.setFixedWidth(width+50)
        self.view.setFixedHeight(550)
        self.view.move(-10,-10)
        
        
        self.label = QtGui.QLabel('', self)
        self.label.move(50,490)
        self.label.setFixedHeight(50)
        self.label.setFixedWidth(width+50) 


        self.progress = QtGui.QProgressBar(self)
        self.progress.setGeometry(0, 0, width+50, 20) 
        #self.progress.setMinimum(0)
        #self.progress.setMaximum(100)
        self.progress.setFixedWidth(width+33)
        self.progress.move(1,530)
        
        self.button = QtGui.QPushButton('Exporteer .IFC bestand naar Excel', self)
        self.button.move(0, 550)
        self.button.setFixedWidth(width)
        self.show()

        self.button.clicked.connect(self.onButtonClick)

    def onButtonClick(self):
        
        self.calc = External()
        self.calc.countChanged.connect(self.onCountChanged)
        self.calc.start()
        
    def onCountChanged(self, value):
        
        if value == 0:
            self.progress.setMaximum(value)
            self.label.setText('...IFC data wordt verwerkt en weggeschreven naar Excel.')
        
        if value == 1: 
            self.progress.setRange(0,1)
            self.label.setText('...Het Excel bestand is gereed.')
            


if __name__ == "__main__":

    app = QtGui.QApplication(sys.argv)

    window = Actions()
    window.show()
    app.exec_()
    sys.exit(app.exec_())
             

    