'''
Created on Sep 3, 2014

@author: santhosh
'''

import sys, os 
from PyQt4.QtCore import *
from PyQt4.QtGui import *

class popup(QWidget):
    def __init__(self, parent=None):
        global writeFG
        writeFG = False
        global val
        val =0
        global selection
        selection = "Maximum"
        QWidget.__init__(self, parent)
        self.setWindowTitle('Demo: PyQt with matplotlib')
        self.main_frame = QWidget()
        self.label = QLabel("Enter the number of simulations")
        self.errorlabel = QLabel("Enter the number of simulations")
        self.errorlabel.setStyleSheet('color:red')
        self.errorlabel.setVisible(False)
        self.line = QLineEdit("")
        self.line.setMaximumWidth(60)
        self.button = QPushButton("Run")
        self.connect(self.button,SIGNAL("clicked()"),self.on_click)
        self.dropdown = QComboBox()
        self.dropdown.setToolTip("Select the statistic you want to highlight")
        
        self.dropdown.addItem("Maximum")
        self.dropdown.addItem("Minimum")
        self.dropdown.addItem("Mean")
        self.dropdown.addItem("Median")
        self.dropdown.addItem("Mode")
        self.writecb = QCheckBox("Write data into new spreadsheet")
        
        
        hbox = QHBoxLayout()
        hbox.addWidget(self.label)
        hbox.setAlignment(self.label, Qt.AlignLeft)
        
        hbox.addWidget(self.line)
        hbox.setAlignment(self.line,Qt.AlignCenter)
        
#         hbox.addWidget(self.dropdown)
#         hbox.setAlignment(self.dropdown, Qt.AlignCenter)
        
        hbox.addWidget(self.writecb)
        hbox.setAlignment(self.dropdown, Qt.AlignRight)
        
        vbox = QVBoxLayout()
        vbox.addLayout(hbox)
        vbox.addWidget(self.button)
        vbox.addWidget(self.errorlabel)
        self.main_frame.setLayout(vbox)
        
        self.main_frame.show()
        
    def on_click(self):
        text = str(self.line.text())
        
        if not text:
            self.errorlabel.setVisible(True)
        else:
            self.errorlabel.setVisible(False)
            global val
            global selection
            global writeFG
#             val = str(0)
            val = int(text)
            if self.writecb.isChecked():
                writeFG = True
            else:
                writeFG = False 
            selection = str(self.dropdown.currentText())
            self.main_frame.close()
    
def main():
    try :
        app = QApplication(sys.argv)
        form = popup()
        app.exec_()
    except:
        pass
    
def retval():
    global val
    return val 

def writeflag():
    global writeFG
    return writeFG

def retsel():
    global selection
    if selection=='Maximum':
        return 0
    if selection=='Minimum':
        return 1
    if selection=='Mean':
        return 2
    if selection=='Median':
        return 3
    if selection=='Mode':
        return 4
    return 9

if __name__ == "__main__":
    global val, selection
    main()
    print val