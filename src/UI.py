import string
import sys, os 

from PyQt4.QtCore import *
from PyQt4.QtGui import *
# import matplotlib
from matplotlib.backends.backend_qt4agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt4agg import NavigationToolbar2QTAgg as NavigationToolbar
from matplotlib.figure import Figure
import numpy
from numpy.random import normal, exponential
from numpy import average
import scipy.stats
import multiprocessing


class AppForm(QMainWindow):
    def __init__(self,op,inp,sel,parent=None):
        QMainWindow.__init__(self, parent)
#         NavigationToolbar.__init__(self,self.canvas, parent)
        global input
        input = inp
        global output
        output=None
        output = op
        global bin
        bin = 10
        global selection
        selection = sel
        self.setWindowTitle("MiSim Application")
        self.create_menu()
        self.create_main_frame()
        self.create_status_bar()
        global xlow
        global xup
        global upd_lim
        upd_lim = False 

    def create_menu(self):
        self.file_menu = self.menuBar().addMenu("&File")
        
        load_file_action = self.create_action("&Save plot",
            shortcut="Ctrl+S", slot=self.save_plot, 
            tip="Save the plot")
        quit_action = self.create_action("&Quit", slot=self.close, 
            shortcut="Ctrl+Q", tip="Close the application")
        
        self.add_actions(self.file_menu, 
            (load_file_action, None, quit_action))
        
        self.help_menu = self.menuBar().addMenu("&Help")
        about_action = self.create_action("&About", 
            shortcut='F1', slot=self.on_about, 
            tip='About the demo')
        
        self.add_actions(self.help_menu, (about_action,))

    def add_actions(self, target, actions):
        for action in actions:
            if action is None:
                target.addSeparator()
            else:
                target.addAction(action)

    def create_action(  self, text, slot=None, shortcut=None, 
                        icon=None, tip=None, checkable=False, 
                        signal="triggered()"):
        action = QAction(text, self)
        if icon is not None:
            action.setIcon(QIcon(":/%s.png" % icon))
        if shortcut is not None:
            action.setShortcut(shortcut)
        if tip is not None:
            action.setToolTip(tip)
            action.setStatusTip(tip)
        if slot is not None:
            self.connect(action, SIGNAL(signal), slot)
        if checkable:
            action.setCheckable(True)
        return action

    def on_about(self):
        msg = """ MiSim - Developed by Kasa and Baffy under the guidance of Dr. Roehrig
        Version 1.0
        
        Contact kasa@cmu.edu for any suggestions/bugs.
        """
        QMessageBox.about(self, "About the demo", msg.strip())

    def save_plot(self):
        file_choices = "PNG (*.png)|*.png"
        
        path = unicode(QFileDialog.getSaveFileName(self, 
                        'Save file', '', 
                        file_choices))
        if path:
            self.canvas.print_figure(path, dpi=self.dpi)
            self.statusBar().showMessage('Saved to %s' % path, 2000)

    
    def create_main_frame(self):
        global input
        self.main_frame = QWidget()
        
        
        self.dropdown = QComboBox()
        self.dropdown.setMaximumWidth(100)
        
        for i in input:
            self.dropdown.addItem("Input - "+str(i))
        
        self.dpi = 100
        self.fig = Figure((8.0, 7.0), dpi=self.dpi)
        self.canvas = FigureCanvas(self.fig)
        self.canvas.setParent(self.main_frame)
        
        self.axes = self.fig.add_subplot(111)
        global counts, bins, patches
        counts, bins, patches = self.axes.hist(output[0],bins=50) 
        # Create the navigation toolbar, tied to the canvas
        #
        self.mpl_toolbar = NavToolbar(self.canvas, self.main_frame)
        
        self.label1 = QLabel("Enter the number of bins")
        self.textbox = QLineEdit()
        self.textbox.setMaximumWidth(50)
        self.textbox.setText('50')
        self.xlimlabel = QLabel("X Axis Limits")
        self.xlowlim = QLineEdit()
        self.xlowlim.setMaximumWidth(50)
        self.xuplim = QLineEdit()
        self.xuplim.setMaximumWidth(50)
        self.xdash = QLabel("-")
        self.selop = QLabel("select the output")
        self.limerrortext = QLabel()
        self.limerrortext.setVisible(False)
        self.limerrortext.setStyleSheet('color:red')
        
        self.problabel = QLabel("Percentile")
        self.probvalue = QLineEdit()
        self.probvalue.setReadOnly(True)
        
        self.probvalue.setMaximumWidth(50)
        self.probvalue.setStyleSheet("background-color: #c0c0c0 ;")

        
        self.rdbutton1 = QRadioButton('Interval Limits')
#         self.rdbutton2 = QRadioButton('Percentile')
        
        self.rdgroup = QButtonGroup()
        self.rdgroup.addButton(self.rdbutton1)
#         self.rdgroup.addButton(self.rdbutton2)
        self.rdbutton1.setChecked(True)
        
        self.rdbutton1.toggled.connect(self.chk_rdbuttons)
#         self.rdbutton2.toggled.connect(self.chk_rdbuttons)
        
        self.connect(self.xlowlim, SIGNAL('editingFinished()'),self.check_lim)
        self.connect(self.xuplim, SIGNAL('editingFinished()'),self.check_lim)
        self.connect(self.textbox, SIGNAL('editingFinished ()'), self.on_draw)
        
        
        
        self.pdf_cb = QCheckBox("Prob. Dist.")
        self.pdf_cb.setChecked(False)
        self.connect(self.pdf_cb, SIGNAL('stateChanged(int)'), self.on_draw)
        
        self.cum_pdf_cb = QCheckBox("Cum Prob. Dist.")
        self.cum_pdf_cb.setChecked(False)
        self.cum_pdf_cb.setShown(False)
        self.connect(self.cum_pdf_cb, SIGNAL('stateChanged(int)'), self.on_draw)
        self.connect(self.pdf_cb, SIGNAL('toggled(bool)'),self.cum_pdf_cb.setVisible)
        
        self.grid_cb = QCheckBox("Show &Grid")
        self.grid_cb.setChecked(False)
        self.connect(self.grid_cb, SIGNAL('stateChanged(int)'), self.on_draw)
        
        self.dropdown.connect(self.dropdown, SIGNAL("currentIndexChanged(int)"),self.refresh)
        
        self.textarea = QScrollArea()
        self.text = QTableWidget()
        self.text.setWindowTitle("Statistics")
#         self.text.resize(400,250)
        self.text.resize(self.text.sizeHint())
        self.text.setRowCount(len(output))
#        print min(gaussian_numbers[1])
        self.text.setColumnCount(9)
        self.text.setHorizontalHeaderLabels(QString("Input Value;Mean;Median;Std Dev;Variance;Skewness;Kurtosis;Maximum;Minimum").split(";"))
        self.text.setWindowTitle
        
#         someString = str(min(output[1]))
        
 #       self.text.setItem(0,1,QTableWidgetItem(someString))
 #       self.text.setItem(0,0,QTableWidgetItem(str(stats.mode(gaussian_numbers[1]))))
        rowcount=0
        for arrElement in output:
#             print max(arrElement)
            self.text.setItem(rowcount,7,QTableWidgetItem(str(max(output[rowcount]))))
            self.text.setItem(rowcount,8,QTableWidgetItem(str(min(output[rowcount]))))
            self.text.setItem(rowcount,1,QTableWidgetItem(str(average(output[rowcount]))))
            self.text.setItem(rowcount,2,QTableWidgetItem(str(numpy.median(output[rowcount]))))
            self.text.setItem(rowcount,3,QTableWidgetItem(str(numpy.std(output[rowcount]))))
            self.text.setItem(rowcount,4,QTableWidgetItem(str(numpy.var(output[rowcount]))))
            self.text.setItem(rowcount,5,QTableWidgetItem(str(scipy.stats.skew((output[rowcount])))))
            self.text.setItem(rowcount,6,QTableWidgetItem(str(scipy.stats.kurtosis((output[rowcount])))))
            self.text.setItem(rowcount,0,QTableWidgetItem(str(input[rowcount])))
            rowcount=rowcount + 1
        
        self.text.setEditTriggers(QAbstractItemView.NoEditTriggers)
        
        self.copyAction = QAction("Copy",  self)
        self.copyAction.setShortcut("Ctrl+C")
        self.addAction(self.copyAction)
        self.connect(self.copyAction, SIGNAL("triggered()"), self.on_select)
        
        global selection
        rows = self.text.rowCount()
#         for x in range(rows):
# #             print x
#             item = self.text.item(x,int(selection))
#             item.setBackgroundColor(QColor(255, 0, 0, 127))
#          
#         self.text.show()
        self.textarea.setWidget(self.text)
        
        hb = QHBoxLayout()
        hb.addWidget(self.selop)
        hb.addWidget(self.dropdown)
        hb.addWidget(self.label1)
        hb.addWidget(self.textbox)
        hb.addWidget(self.pdf_cb)
        hb.addWidget(self.cum_pdf_cb)
        hb.addWidget(self.grid_cb)
#         hb.addWidget(self.rdbutton1)
        
        hb.addWidget(self.xlimlabel)
        
        hb.addWidget(self.xlowlim)
        
        hb.addWidget(self.xdash)
        
        hb.addWidget(self.xuplim)
        hb.addWidget(self.problabel)
        hb.addWidget(self.probvalue)
        
          
        vbox = QVBoxLayout()
        vbox.addWidget(self.canvas)
        vbox.addWidget(self.mpl_toolbar)
        vbox.addLayout(hb)
        vbox.addWidget(self.limerrortext)
        vbox.addWidget(self.text)
        
        self.main_frame.setLayout(vbox)
        self.setCentralWidget(self.main_frame)

    
    def chk_rdbuttons(self):
        if (self.rdbutton1.isChecked()):
            self.probvalue.setText('')
            self.probvalue.setReadOnly(True)
            self.xlowlim.setReadOnly(False)
            self.xuplim.setReadOnly(False)
        
        self.check_lim()
            
            
    def refresh(self):
        global upd_lim
        self.xlowlim.setText('')
        self.xuplim.setText('')
        self.probvalue.setText('')
        upd_lim = False
        
        self.on_draw()
    
    def on_select(self):
        self.clip = QApplication.clipboard()
        print'on Selection' 
        selected = self.text.selectedRanges()
        s=''
        for r in xrange(selected[0].topRow(),selected[0].bottomRow()+1):
            for c in xrange(selected[0].leftColumn(),selected[0].rightColumn()+1):
                try:
                    print 'appending'
                    s += str(self.text.item(r,c).text()) + "\t"
                except AttributeError:
                    print 'exception'
                    s += "\t"
            s = s[:-1] + "\n" #eliminate last '\t'
        
        self.clip.setText(s)
                    
       
    def check_lim(self):
        xlowtext = str(self.xlowlim.text())
        xuptext = str(self.xuplim.text())
        xcitext = str(self.problabel.text())
        global upd_lim
        global xlow
        global xup
        global xci
        global bins
        global upd_flag
        
        if      (xuptext=='' and xlowtext!=''):
            upd_flag= 'L'
            xlow = float(xlowtext)
            xup = bins[-1]
         
        if      (xuptext!='' and xlowtext==''):
            upd_flag= 'L'
            xup = float(xuptext)
            xlow = bins[0]
                
        if       (xuptext!='' and xlowtext!=''):
            upd_flag= 'L'
            if (float(xlowtext)<bins[1]):
                xlow = bins[0]
            else:
                xlow = float(xlowtext)
            
            if (float(xuptext)>bins[-1]):
                xup = bins[-1]
            else:               
                xup = float(xuptext)
        
        upd_lim = True

        self.on_draw()
 
    
    def create_status_bar(self):
        self.status_text = QLabel("This is a demo")
        self.statusBar().addWidget(self.status_text, 1)

    def on_draw(self):
        """ Redraws the figure
        """
        text=str(self.textbox.text())
        
        global bin
        if text=='':
            text='50'
        bin = int(text)
            
        self.axes.clear()        
        inputvalstr = self.dropdown.currentIndex()
        inputval = int(inputvalstr)
        
        global upd_lim
        global upd_flag
        global xlow
        global xup
        global xci
        global counts, bins, patches 
        
        array = output[inputval]
        
        if self.pdf_cb.isChecked():
            if self.cum_pdf_cb.isChecked():
                counts, bins, patches =self.axes.hist(array,color = 'b', bins =bin,normed = True,cumulative= True)
            else:
                
                counts, bins, patches = self.axes.hist(array,color = 'b',bins =bin,normed = True)
        else:   
            
                counts, bins, patches = self.axes.hist(array,bins =bin,color = 'b')
        
        
        if (upd_lim):
#             if (upd_flag=='L'):
                self.axes.axvline(xlow,color = 'r',linestyle='dashed', linewidth=1)
                self.axes.axvline(xup,color = 'r',linestyle='dashed', linewidth=1)
                for patch, rightside, leftside in zip(patches, bins[1:], bins[:-1]):
                    if rightside < xlow:
                        patch.set_facecolor('green')
                    elif leftside > xup:
                        patch.set_facecolor('red')
                bins1 =(bins[0],xlow,xup,bins[-1])
                hist1,edges = numpy.histogram(array, bins1)
                self.probvalue.setText(str(hist1[1]*100/len(array))+"%")
            
#             if (upd_lim=='C'):
#                 xclow = (1-xci)/2
#                 xcup = xci + xclow
#                 cires = stats.rv_discrete.ppf(xci)
                
            
        self.axes.grid(self.grid_cb.isChecked())        
#         
        self.canvas.draw()
 
class NavToolbar(NavigationToolbar ):
    def __init__(self, canvas, parent ):
        NavigationToolbar.__init__(self,canvas,parent)
        for c in self.findChildren(QToolButton):
            if str(c.text()) in ('Subplots','Customize','Forward','Back'):
                c.defaultAction().setVisible(False)
                continue 
        
 
def main(op):
    app = QApplication(sys.argv)
    
    form = AppForm(op)
    form.show()
    app.quit()
    status = app.exec_()
    sys.exit(status)


def f(op,inp,usr_selection):
    app = QApplication(sys.argv)
    form = AppForm(op,inp,usr_selection)
    form.show()
    status =app.exec_()

def draw(op,inp,usr_selection):
    # separate process so that excel is not blocked
    multiprocessing.set_executable(os.path.join(sys.exec_prefix, 'pythonw.exe')) 
    p = multiprocessing.Process(target=f, args=(op,inp,usr_selection))
    p.start()

 

if __name__ == "__main__":
    op=[]
    main(op)               