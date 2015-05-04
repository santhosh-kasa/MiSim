'''
Created on Aug 26, 2014

@author: santhosh
'''
from pyxll import xl_func, xl_menu,xlcCalculateNow,xl_on_reload,xlcCalculation,xl_on_open,xlcAlert,xlcCalculateDocument,get_active_object,xlfCaller,xlCalculationManual,\
    get_config
import win32com.client
import csv
import datetime
import time
from numpy import array
from numpy import zeros
from numpy import nditer
import numpy
import os 
from PyQt4 import QtGui, QtCore
import timer
import pyxll
import sys
import UI
import threading
from PyQt4.QtGui import QApplication    
import Dialog

import numpy.random as np  
from _ast import Pass

@xl_func("float trials, float success_prob: float",category="MiSim",macro=True,volatile=True)
def MiBinom(trials, success_prob):
    """Binomial Distribution Function
        Trials: Number of trials
        Success_prob : Probability of Success """
    global manflag
    if not manflag:
        setManual()    
    return np.binomial(trials,success_prob,1)

@xl_func("float start, float end: float",category="MiSim",volatile=True,macro=True)
def MiUniform(start,end):
    """Uniform Distribution Function
        start: Minimum
        end: Maximum"""
    global manflag
    if not manflag:
        setManual()    
    return np.uniform(start,end,1)

@xl_func("float rate: float",category="MiSim",volatile=True,macro=True)
def MiExponential(rate):
    """Exponential Distribution Function
        rate: Gamma of exp distribution"""
    global manflag
    if not manflag:
        setManual()
    return np.exponential(rate,1) 

@xl_func("float dof: float",category="MiSim",volatile=True,macro=True)
def MiChisquare(dof):
    """ChiSquare Distribution Function
        dof: Degrees of Freedom"""
    global manflag
    if not manflag:
        setManual()
    return np.chisquare(dof,1)

@xl_func("float left, float mode, float right: float",category="MiSim",volatile=True,macro=True)
def MiTriangular(left, mode, right):
    """Triangular Distribution Function
        left: Lower Limit
        mode: Mode
        right: Upper Limit"""
    global manflag
    if not manflag:
        setManual()
    return np.triangular(left,mode,right,1)

@xl_func("float mean, float sd: float",category="MiSim",volatile=True,macro=True)
def MiLognormal(mean,sd):
    """LogNormal Distribution Function
        mean: Mean
        sd: Standard Distribution"""
    global manflag
    if not manflag:
        setManual()
    return np.lognormal(mean,sd,1)

@xl_func("float mean, float sd: float",category="MiSim",volatile=True,macro=True)
def MiNormal(mean,sd):
    """Normal Distribution Function
        mean: Mean
        sd: Standard Distribution"""
#     xlcCalculation(3) 
    global manflag
    if not manflag:
        setManual()
    return np.normal(mean,sd,1)
    
@xl_func("float lam: float",category="MiSim",volatile=True,macro=True)
def MiPoisson(lam):
    """Poisson Distribution Function
        lam: Expectation of interval"""
    global manflag
    if not manflag:
        setManual()
    return np.poisson(lam,1)

@xl_func("string name: string",macro=True)
def hello(name):
    """return a familiar greeting"""
#     return "Hello, %s" % name
    global op
    s=''
    for x in nditer(op):
        s = s + "," + str(x)
    return s


@xl_func("string name: string",macro=True)
def hello1(name):
    """return a familiar greeting"""
#     return "Hello, %s" % name
    global op2
    s=''
    for x in nditer(op2[0]):
        s = s + "," + str(x)
    return s


@xl_func("string name: string",macro=True)
def hello2(name):
    """return a familiar greeting"""
#     return "Hello, %s" % name
    global op2
    s=''
    for x in nditer(op2[1]):
        s = s + "," + str(x)
    return s

  
@xl_func("float value:float",macro=True,volatile=True,category="MiSim",)
def MiOutput(value):
    global op
    global current_inp
    global iter_num
    global output_flag
    op[current_inp][iter_num] = value
    output_flag = True 
    return value


@xl_func("int index, float value:float",macro=True,volatile=True,category="MiSim",)
def MiOutput2(index,value):
    global op2
    global col2
#     col2 = col2+1
    op2[index].append(value)   
    return value

  
#     xlcCalculateDocument()
#     xlcCalculateNow()
#     
# 

# @xl_func("float value:string")
# def Mistestnum(value):
#     testarray = zeros((1,3))
#     testarray[0][0] = 1.1
#     testarray[0][1]=1.2
#     testarray[0][2] = value
#     s=''
#     for x in nditer(testarray):
#         s = s + "  " + str(x)
#     return s
#    

@xl_func("float[] array:string",macro=True,volatile=True,category="MiSim")
def MiInputArray(array):
    global inp
    global arrayvals
    arrayvals = array 
    inp=[]
    for row in array:
        for cell_value in row:
            inp.append(cell_value)
    global num_of_inp        
    num_of_inp = len(inp)   
    global inparraycell
    inparraycell = xlfCaller()  
    return "Array of Input(s)"


@xl_func("xl_cell cell:float",macro=True,volatile=True,category="MiSim")
def MiInput(cell):
    global inp
    global inpcell
    global formstr
    global current_inp 
    global arrayvals
    inpcell = xlfCaller()
    formstr = str(inpcell.formula)
#
#     inp=[]
#     for row in arrayvals:
#         for cell_value in row:
#             inp.append(cell_value)
#     rect1 = inpcell.rect
#     xlcAlert(str(rect1.first_row))
#     xlcAlert(str(rect1.first_col))
    return inp[current_inp]
    

@xl_menu("Load",menu="Load", menu_order=1)
def on_Load():
    xlcCalculation(3) 
    xlcCalculateNow()
#     global op
#     s=''
#     for x in nditer(op):
#         s = s + "," + str(x)
#     xlcAlert(s)
    global simcomplete
    simcomplete = False

@xl_on_reload
def on_reload(import_info):
    global op
    op=[]
    op= zeros((1,1))
    global simcomplete
    simcomplete = False
    xlcCalculation(3)
    xlcCalculateNow()
    xlcCalculateDocument()
    setManual()


@xl_on_open   
def xl_on_open(func):
#     global op
#     op=[]
#     xlcCalculation(3)
    global manflag
    manflag = False
    global op
    op = zeros((1,1)) 
    global iter_num
    iter_num = 0
    xlcCalculateNow()
    global inp
    inp = []
    global num_of_inp
    num_of_inp = 1
    global current_inp
    current_inp = 0
    global inpcell
    inpcell = None
    global output_flag
    output_flag  = False
    global simcomplete
    simcomplete = False
    global op2
    op2 = zeros((2,100))
    global col2
    col2=0
    
    
def setManual():
#     global manflag
#     manflag = True
#     xl_window = get_active_object()
#     xl_app = win32com.client.Dispatch(xl_window).Application
#     xl_app.Calculation=xlCalculationManual   
    pass

    
# @xl_menu("Show_Stats",menu="Show Stats", menu_order=4)
# def Show_Stats():
#     global simcomplete
#     xl_window = get_active_object()
#     xl_app = win32com.client.Dispatch(xl_window).Application
# 
#     # get the current selected range
#     a= (10,20)
#     xl_app.ActiveWorkbook.Names.Add("test",a)


    
@xl_menu("Simulate",menu="Simulate", menu_order=3)
def Simulate():
    global num_of_sim 
    global op  
    global num_of_inp
    global iter_num
    global current_inp
    global inpcell
    global output_flag
    global usr_selection
    global inparraycell
    global formstr
    global arrayvals
    global simcomplete
    global inp
    
    xlcCalculateNow()
    xlcCalculateDocument()
    if not output_flag:
        xlcAlert("Select a output cell. If selected, click reload")
        return
#     
    checkMIArray=False
    win32com.client.gencache.Rebuild()
    xl_window = get_active_object()
    xl_app = win32com.client.Dispatch(xl_window).Application
    xl_app.Calculation=xlCalculationManual
#     xl_app = win32com.client.GetActiveObject("Excel.Application")
#     win32com.client.gencache.EnsureDispatch(xl_app)
    
    num_of_sim = 0
    if ((inpcell is not None) and formstr == str(inpcell.formula) ): 
        checkMIArray=True
#         xlcAlert('inside looop:'+ str(inp).strip('[]'))
        rect = inpcell.rect
#         xlcAlert(inpcell.address)
#         xlcAlert(str(rect.first_row))
#         xlcAlert(str(rect.first_col))
        selection = xl_app.ActiveSheet.Cells(int(rect.first_row)+1,int(rect.first_col)+1)
#         selection.value = 100
#     inpcell.value = 100
#     xlcAlert("click OK")
    
    app1 = QApplication(sys.argv)
    form = Dialog.popup()
    app1.exec_()
    num_of_sim = int(Dialog.retval())
    usr_selection = str(Dialog.retsel())
    writeFG = Dialog.writeflag()
    if num_of_sim > 0:
        current_inp =1
#         xlcAlert(str(num_of_sim))
        start = time.time()
        op=None
        op = zeros((num_of_inp,num_of_sim))        
        xl_app.ScreenUpdating = False
        xl_app.DisplayStatusBar = False
        xl_app.EnableEvents = False
        for j in range(num_of_inp):
            current_inp = j
            if (checkMIArray):
                selection.Value = inp[current_inp]
            for i in range(num_of_sim):
                iter_num = i
                xlcCalculateDocument()
        end = time.time()
        xl_app.ScreenUpdating = True
        xl_app.DisplayStatusBar = True
        xl_app.EnableEvents = True
        current_inp = 0
        simcomplete= True
#         selection.value = inp[current_inp]
#         selection.Formula = inpcell.formula
        if checkMIArray:    
            fstr = '=MiInput(' + inparraycell.address + ')'
            selection.Formula = fstr
#         xlcAlert(str(checkMIArray))
        if not checkMIArray:
            inp = [0]
        UI.draw(op, inp,usr_selection)
        # Store data in a CSV format
        if checkMIArray:
            popupstr = ''
            for idx,inpvalue in enumerate(inp):
                tupleop = None
                tupleop = tuple(op[idx])
                datastr = "data" + str(idx)
                xl_app.ActiveWorkbook.Names.Add(datastr,tupleop,False)
                popupstr = popupstr + "Output variable for Input " + str(inpvalue) + "is: "+str(datastr) + "\n"
            popupstr = popupstr + "You can use all the excel statistical functions on these variables"
            xlcAlert(popupstr) 
        else:
            tupleop = tuple(op)
            xl_app.ActiveWorkbook.Names.Add("data",tupleop,False)
            xlcAlert("Your Output Variable is 'data'" + "\n" + "You can use all the excel statistical functions on this variable")
        if writeFG:
            config = get_config()
            config.read('pyxll.cfg')
            dir_path = config.get('LOG','path')
            xlcAlert("Data stored at "+str(dir_path))
            file_name = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            if checkMIArray:
#                 xlcAlert(str(len(op[1])))
                for idx,inpvalue in enumerate(inp):
                    file_name1 = file_name +"-input "+str(inpvalue) +'.csv'
                    if os.path.exists(dir_path): 
                        myfile = open(os.path.join(dir_path, file_name1), 'wb')
                        wr = csv.writer(myfile, dialect='excel')
                        wr.writerow(op[idx])
            else:
                if os.path.exists(dir_path):
                    file_name = file_name + '.csv' 
                    myfile = open(os.path.join(dir_path, file_name), 'wb')
                    wr = csv.writer(myfile, dialect='excel')
                    wr.writerows(op)
            
        