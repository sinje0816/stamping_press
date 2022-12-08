from PyQt5 import QtCore, QtGui, QtWidgets
import sys, datetime, os
from GUI import Ui_Dialog
import main_program as mprog
import win32com.client as win32
import N_Test as N

projection = {"xy":(1, 0, 0, 0, 1, 0) , "-xy":(-1, 0, 0, 0, 1, 0) , "yz":(0, 1, 0, 0, 0, 1)
    , "-yz":(0, -1, 0, 0, 0, 1), "zx":(1, 0, 0, 0, 0, 1) , "-zx":(-1, 0, 0, 0, 0, 1)}
P = ["xy" , "-xy" , "yz" , "-yz" , "zx" , "-zx"]
p = ()
X1 = ()
Y1 = ()
X2 = ()
Y2 = ()
X3 = ()
Y3 = ()


mprog.set_CATIA_workbench_env()
mprog.OPEN_Drawing()
catapp = win32.Dispatch('CATIA.Application')
drawingDocument = catapp.ActiveDocument
drawingSheets = drawingDocument.Sheets
drawingSheet = drawingSheets.Item("Sheet.1")
drawingViews1 = drawingSheet.Views
projection_file_name_list = ['FRAME1', 'FRAME2']
for x in projection_file_name_list:
    mprog.import_part("C:\\Users\\USER\\Desktop\\stamping_press", x)
    if x == 'FRAME1':
        p = 4
        p2 = 5
        X1 = 145
        Y1 = 640
        X2 = 375
        Y2 = 640
        N.Select_projection_surface(p, X1, Y1)
        N.Select_projection_surface(p2, X2, Y2)
    elif x == 'FRAME2':
        p = 4
        p2 = 5
        X1 = 145
        Y1 = 300
        X2 = 375
        Y2 = 300
        N.Select_projection_surface(p, X1, Y1)
        N.Select_projection_surface(p2, X2, Y2)


# N.Select_projection_surface(p , X1 , Y1)
# N.Select_projection_surface(p2 , X2 , Y2)