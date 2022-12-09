from PyQt5 import QtCore, QtGui, QtWidgets
import sys, datetime, os
from GUI import Ui_Dialog
import main_program as mprog
import win32com.client as win32


projection = {"xy":(1, 0, 0, 0, 1, 0) , "-xy":(-1, 0, 0, 0, 1, 0) , "yz":(0, 1, 0, 0, 0, 1)
    , "-yz":(0, -1, 0, 0, 0, 1), "zx":(1, 0, 0, 0, 0, 1) , "-zx":(-1, 0, 0, 0, 0, 1)}
P = ["xy" , "-xy" , "yz" , "-yz" , "zx" , "-zx"]


mprog.set_CATIA_workbench_env()
mprog.OPEN_Drawing()
catapp = win32.Dispatch('CATIA.Application')
drawingDocument = catapp.ActiveDocument
drawingSheets = drawingDocument.Sheets
drawingSheet = drawingSheets.Item("Sheet.1")
drawingViews1 = drawingSheet.Views
projection_file_name_list = ['FRAME1']
for x in projection_file_name_list:
    mprog.import_part("C:\\Users\\USER\\Desktop\\stamping_press", x)
    if x == 'FRAME1':
        p = 4
        p2 = 5




def Select_projection_surface(surface , XLocation , YLocation):
    projection_file_name_list = ['FRAME1']
    for x in projection_file_name_list:
            drawingView1 = drawingViews1.Add("AutomaticNaming")
            drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks
            drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
            documents1 = catapp.Documents
            partDocument1 = documents1.Item(x + ".CATPart")
            product1 = partDocument1.GetItem(x)
            drawingViewGenerativeBehavior1.Document = product1
            drawingViewGenerativeBehavior1.DefineFrontView(projection[P[surface]][0], projection[P[surface]][1]
                                                           , projection[P[surface]][2], projection[P[surface]][3]
                                                           , projection[P[surface]][4], projection[P[surface]][5])
            drawingView1.X = XLocation
            drawingView1.Y = YLocation
            drawingView1.Scale = float(1 / 18)
            drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
            drawingViewGenerativeBehavior1.Update()


