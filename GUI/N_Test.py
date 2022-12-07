from PyQt5 import QtCore, QtGui, QtWidgets
import sys, datetime, os
from GUI import Ui_Dialog
import main_program as mprog
import win32com.client as win32


mprog.set_CATIA_workbench_env()
mprog.OPEN_Drawing()
catapp = win32.Dispatch('CATIA.Application')
drawingDocument = catapp.ActiveDocument
drawingSheets = drawingDocument.Sheets
drawingSheet = drawingSheets.Item("Sheet.1")
drawingViews1 = drawingSheet.Views
drawingView1 = drawingViews1.Add("AutomaticNaming")
drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks
drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
drawingSheet.Scale = float(1/18)
documents1 = catapp.Documents
partDocument1 = documents1.Item("FRAME1.CATPart")
product1 = partDocument1.GetItem("FRAME1")
drawingViewGenerativeBehavior1.Document = product1
drawingViewGenerativeBehavior1.DefineFrontView (-1, -0, -0, 0, 0, 1)
drawingView1.X = 375
drawingView1.Y = 640
drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
drawingViewGenerativeBehavior1.Update()

