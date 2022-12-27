from PyQt5 import QtCore, QtGui, QtWidgets
import sys, datetime, os
from GUI import Ui_Dialog
import main_program as mprog
import win32com.client as win32


projection = {"front view":(0, 1 , 0, 0 , 0, 1) , "Rear view":(0, -1, 0, 0 , 0, 1) , "top view":(0, 1 , 0, -1 , 0 , 0)
    , "bottom view":(0, 1, 0, 1, 0, 0), "Left view":(1, 0, 0, 0, 0, 1) , "right view":(-1, 0, 0, 0, 0, 1)
              , 'top view(left horizontal)':(-1 , 0 , 0 , 0 , -1 , 0)  , "top view(Y inverse)":(0 , -1 , 0 , -1 , 0 , 0) ,
              'bottom view(right horizontal)':(-1 , 0 , 0 , 0 , 1 , 0) , 'right view(left horizontal)':(0 , 0 , -1 , -1 , 0 , 0),
              'Rear view(left horizontal)':(0 , 0 , -1 , 0 , -1 , 0) , 'top view(right horizontal)':(1 , 0 , 0 , 0 , 1 , 0) ,
              'top view(180 degree)':(0 , -1 , 0 , 1 , 0 , 0) , 'bottom view(left horizontal)':(1 , 0 , 0 , 0 , -1 , 0)}
U = ["front view" , "Rear view" , "top view" , "bottom view" , "Left view" , "right view" , 'top view(left horizontal)' ,
     "top view(Y inverse)" , 'bottom view(right horizontal)' , 'right view(left horizontal)','Rear view(left horizontal)',
     'top view(right horizontal)' , 'top view(180 degree)' , 'bottom view(left horizontal)']




def Material_diagram_projection(surface , XLocation , YLocation , x , scale):
    catapp = win32.Dispatch('CATIA.Application')
    ActWin = catapp.Windows.item("Drawing1.CATDrawing")
    ActWin.Activate()
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews1 = drawingSheet.Views
    drawingView1 = drawingViews1.Add(x)
    drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    documents1 = catapp.Documents
    partDocument1 = documents1.Item(x + ".CATPart")
    product1 = partDocument1.GetItem(x)
    drawingViewGenerativeBehavior1.Document = product1
    drawingViewGenerativeBehavior1.DefineFrontView(projection[U[surface]][0], projection[U[surface]][1]
                                                           , projection[U[surface]][2], projection[U[surface]][3]
                                                           , projection[U[surface]][4], projection[U[surface]][5])
    drawingView1.X = XLocation
    drawingView1.Y = YLocation
    drawingView1.Scale = scale
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    drawingViewGenerativeBehavior1.Update()
    drawingView1.Activate()
    # drawingview2 = drawingViews1.Item('Front view')
    # drawingtexts1 = drawingview2.Texts
    # drawingtext1 = drawingtexts1.Item(1)
    # # drawingtexts1 = drawingtext1.Parent
    # selection = partDocument1.Selection
    # selection.Add(drawingtext1)
    # selection.Delete()
    # selection.Clear()
    drawingView1.FrameVisualization = False


def Material_diagram_balloons(view , name , XLocation , YLocation):
    catapp = win32.Dispatch("CATIA.Application")
    partdoc = catapp.ActiveDocument
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingviews = drawingsheet.Views
    drawingview = drawingviews.Item(view)
    drawingview.Activate()
    DrawTexts_balloons = drawingview.Texts
    DrawText = DrawTexts_balloons.Add(name, XLocation , YLocation)
    DrawText.SetFontName(0, 0, 'Arial Unicode MS (TrueType)')
    DrawText.SetFontSize(0, 0, 10)  # 調整字體位置和大小(x, y, 字體大小)
    # DrawLeader_DrawTexts_balloons = DrawText.Leaders.Add(point_positionX, point_positionY)  # 圓點位置
    DrawText.FrameType = 3  # 圓類型
    DrawText.x = XLocation
    # DrawText.y = YLocation
    # DrawText.DeactivateFrame(3)
    # DrawText.Deactivates = 1
    # DrawText.Blanking(0)
    # DrawLeader_DrawTexts_balloons.AllAround = 0
    # DrawLeader_DrawTexts_balloons.ModifyPoint(0, leader_line_length, 0)  # 選擇引線點 -> (0, 1, 2)並調整座標位置
    # DrawLeader_DrawTexts_balloons.HeadSymbol = 20  # 引號標點類型
    # DrawText.StandardBehavior = 1
    # MyViewGenBehavior = MyView.GenerativeBehavior




