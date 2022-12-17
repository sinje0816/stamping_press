import os
import win32com.client as win32
import datetime, time, math
import main_program as mprog

draft_area_center_initX = 500
draft_area_center_initY = 220
draft_X_clearence = 10
draft_Y_clearence = 25
# 圖面最大範圍框
drafting_view_min_Y = 43
drafting_view_max_Y = 810
drafting_view_min_X = 20
drafting_view_max_X = 982
isometric_view_center_X = 270
isometric_view_center_Y = 600

def drafting_parameter_calculation(width, height, depth, S, T):  # 電子型錄WHD, 比例
    scale_p = 1
    drafting_area_centerX = draft_area_center_initX  # 前視圖中心
    drafting_area_centerY = draft_area_center_initY
    drafting_isometric_area_centerY = isometric_view_center_Y
    drafting_isometric_area_centerX = (drafting_view_max_X - drafting_view_min_Y) / 2

    while True:
        scale_tmp = scale_p  # temping original scale
        scale = 1 / scale_p  # proportion convert to ratio
        w_scale = width * scale  # width after scaling
        h_scale = height * scale  # height after scaling
        d_scale = depth * scale
        drafting_area_X_range = w_scale + d_scale * 2 + draft_X_clearence * 6
        drafting_area_Y_range = h_scale + draft_Y_clearence * 2
        drafting_area_extremum = [drafting_area_centerX - drafting_area_X_range / 2,  # X-min[0]
                                  drafting_area_centerX + drafting_area_X_range / 2,  # X-max[1]
                                  drafting_area_centerY - drafting_area_Y_range / 2,  # Y-min[2]
                                  drafting_area_centerY + drafting_area_Y_range / 2]  # Y-max[3]
        drafting_center_Y = drafting_view_max_Y / 2 + drafting_view_min_Y
        drafting_center_X = drafting_view_max_X / 2 + drafting_view_min_X
        drafting_root_X = drafting_view_min_X + 5  # 反迴圈框_X
        drafting_root_Y = drafting_view_min_Y + 5
        # ---------isometric view-------------
        drafting_isometric_X_range = w_scale * math.cos(math.radians(45)) * 3 + d_scale * math.cos(math.radians(45)) + scale * 3200 * math.cos(math.radians(45)) + 3500 * scale * math.cos(math.radians(45)) + draft_X_clearence * 2  # W多乘2次為增長邊界長度
        drafting_isometric_Y_range = (S + T) * scale * math.cos(math.radians(35.7)) + math.cos(math.radians(35.7)) + h_scale * math.cos(math.radians(35.7))  # h為增長邊界範圍
        drafting_isometric_area_extremum = [drafting_isometric_area_centerX - drafting_isometric_X_range / 2,  # X-min[0]
                                            drafting_isometric_area_centerX + drafting_isometric_X_range / 2,  # X-max[1]
                                            drafting_isometric_area_centerY - drafting_isometric_Y_range / 2,  # Y-min[2]
                                            drafting_isometric_area_centerY + drafting_isometric_Y_range / 2]  # Y-max[3]
        drafting_isometric_root_Y_min = drafting_center_Y + 5
        drafting_isometric_root_X_min = drafting_view_min_X + 5

        # ----------圖面比例位置判斷------------
        if scale_p % 2 != 0 or scale_p % 5 != 0 or scale_p % 10 != 0:
            scale_p += 1
        # ----------三視圖----------
        if drafting_area_extremum[2] > drafting_view_min_Y and drafting_area_extremum[3] > drafting_center_Y and drafting_area_extremum[2] < drafting_root_Y:  # 若Ymin再返迴圈和邊框之間同時Ymax同時大於圖框中心則比例縮小
            scale_p += 1
        elif drafting_area_extremum[2] < drafting_view_min_Y and drafting_area_extremum[3] > drafting_center_Y:  # 若Ymin和Ymax同時大於圖框則比例縮小
            scale_p += 1
        elif drafting_area_extremum[0] > drafting_view_min_X and drafting_area_extremum[1] > drafting_center_X and drafting_area_extremum[0] < drafting_root_X:  # 若Xmin再返迴圈和邊框之間同時Xmax同時大於圖框中心則比例縮小
            scale_p += 1
        elif drafting_area_extremum[0] < drafting_view_min_X and drafting_area_extremum[1] > drafting_center_X:  # 若Xmin和Xmax同時大於圖框則比例縮小
            scale_p += 1
        # ----------等角圖----------
        elif drafting_isometric_area_extremum[2] > drafting_center_Y and drafting_isometric_area_extremum[2] < drafting_isometric_root_Y_min and drafting_isometric_area_extremum[3] > drafting_view_max_Y:
            scale_p += 1
        elif drafting_isometric_area_extremum[2] < drafting_center_Y and drafting_isometric_area_extremum[3] > drafting_view_max_Y:
            scale_p += 1
        elif drafting_isometric_area_extremum[0] > drafting_view_min_X and drafting_isometric_area_extremum[0] < drafting_isometric_root_X_min and drafting_isometric_area_extremum[1] > drafting_view_max_X:  # 若Ymin和Ymax同時大於圖框則比例縮小
            scale_p += 1
        elif drafting_isometric_area_extremum[0] < drafting_view_min_X and drafting_isometric_area_extremum[1] > drafting_view_max_X:
            scale_p += 1
        else:
            break
    # ------------三視圖位置-------------
    while True:
        drafting_area_X_range = w_scale + d_scale * 2 + draft_X_clearence * 6
        drafting_area_Y_range = h_scale + draft_Y_clearence * 2
        drafting_area_extremum = [drafting_area_centerX - drafting_area_X_range / 2,  # X-min[0]
                                  drafting_area_centerX + drafting_area_X_range / 2,  # X-max[1]
                                  drafting_area_centerY - drafting_area_Y_range / 2 - h_scale*1 / 3,  # Y-min[2]
                                  drafting_area_centerY + drafting_area_Y_range / 2]  # Y-max[3]
        drafting_center_Y = drafting_view_max_Y / 2 + drafting_view_min_Y
        drafting_center_X = drafting_view_max_X / 2
        drafting_root_X = drafting_view_min_X + 5  # 反迴圈框_X
        drafting_root_Y = drafting_view_min_Y + 5
        # ------------Y方向-------------
        if drafting_area_extremum[2] <= drafting_view_min_Y:
            drafting_area_centerY += 1
        elif drafting_area_extremum[2] > drafting_root_Y:
            drafting_area_centerY -= 1
        # ------------X方向-------------
        elif drafting_area_extremum[0] <= drafting_view_min_X:
            drafting_area_centerX += 1
        elif drafting_area_extremum[0] <= drafting_root_X:
            drafting_area_centerX -= 1
        else:
            break
    # ---------isometric position-----------
    while True:
        drafting_isometric_X_range = w_scale * math.cos(math.radians(45)) * 3 + d_scale * math.cos(math.radians(45)) + scale * 4250 * math.cos(math.radians(45)) + 3500 * scale * math.cos(math.radians(45)) + draft_X_clearence * 2  # W多乘2次為增長邊界長度
        drafting_isometric_Y_range = (S + T) * scale * math.cos(math.radians(35.7)) + math.cos(math.radians(35.7)) + h_scale * math.cos(math.radians(35.7))  # h為增長邊界範圍
        drafting_isometric_area_extremum = [drafting_isometric_area_centerX - drafting_isometric_X_range / 2,  # X-min[0]
                                            drafting_isometric_area_centerX + drafting_isometric_X_range / 2,  # X-max[1]
                                            drafting_isometric_area_centerY - drafting_isometric_Y_range / 2,  # Y-min[2]
                                            drafting_isometric_area_centerY + drafting_isometric_Y_range / 2]  # Y-max[3]
        drafting_isometric_root_Y_min = drafting_center_Y + 5
        drafting_isometric_root_X_min = drafting_view_min_X + 5
        # ------------Y方向-------------
        if drafting_isometric_area_extremum[2] <= drafting_center_Y:  # 若爆炸圖圖面Y最小值小於中心軸則爆炸圖1圖面中心上移1
            drafting_isometric_area_centerY += 1
        elif drafting_isometric_area_extremum[2] >= drafting_isometric_root_Y_min:  # 若爆炸圖圖面Y最小值大於返回圈框則爆炸圖1圖面中心下移1
            drafting_isometric_area_centerY -= 1
        # ------------X方向-------------
        elif drafting_isometric_area_extremum[0] <= drafting_view_min_X:  # 若爆炸圖圖面X最小值小於中心軸則爆炸圖1圖面中心右移1
            drafting_isometric_area_centerX += 1
        elif drafting_isometric_area_extremum[0] > drafting_isometric_root_X_min:  # 若爆炸圖圖面X最小值小於中心軸則爆炸圖1圖面中心左移1
            drafting_isometric_area_centerX -= 1
        else:
            break
    # 三視圖位置
    drafting_Coordinate_Position = {'Front View': (drafting_area_centerX, drafting_area_centerY),
                                    'Left View': (drafting_area_centerX - w_scale / 2 - d_scale / 2, drafting_area_centerY),
                                    'Right View': (drafting_area_centerX + w_scale / 2 + d_scale / 2, drafting_area_centerY)}
    # 等角圖位置
    drafting_isometric_area_extremum_X = drafting_isometric_area_extremum[1] - drafting_isometric_area_extremum[0]
    drafting_isometric_area_extremum_Y = drafting_isometric_area_extremum[3] - drafting_isometric_area_extremum[2]
    drafting_isometric_Coordinate_Position = {'exploded_1': (drafting_center_X - (drafting_isometric_area_extremum[1] - drafting_isometric_area_extremum[0]) / 4, drafting_center_Y + (drafting_isometric_area_extremum[3] - drafting_isometric_area_extremum[2]) / 2),
                                              'exploded_2': (drafting_center_X + (drafting_isometric_area_extremum[1] - drafting_isometric_area_extremum[0]) / 4, drafting_center_Y + (drafting_isometric_area_extremum[3] - drafting_isometric_area_extremum[2]) / 2 - 35)}
    return drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale_p

def change_Drawing_scale(value):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingSheet.Scale = float(value)

def exploded_Drawing_1(X_coordinate, Y_coordinate, scale):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingViewGenerativeLinks = drawingView.GenerativeLinks
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item("SN1-110.CATProduct")
    product = productDocument.Product
    drawingViewGenerativeBehavior.Document = product
    drawingViewGenerativeBehavior.DefineIsometricView(-0.707107, 0.707107, 0, -0.408248, -0.408248, 0.816497)
    drawingView.X = X_coordinate
    drawingView.Y = Y_coordinate
    drawingView.Scale = 1 / scale
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    drawingViewGenerativeBehavior.Update()
    drawingView.Activate()
    selection = productDocument.Selection
    selection.Clear()
    drawingview1 = drawingViews.Item('Isometric view')
    drawingtexts1 = drawingview1.Texts
    drawingtext1 = drawingtexts1.Item(1)
    drawingtexts1 = drawingtext1.Parent
    selection.Add(drawingtext1)
    selection.Delete()
    selection.Clear()

def exploded_Drawing_2(X_coordinate, Y_coordinate, scale):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingViewGenerativeLinks = drawingView.GenerativeLinks
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item("SN1-110.CATProduct")
    product = productDocument.Product
    drawingViewGenerativeBehavior.Document = product
    drawingViewGenerativeBehavior.DefineIsometricView(0.707107, -0.707107, 0, 0.408248, 0.408248, 0.816497)
    drawingView.X = X_coordinate
    drawingView.Y = Y_coordinate
    drawingView.Scale = 1 / scale
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    drawingViewGenerativeBehavior.Update()
    drawingView.Activate()
    drawingtexts = drawingView.Texts
    drawingtext = drawingtexts.Item(1)
    drawingtexts = drawingtext.Parent
    selection = productDocument.Selection
    selection.Add(drawingtext)
    selection.Delete()
    selection.Clear()

def Front_View_Drawing(X_coordinate, Y_coordinate, scale, type):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item(type + ".CATProduct")
    product = productDocument.Product
    drawingViewGenerativeBehavior.Document = product
    drawingViewGenerativeBehavior.DefineFrontView(0, 1, 0, 0, 0, 1)
    drawingView.X = X_coordinate
    drawingView.Y = Y_coordinate
    drawingView.Scale = 1 / scale
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    drawingViewGenerativeBehavior.Update()
    drawingView.Activate()
    # drawingView = drawingViews.Add("Front view")
    # drawingviews1
    selection = productDocument.Selection
    selection.Clear()
    drawingview1 = drawingViews.Item('Front view')
    drawingtexts1 = drawingview1.Texts
    drawingtext1 = drawingtexts1.Item(1)
    drawingtexts1 = drawingtext1.Parent
    selection.Add(drawingtext1)
    selection.Delete()
    selection.Clear()

def Left_View_Drawing(X_coordinate, Y_coordinate, scale, type):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingViewGenerativeLinks = drawingView.GenerativeLinks
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item(type + ".CATProduct")
    product = productDocument.Product
    drawingViewGenerativeBehavior.Document = product
    drawingViewGenerativeBehavior.DefineFrontView(1, 0, 0, 0, 0, 1)
    drawingView.X = X_coordinate
    drawingView.Y = Y_coordinate
    drawingView.Scale = 1 / scale
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    drawingViewGenerativeBehavior.Update()
    drawingView.Activate()
    drawingtexts = drawingView.Texts
    drawingtext = drawingtexts.Item(1)
    drawingtexts = drawingtext.Parent
    selection = productDocument.Selection
    selection.Add(drawingtext)
    selection.Delete()
    selection.Clear()

def Right_View_Drawing(X_coordinate, Y_coordinate, scale, type):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingViewGenerativeLinks = drawingView.GenerativeLinks
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item(type + ".CATProduct")
    product = productDocument.Product
    drawingViewGenerativeBehavior.Document = product
    drawingViewGenerativeBehavior.DefineFrontView(-1, 0, 0, 0, 0, 1)
    drawingView.X = X_coordinate
    drawingView.Y = Y_coordinate
    drawingView.Scale = 1 / scale
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    drawingViewGenerativeBehavior.Update()
    drawingView.Activate()
    drawingtexts = drawingView.Texts
    drawingtext = drawingtexts.Item(1)
    drawingtexts = drawingtext.Parent
    selection = productDocument.Selection
    selection.Add(drawingtext)
    selection.Delete()
    selection.Clear()

def coordinate():
    catapp = win32.Dispatch('CATIA.Application')

    catDrwDoc = catapp.ActiveDocument
    catDrwSel = catDrwDoc.Selection
    catDrwSelLb = catDrwSel

def balloons(view, circle_position_1, circle_position_2, circle_position_3, point_position_1, point_position_2):
    catapp = win32.Dispatch("CATIA.Application")
    partdoc = catapp.ActiveDocument
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingviews = drawingsheet.Views
    drawingview = drawingviews.Item('Isometric view')
    drawingview.Activate()
    DrawTexts_ground_screw = drawingview.Texts
    DrawText = DrawTexts_ground_screw.Add(circle_position_1, circle_position_2, circle_position_3) #線段長度
    DrawText.SetFontName(0, 0, 'Arial Unicode MS (TrueType)')
    DrawText.SetFontSize(0, 0, 10)
    DrawLeader_nut_ground_screw = DrawText.Leaders.Add(point_position_1, point_position_2)  # 圓點位置
    DrawText.FrameType = 3  # 圓類型
    DrawLeader_nut_ground_screw.HeadSymbol = 20