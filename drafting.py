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
drafting_view_max_X = 1250
isometric_view_center_X = 270
isometric_view_center_Y = 600

def drafting_parameter_calculation(width, height, depth, T):  # 電子型錄WHD, 比例
    scale_p = 1
    drafting_area_centerX = draft_area_center_initX  # 前視圖中心
    drafting_area_centerY = draft_area_center_initY
    drafting_isometric_area_centerY = isometric_view_center_Y
    drafting_isometric_area_centerX = (drafting_view_max_X - drafting_view_min_Y) / 2

    while True:
        # 排除餘數不為 2, 5, 10 的數
        if scale_p % 2 != 0 or scale_p % 5 != 0 or scale_p % 10 != 0:
            scale_p += 1
        scale_tmp = scale_p  # temping original scale
        scale = 1 / scale_p  # proportion convert to ratio
        w_scale = width * scale  # width after scaling
        h_scale = height * scale  # height after scaling
        d_scale = depth * scale
        drafting_area_X_range = w_scale + d_scale * 2 + draft_X_clearence * 6
        drafting_area_Y_range = h_scale + draft_Y_clearence * 2
        drafting_center_Y = (drafting_view_max_Y / 2 + drafting_view_min_Y)
        # ---------isometric view-------------
        drafting_isometric_X_range = w_scale * math.cos(math.radians(45)) * 3 + d_scale * math.cos(math.radians(45)) * 2 + scale * 4250 * math.cos(math.radians(45)) + 3500 * scale * math.cos(math.radians(45)) + draft_X_clearence * 2  # W多乘2次為增長邊界長度
        drafting_isometric_Y_range = ((1700 + 500) * scale + T * scale + h_scale)  # h為增長邊界範圍(500為上下range)
        # ----------圖面比例位置判斷------------
        # ----------三視圖----------
        if drafting_area_X_range > drafting_view_max_X - drafting_view_min_X:
            scale_p += 1
        elif drafting_area_Y_range > drafting_center_Y - drafting_view_min_Y:
            scale_p += 1
        elif drafting_isometric_X_range > drafting_view_max_X - drafting_view_min_X:
            scale_p += 1
        elif drafting_isometric_Y_range > drafting_view_max_Y - drafting_center_Y:
            scale_p += 1
        else:
            break
    # ------------位置---------------
    # ------------三視圖-------------
    while True:
        drafting_area_X_range = w_scale + d_scale * 2 + draft_X_clearence * 6
        drafting_area_Y_range = h_scale + draft_Y_clearence * 2
        drafting_area_extremum = [drafting_area_centerX - drafting_area_X_range / 2,  # X-min[0]
                                  drafting_area_centerX + drafting_area_X_range / 2,  # X-max[1]
                                  drafting_area_centerY - drafting_area_Y_range / 2 - h_scale * 1 / 3,  # Y-min[2]
                                  drafting_area_centerY + drafting_area_Y_range / 2]  # Y-max[3]
        drafting_center_Y = drafting_view_max_Y / 2 + drafting_view_min_Y
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
        drafting_isometric_X_range = w_scale * math.cos(math.radians(45)) * 3 + d_scale * math.cos(
            math.radians(45)) * 2 + scale * 4250 * math.cos(math.radians(45)) + 3500 * scale * math.cos(
            math.radians(45)) + draft_X_clearence * 2  # W多乘2次為增長邊界長度
        drafting_isometric_Y_range = ((1700 + 500) * scale + T * scale + h_scale) # h為增長邊界範圍(500為上下range)
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
    drafting_isometric_Coordinate_Position = {'exploded_1': (drafting_area_centerX + 25 - (drafting_isometric_area_extremum[1] - drafting_isometric_area_extremum[0]) / 5, drafting_view_max_Y - (drafting_isometric_area_extremum[3] - drafting_isometric_area_extremum[2]) * 2 / 3),
                                              'exploded_2': (drafting_area_centerX + 50 + (drafting_isometric_area_extremum[1] - drafting_isometric_area_extremum[0]) / 5, drafting_view_max_Y - (drafting_isometric_area_extremum[3] - drafting_isometric_area_extremum[2]) * 2 / 3 - 50)}
    return drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale_p

def change_Drawing_scale(value):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingSheet.Scale = float(value)

def exploded_Drawing_1(Type, X_coordinate, Y_coordinate, scale):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingView.SetViewName("Isometric view1", '', '')  # 更改視圖名稱
    drawingViewGenerativeLinks = drawingView.GenerativeLinks
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item(str(Type) + ".CATProduct")
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
    drawingview1 = drawingViews.Item('Isometric view1')
    drawingtexts1 = drawingview1.Texts
    drawingtext1 = drawingtexts1.Item(1)
    drawingtexts1 = drawingtext1.Parent
    selection.Clear()
    selection.Add(drawingtext1)
    selection.Delete()
    selection.Clear()

def exploded_Drawing_2(Type, X_coordinate, Y_coordinate, scale):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingView.SetViewName("Isometric view2", '', '')  # 更改視圖名稱
    drawingViewGenerativeLinks = drawingView.GenerativeLinks
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item(str(Type) + ".CATProduct")
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

def Front_View_Drawing(Type, X_coordinate, Y_coordinate, scale):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingView.SetViewName("Front view", '', '')  # 更改視圖名稱
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item(str(Type) + ".CATProduct")
    product = productDocument.Product
    drawingViewGenerativeBehavior.Document = product
    drawingViewGenerativeBehavior.DefineFrontView(0, 1, 0, 0, 0, 1)
    drawingView.X = X_coordinate
    drawingView.Y = Y_coordinate
    drawingView.Scale = 1 / scale
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    drawingViewGenerativeBehavior.Update()
    drawingView.Activate()
    selection = productDocument.Selection
    drawingview1 = drawingViews.Item('Front view')
    drawingtexts1 = drawingview1.Texts
    drawingtext1 = drawingtexts1.Item(1)
    drawingtexts1 = drawingtext1.Parent
    selection.Add(drawingtext1)
    selection.Delete()
    selection.Clear()

def Left_View_Drawing(Type, X_coordinate, Y_coordinate, scale):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingView.SetViewName("Left view", '', '')  # 更改視圖名稱
    drawingViewGenerativeLinks = drawingView.GenerativeLinks
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item(str(Type) + ".CATProduct")
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

def Right_View_Drawing(Type, X_coordinate, Y_coordinate, scale):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingView.SetViewName("Right view", '', '')  # 更改視圖名稱
    drawingViewGenerativeLinks = drawingView.GenerativeLinks
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item(str(Type) + ".CATProduct")
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

# 圈碼圖
def balloons(view, circle_position_1, circle_position_2, circle_position_3, point_position_1, point_position_2, leader_line_length):
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
    DrawText = DrawTexts_balloons.Add(circle_position_1, circle_position_2, circle_position_3) #線段長度
    DrawText.SetFontName(0, 0, 'Arial Unicode MS (TrueType)')
    DrawText.SetFontSize(0, 0, 10)  # 調整字體位置和大小(x, y, 字體大小)
    DrawLeader_DrawTexts_balloons = DrawText.Leaders.Add(point_position_1, point_position_2)  # 圓點位置
    DrawText.FrameType = 3  # 圓類型
    # DrawText.DeactivateFrame(3)
    # DrawText.Deactivates = 1
    # DrawText.Blanking(0)
    DrawLeader_DrawTexts_balloons.AllAround = 0
    DrawLeader_DrawTexts_balloons.ModifyPoint(0, leader_line_length, 0)  # 選擇引線點 -> (0, 1, 2)並調整座標位置
    DrawLeader_DrawTexts_balloons.HeadSymbol = 20  # 引號標點類型
    # DrawText.StandardBehavior = 1
    # MyViewGenBehavior = MyView.GenerativeBehavior

# 爆炸圖中心線
def create_center_line(view_name, x_value_1, y_value_1, x_value_2, y_value_2):
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingviews = drawingsheet.Views
    drawingview = drawingviews.Item(view_name)
    drawingview.Activate()
    factory2d = drawingview.Factory2D
    dotted_line = factory2d.CreateLine(x_value_1, y_value_1, x_value_2, y_value_2)
    dotted_line.Name = 'chain_line'
    object = catapp.ActiveDocument.Selection
    object.Search('Name=chain_line,all')
    vis = catapp.ActiveDocument.Selection.VisProperties
    vis.SetRealLineType(2, 0.3)
    vis.SetRealWidth(1, 0.13)

def close_broken_line_block_diagram(view_name):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Sheet.1")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item(view_name)
    drawingView1.FrameVisualization = False

# 尺寸標註
def add_dimension_to_view(view_name, item_name, catDimDistance, x_value_1, y_value_1, x_value_2, y_value_2, angle):  # 標註型式, 座標1(X, Y), 座標2(X, Y)
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingviews = drawingsheet.Views
    drawingview = drawingviews.Item(view_name)  # 圖框名稱
    drawingview.Activate()  # 選擇圖框
    factory2d = drawingview.Factory2D
    Point1 = factory2d.CreatePoint(x_value_1, y_value_1)  # 建點
    Point2 = factory2d.CreatePoint(x_value_2, y_value_2)
    iType = catDimDistance  # 標註類型
    geoelem = [Point1, Point2]  # 選擇標註點
    geocoordElem = [x_value_1, y_value_1, x_value_2, y_value_2]  # 座標
    dim = drawingview.Dimensions.Add2(iType, geoelem, geocoordElem, None, angle)  # (標註類型, 選擇標註點, 座標, 參考元素(可空著), 角度)
    dim.Name = item_name
    drawingdim = drawingview.Dimensions.Item(dim.Name)
    if x_value_1 > x_value_2:
        if x_value_1 >= 0:
            x_value = x_value_1 + 300
        else:
            x_value = x_value_1 - 300
    else:
        if x_value_2 >= 0:
            x_value = x_value_2 + 300
        else:
            x_value = x_value_2 - 300
    if y_value_1 > y_value_2:
        if y_value_1 >= 0:
            y_value = y_value_1 + 300
        else:
            y_value = y_value_1 - 300
    else:
        if y_value_2 >= 0:
            y_value = y_value_2 + 300
        else:
            y_value = y_value_2 - 300
    drawingdim.MoveValue(x_value, y_value, 2, 0)  # 標註偏移(X, Y, 選擇偏移部位, 角度尺寸)