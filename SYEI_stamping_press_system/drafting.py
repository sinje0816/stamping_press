import os
import win32com.client as win32
import datetime, time, math
import main_program as mprog
import parameter as par

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


def drafting_parameter_calculation(l, width, height, depth, T):  # 電子型錄WHD, 比例
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
        drafting_area_Y_range = h_scale + draft_Y_clearence * 6
        drafting_center_Y = (drafting_view_max_Y / 2 + drafting_view_min_Y)
        # ---------isometric view-------------
        drafting_isometric_X_range = w_scale * math.cos(math.radians(45)) * 3 + d_scale * math.cos(
            math.radians(45)) * 2 + scale * 4250 * math.cos(math.radians(45)) + 3500 * scale * math.cos(
            math.radians(45)) + draft_X_clearence * 2  # W多乘2次為增長邊界長度
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
    # -------------位置--------------
    # ------------三視圖-------------
    while True:
        drafting_area_X_range = w_scale + d_scale * 2 + draft_X_clearence * 6
        drafting_area_Y_range = h_scale + draft_Y_clearence * 2
        if l == 0:
            drafting_area_extremum = [drafting_area_centerX - drafting_area_X_range / 2,  # X-min[0]
                                      drafting_area_centerX + drafting_area_X_range / 2,  # X-max[1]
                                      drafting_area_centerY - drafting_area_Y_range / 2,  # Y-min[2]
                                      drafting_area_centerY + drafting_area_Y_range / 2]  # Y-max[3]
        else:
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
        drafting_isometric_Y_range = ((1700 + 500) * scale + T * scale + h_scale)  # h為增長邊界範圍(500為上下range)
        drafting_isometric_area_extremum = [drafting_isometric_area_centerX - drafting_isometric_X_range / 2,
                                            # X-min[0]
                                            drafting_isometric_area_centerX + drafting_isometric_X_range / 2,
                                            # X-max[1]
                                            drafting_isometric_area_centerY - drafting_isometric_Y_range / 2,
                                            # Y-min[2]
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
                                    'Left View': (
                                    drafting_area_centerX - w_scale / 2 - d_scale / 2, drafting_area_centerY),
                                    'Right View': (
                                    drafting_area_centerX + w_scale / 2 + d_scale / 2, drafting_area_centerY)}
    # 等角圖位置
    drafting_isometric_Coordinate_Position = {'exploded_1':
                                                  (drafting_area_centerX + 25 - (drafting_isometric_area_extremum[1] - drafting_isometric_area_extremum[0]) / 5,
                                                   drafting_view_max_Y - (drafting_isometric_area_extremum[3] - drafting_isometric_area_extremum[2]) * 2 / 3),
                                              'exploded_2':
                                                  (drafting_area_centerX + 50 + (drafting_isometric_area_extremum[1] - drafting_isometric_area_extremum[0]) / 5,
                                                   drafting_view_max_Y - (drafting_isometric_area_extremum[3] - drafting_isometric_area_extremum[2]) * 2 / 3 - 50)}
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
    selection = catapp.ActiveDocument.Selection
    selection.Clear()
    selection.Add(catapp.ActiveDocument.sheets.ActiveSheet.Views.ActiveView)
    selection.Search("Drafting.Text,sel")
    selection.Delete()

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
    selection = catapp.ActiveDocument.Selection
    selection.Clear()
    selection.Add(catapp.ActiveDocument.sheets.ActiveSheet.Views.ActiveView)
    selection.Search("Drafting.Text,sel")
    selection.Delete()


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
    selection = catapp.ActiveDocument.Selection
    selection.Clear()
    selection.Add(catapp.ActiveDocument.sheets.ActiveSheet.Views.ActiveView)
    selection.Search("Drafting.Text,sel")
    selection.Delete()


def Right_Top_View_Drawing(Type, X_coordinate, Y_coordinate, scale):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Add("AutomaticNaming")
    drawingView.SetViewName("Top view", '', '')  # 更改視圖名稱
    drawingViewGenerativeLinks = drawingView.GenerativeLinks
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    documents = catapp.Documents
    productDocument = documents.Item(str(Type) + ".CATProduct")
    product = productDocument.Product
    drawingViewGenerativeBehavior.Document = product
    drawingViewGenerativeBehavior.DefineFrontView(-1, 0, 0, 0, -1, 0)
    drawingView.X = X_coordinate
    drawingView.Y = Y_coordinate
    drawingView.Scale = 1 / scale
    drawingViewGenerativeBehavior = drawingView.GenerativeBehavior
    drawingViewGenerativeBehavior.Update()
    drawingView.Activate()
    selection = catapp.ActiveDocument.Selection
    selection.Clear()
    selection.Add(catapp.ActiveDocument.sheets.ActiveSheet.Views.ActiveView)
    selection.Search("Drafting.Text,sel")
    selection.Delete()

def coordinate():
    catapp = win32.Dispatch('CATIA.Application')
    catDrwDoc = catapp.ActiveDocument
    catDrwSel = catDrwDoc.Selection
    catDrwSelLb = catDrwSel

# 圈碼圖
def balloons(view, content, circle_position_1, circle_position_2, point_position_1, point_position_2):
    catapp = win32.Dispatch("CATIA.Application")
    partdoc = catapp.ActiveDocument
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingviews = drawingsheet.Views
    drawingview = drawingviews.Item(view)
    drawingview.Activate()
    # 圈碼內容及位置
    DrawText = drawingview.Texts.Add(content, circle_position_1, circle_position_2)  # (輸入內容, 線段長度X, 線段長度Y)
    DrawText.SetFontName(0, 0, 'Arial Unicode MS (TrueType)')
    DrawText.SetFontSize(0, 0, 10)  # 調整字體位置和大小(x, y, 字體大小)
    DrawText.FrameType = 3  # 圓類型
    # 引線點標註位置
    DrawLeader_DrawTexts_balloons = DrawText.Leaders.Add(point_position_1, point_position_2)  # 圓點位置
    DrawLeader_DrawTexts_balloons.AllAround = 0
    DrawLeader_DrawTexts_balloons.ModifyPoint(0, circle_position_1, 0)  # 選擇引線點 -> (0, 1, 2)並調整座標位置
    DrawLeader_DrawTexts_balloons.HeadSymbol = 20  # 引號標點類型

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
    # 更換線段type
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
def add_dimension_to_view(view_name, item_name, catDimDistance, x_value_1, y_value_1, x_value_2, y_value_2,
                          angle):  # 標註型式, 座標1(X, Y), 座標2(X, Y)
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


def symbol_of_weld(view, WeldingSymbol, lead_X, lead_Y, WeldingTail, parameter):  # 焊接符號
    catapp = win32.Dispatch("CATIA.Application")
    partdoc = catapp.ActiveDocument
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingviews = drawingsheet.Views
    drawingview = drawingviews.Item(view)
    MyWelding = drawingview.Weldings.Add(WeldingSymbol, lead_X, lead_Y)  # 引線標註點位置(引線符號, 座標X, 座標Y)
    # 基線位置
    # 基線計算位置時要注意需呈現為60度
    if lead_X >= 0:
        MyWelding.X = lead_X + 100 * math.cos(math.radians(60))
        MyWelding.angle = 0
    else:
        MyWelding.X = lead_X - 300
        MyWelding.angle = 0
    if lead_Y >= 0:
        MyWelding.Y = lead_Y + 300 * math.cos(math.radians(60)) / math.sin(math.radians(60))
    else:
        MyWelding.Y = lead_Y - 300 * math.cos(math.radians(60)) / math.sin(math.radians(60))
    MyWelding.WeldingTail = WeldingTail  # 是否開啟尾叉 0 or 1

    for i in range(1, 14):
        try:
            drawingWelding1 = MyWelding
            # drawingWelding1 = drawingWeldings1.Item('Welding Symbol.1')
            catWeldingFieldOne = i
            textRange1 = drawingWelding1.GetTextRange(catWeldingFieldOne)
            char1 = textRange1.Text
            textRange1.Text = parameter[i]
            TextProperties = drawingWelding1.TextProperties
            TextProperties.Update()
        except:
            pass


def background():
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingview = drawingsheet.Views.Item('Background View').Activate()


def drafting_welding_view_parameter_calculation(width, height, depth, S, Z, T):  # 電子型錄WHD, 比例
    scale_p = 1
    drafting_area_centerX = par.drafting_front_area_centerX  # 上:上視圖中心 & 下:前視圖中心
    drafting_down_area_centerY = par.drafting_down_area_centerY
    drafting_up_area_centerY = par.drafting_up_area_centerY

    while True:
        # 排除餘數不為 2, 5, 10 的數
        if scale_p % 2 != 0 or scale_p % 5 != 0 or scale_p % 10 != 0:
            scale_p += 1
        scale = 1 / scale_p  # proportion convert to ratio
        w_scale = width * scale  # width after scaling
        h_scale = height * scale  # height after scaling
        d_scale = depth * scale
        S_scale = S * scale
        Z_T_scale = (Z - T) * scale
        T_scale = T * scale
        # ------------下圖總範圍-------------
        drafting_down_area_X_range = w_scale * 2 * 1.5 + d_scale * 2 + par.draft_X_clearence * 15
        drafting_down_area_Y_range = h_scale + par.draft_Y_clearence * 2
        drafting_up_max_range_center_Y = (821 - par.drafting_view_min_X) * 2 / 3 + par.drafting_view_min_Y
        # ------------上圖總範圍-------------
        drafting_up_area_X_range = w_scale * 2 + d_scale * 2 + par.draft_X_clearence * 5
        drafting_up_area_Y_range = w_scale + par.draft_Y_clearence * 2
        # ----------圖面比例判斷------------
        # ----------下圖----------
        if drafting_down_area_X_range > par.drafting_view_max_X - par.drafting_view_min_X:
            scale_p += 1
        elif drafting_down_area_Y_range > drafting_up_max_range_center_Y - par.drafting_view_min_Y:
            scale_p += 1
        # ----------上圖----------
        elif drafting_up_area_X_range > par.drafting_view_max_X - par.drafting_view_min_X:
            scale_p += 1
        elif drafting_up_area_Y_range > par.drafting_view_max_Y - drafting_up_max_range_center_Y:
            scale_p += 1
        else:
            break
    # ------------下圖位置判斷-------------
    while True:
        drafting_area_X_left_range = w_scale * 1.5 + par.draft_X_clearence * 5
        drafting_area_X_right_range = w_scale + d_scale * 1.5 + par.draft_X_clearence * 10
        drafting_down_area_extremum = [drafting_area_centerX - drafting_area_X_left_range,  # X-min[0]
                                       drafting_area_centerX + drafting_area_X_right_range]  # X-max[1]
        drafting_root_X = par.drafting_view_min_X + 5  # 反迴圈框_X
        # ------------X方向-------------
        if drafting_down_area_extremum[0] <= par.drafting_view_min_X:
            drafting_area_centerX += 1
        elif drafting_down_area_extremum[0] <= drafting_root_X:
            drafting_area_centerX -= 1
        else:
            break
    # ------------上圖位置判斷--------------註:X位置共用
    while True:
        drafting_up_area_Y_range = w_scale + par.draft_Y_clearence * 2
        drafting_up_area_extremum = [drafting_up_area_centerY - drafting_up_area_Y_range / 2,  # Y-min[0]
                                     drafting_up_area_centerY + drafting_up_area_Y_range / 2]  # Y-max[1]
        drafting_up_root_Y_min = drafting_up_max_range_center_Y + 5
        # ------------Y方向-------------
        if drafting_up_area_extremum[0] >= drafting_up_root_Y_min:
            drafting_up_area_centerY -= 1
        elif drafting_up_area_extremum[0] <= drafting_up_max_range_center_Y:  # 若爆炸圖圖面Y最小值小於中心軸則爆炸圖1圖面中心上移1
            drafting_up_area_centerY += 1
        else:
            break
    # 下圖位置
    drafting_down_Coordinate_Position = {
        'Right View': (drafting_area_centerX - drafting_area_X_left_range, drafting_down_area_centerY),
        'Front View': (drafting_area_centerX, drafting_down_area_centerY),
        'Section view A-A': (drafting_area_centerX + drafting_area_X_left_range + (d_scale - w_scale) / 5, drafting_down_area_centerY),
        'Section view B-B': (drafting_area_centerX + drafting_area_X_left_range * 2 + (d_scale - w_scale) / 5, drafting_down_area_centerY),
        'Section view F-F': (drafting_area_centerX + drafting_area_X_left_range * 2 + (d_scale - w_scale) / 5 - w_scale / 3, drafting_down_area_centerY - h_scale / 3),
        'Section view G-G': (drafting_area_centerX + drafting_area_X_left_range * 2 + (d_scale - w_scale) / 5 + w_scale / 3, drafting_down_area_centerY - h_scale / 3)}
    # 上圖位置
    drafting_up_Coordinate_Position = {
        'Section view D-D': (drafting_area_centerX - drafting_area_X_left_range, drafting_up_area_centerY),
        'Top View': (drafting_area_centerX, drafting_up_area_centerY),
        'Section view C-C': (drafting_area_centerX + drafting_area_X_left_range + (d_scale - w_scale) / 5, drafting_up_area_centerY),
        'Section view E-E': (drafting_area_centerX + drafting_area_X_left_range * 2 + (d_scale - w_scale) * 1 / 2, drafting_up_area_centerY)}
    return drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale_p

# 剖面圖
def Section(view_name, Coordinate_X, Coordinate_Y, Scale, Coordinate, change_direction, new_create_view_name):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Sheet.1")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item(view_name)  # 圖框名稱
    drawingView1.Activate()
    drawingSheets1 = drawingDocument1.Sheets
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    drawingView2 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
    drawingView2.X = Coordinate_X
    drawingView2.Y = Coordinate_Y
    drawingView2.Scale = 1 / Scale
    drawingView2.Angle = 0
    SectionProfile =Coordinate
    drawingViewGenerativeBehavior2Variant = drawingViewGenerativeBehavior2
    drawingViewGenerativeBehavior2Variant.DefineSectionView(SectionProfile, "SectionView", "Offset", change_direction, drawingViewGenerativeBehavior1)
    drawingViewGenerativeLinks1 = drawingView2.GenerativeLinks
    drawingViewGenerativeLinks2 = drawingView1.GenerativeLinks
    drawingViewGenerativeLinks2.CopyLinksTo(drawingViewGenerativeLinks1)
    drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
    drawingViewGenerativeBehavior2.Update()
    drawingView2.ReferenceView = drawingView1
    drawingView2.AlignedWithReferenceView()
    # 選定新投影圖面刪除文字方塊
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Sheet.1")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item(new_create_view_name)  # 圖框名稱
    drawingView1.Activate()
    selection = catapp.ActiveDocument.Selection
    selection.Clear()
    selection.Add(catapp.ActiveDocument.sheets.ActiveSheet.Views.ActiveView)
    selection.Search("Drafting.Text,sel")
    selection.Delete()

def Define_Polygonal_Detail_View(view_name, Coordinate, Coordinate_X, Coordinate_Y):  # 留下框選範圍圖面
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Item(view_name)  # 圖框名稱
    drawingView.Activate()
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.ActiveSheet
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.ActiveView
    drawingViewGenerativeBehavior1 = drawingView.GenerativeBehavior
    SectionProfile = Coordinate  # 給定框選範圍點座標
    drawingViewGenerativeBehavior1Variant = drawingViewGenerativeBehavior1
    drawingViewGenerativeBehavior1Variant.DefinePolygonalDetailView(SectionProfile[0:], drawingViewGenerativeBehavior1)
    drawingViewGenerativeBehavior1 = drawingView.GenerativeBehavior
    drawingViewGenerativeBehavior1.ForceUpdate()
    # 重新設定圖面座標
    # try:
    #     drawingView.X = Coordinate_X
    #     drawingView.Y = Coordinate_Y
    # except:
    #     pass

def break_line(view_name, Coordinate, X_direction, Y_direction):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Item(view_name)  # 圖框名稱
    drawingView.Activate()
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    SectionProfile = Coordinate  # 選擇要框選省略之範圍
    drawingViewGenerativeBehavior1Variant = drawingViewGenerativeBehavior1
    drawingViewGenerativeBehavior1Variant.DefineBrokenView(SectionProfile[0:], X_direction, Y_direction)
    # drawingViewGenerativeBehavior1Variant.SetRealLineType(8, 0.2)
    # drawingViewGenerativeBehavior1Variant.visProperties1.SetRealWidth(2, 0.25)

def close_all_Generated_Shape():
    catapp = win32.Dispatch('CATIA.Application')
    selection = catapp.ActiveDocument.Selection
    selection.Clear()
    selection.Add(catapp.ActiveDocument.sheets.ActiveSheet.Views.ActiveView)
    selection.Search("Drafting.Generated Shape,all")
    selection.Delete()

def close_selection_text():
    catapp = win32.Dispatch('CATIA.Application')
    selection = catapp.ActiveDocument.Selection
    selection.Clear()
    selection.Add(catapp.ActiveDocument.sheets.ActiveSheet.Views.ActiveView)
    selection.Search("Drafting.text, sel")
    selection.Delete()

def move_view_Position(view_name, X_Position, Y_Position):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Item(view_name)  # 圖框名稱
    drawingView.Activate()
    drawingView.X = X_Position
    drawingView.Y = Y_Position

def Define_Polygonal_Cipping_View(view_name, Coordinate):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Item(view_name)  # 圖框名稱
    drawingView.Activate()
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.ActiveSheet
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.ActiveView
    drawingViewGenerativeBehavior1 = drawingView.GenerativeBehavior
    SectionProfile = Coordinate  # 給定框選範圍點座標
    drawingViewGenerativeBehavior1Variant = drawingViewGenerativeBehavior1
    drawingViewGenerativeBehavior1Variant.DefinePolygonalClippingView(SectionProfile[0:])
    drawingViewGenerativeBehavior1 = drawingView.GenerativeBehavior
    drawingViewGenerativeBehavior1.ForceUpdate()

# 指定type種類全部刪除
# 使用方法:"⌈輸入type⌋',all"
# 指定元素名稱並刪除
# 使用方法:"Name='⌈輸入名稱⌋',all"
def selection_Search_delete(view_name, selection_name):
    catapp = win32.Dispatch('CATIA.Application')
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews = drawingSheet.Views
    drawingView = drawingViews.Item(view_name)  # 圖框名稱
    drawingView.Activate()
    selection = catapp.ActiveDocument.Selection
    selection.Clear()
    selection.Add(catapp.ActiveDocument.sheets.ActiveSheet.Views.ActiveView)
    selection.Search(selection_name)
    selection.Delete()