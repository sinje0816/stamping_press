import os
import win32com.client as win32
import datetime, time, math
import main_program as mprog

A = [720, 830, 890, 940, 1050, 1160, 1300, 1480, 1560, 1760]
A_15 = [1080, 1245, 1335, 1410, 1575, 1740, 1950, 2220, 2340, 2640]
B = [1058, 1125, 1210, 1315, 1480, 1680, 1985, 2113, 2400, 2700]
B_15 = [1587, 1688, 1815, 1973, 2220, 2520, 2978, 3170, 3600, 4050]
H = [2060, 2185, 2290, 2540, 2755, 2990, 3270, 3725, 4005, 4285]
R = [388, 486, 516, 544, 614, 670, 730, 900, 970, 1040]
R_15 = [582, 729, 774, 816, 921, 1005, 1095, 1350, 1455, 1560]
E = [700, 780, 840, 900, 1050, 1150, 1250, 1400, 1500, 1600]
E_15 = [1051, 1170, 1260, 1350, 1575, 1725, 1875, 2100, 2250, 2400]
D_DH = [250, 280, 330, 350, 380, 430, 490, 550, 580, 610]
D_S = [80, 90, 110, 130, 150, 180, 200, 220, 250, 280]
D_H = [50, 60, 70, 80, 100, 110, 130, 150, 180, 210]
D_P = [35, 40, 45, 50, 60, 70, 80, 90, 100, 110]
DH_S = [230, 250, 270, 300, 330, 350, 400, 450, 450, 500]
DH_H = [200, 220, 240, 270, 300, 320, 360, 400, 400, 450]
DH_P = [200, 220, 240, 270, 300, 320, 360, 400, 400, 450]
S = [983, 1068, 1158, 1285, 1445, 1630, 1809, 2067, 2262, 2357]
H_Z = [1260, 1385, 1490, 1640, 1855, 2086.933, 2370, 2725, 3005, 3285]
O = [1045, 1075, 1125, 1145, 1175, 1225, 1285, 1345, 1375, 1405]
hole_type = [0, 1, 2]
P = [330, 380, 430, 480, 560, 650, 720, 860, 960, 1060]
P_15 = [495, 570, 645, 721, 840, 975, 1080, 1290, 1440, 1590]
Q = [250, 300, 350, 400, 460, 520, 580, 650, 720, 790]
Q_15 = [375, 450, 525, 600, 690, 780, 870, 975, 1080, 1185]
T = [85, 100, 115, 130, 140, 155, 165, 180, 180, 200]
Z = [800, 800, 800, 900, 900, 900, 900, 1000, 1000, 1100]
F = [320, 400, 440, 520, 600, 680, 760, 840, 900, 960]
FRAME20_H = [280, 355, 420, 437.5, 550, 680, 825, 990, 1120, 1170]
FRAME2_lower_depth = [166.016, 246.016, 286.016, 366.016, 446.016, 526.016, 606.016, 686.016, 746.016, 806.016]
FRAME2_lower_depth_15 = [331.016, 451.016, 511.016, 631.016, 751.016, 871.016, 991.016, 1111.016, 1201.016, 1291.016]
FRAME1_lower_high = [1330, 1335, 1340, 1445, 1455, 1460, 1470, 1575, 1595, 1695]
FRAME20_FRAME2_YZ = [805, 805, 979, 979, 979, 979, 979, 979, 979, 979]
BALANCER1_XZ = [204, 253, 268, 282, 317, 345, 375, 460, 575, 690]  # (R+180mm)/2+80mm
FRAME_10_H = [558, 620.5, 673, 798, 905.5, 1023, 1163, 1320.5, 1530.5, 1740.5]
FRAME_32_XY = [0, 0, 0, 1722, 1819.5, 1922, 2052, 2294.5, 2504.5, 2794.5]

draft_Surface_Border = 20
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
        # drafting_isometric_root_Y_max = drafting_view_max_Y - 5
        # drafting_isometric_root_X_max = drafting_view_max_X - 5
        drafting_isometric_root_Y_min = drafting_center_Y + 5
        drafting_isometric_root_X_min = drafting_view_min_X + 5
        # drafting_isometric_root_Y_max = drafting_view_max_Y - 5
        # drafting_isometric_root_X_max = drafting_view_max_X - 5

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
            print(scale_p)
            break
    # ------------三視圖位置-------------
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
                                  drafting_area_centerY - drafting_area_Y_range / 2 - h_scale*1 / 3,  # Y-min[2]
                                  drafting_area_centerY + drafting_area_Y_range / 2]  # Y-max[3]
        drafting_center_Y = drafting_view_max_Y / 2 + drafting_view_min_Y
        drafting_center_X = drafting_view_max_X / 2 + drafting_view_min_X
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
        scale_tmp = scale_p  # temping original scale
        scale = 1 / scale_p  # proportion convert to ratio
        w_scale = width * scale  # width after scaling
        h_scale = height * scale  # height after scaling
        d_scale = depth * scale
        drafting_isometric_X_range = w_scale * math.cos(math.radians(45)) * 3 + d_scale * math.cos(math.radians(45)) + scale * 3200 * math.cos(math.radians(45)) + 3500 * scale * math.cos(math.radians(45)) + draft_X_clearence * 2  # W多乘2次為增長邊界長度
        drafting_isometric_Y_range = (S + T) * scale * math.cos(math.radians(35.7)) + math.cos(math.radians(35.7)) + h_scale * math.cos(math.radians(35.7))  # h為增長邊界範圍
        drafting_isometric_area_extremum = [drafting_isometric_area_centerX - drafting_isometric_X_range / 2,  # X-min[0]
                                            drafting_isometric_area_centerX + drafting_isometric_X_range / 2,  # X-max[1]
                                            drafting_isometric_area_centerY - drafting_isometric_Y_range / 2,  # Y-min[2]
                                            drafting_isometric_area_centerY + drafting_isometric_Y_range / 2]  # Y-max[3]
        drafting_isometric_root_Y_min = drafting_center_Y + 5
        drafting_isometric_root_X_min = drafting_view_min_X + 5
        # drafting_isometric_root_Y_max = drafting_view_max_Y - 5
        # drafting_isometric_root_X_max = drafting_view_max_X - 5
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
    drafting_isometric_Coordinate_Position = {'exploded_1': (drafting_area_centerX - 3500 * scale * math.cos(math.radians(45)), (drafting_view_max_Y - drafting_center_Y) / 2 + drafting_center_Y - 50),
                                              'exploded_2': (drafting_area_centerX + 3500 * scale * math.cos(math.radians(45)), (drafting_view_max_Y - drafting_center_Y) / 2 + drafting_center_Y - 85)}
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
    drawingViewGenerativeLinks = drawingView.GenerativeLinks
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

# mprog.change_Drawing_scale(1 / scale)
# # drafting_Coordinate_Position = drafting_parameter_calculation(A[5], B[5], H[5])

# exploded_drafting_down_Coordinate = {'Front View': (0, 1, 0, 0, 0, 1),
#                                      'Left View': (1, 0, 0, 0, 0, 1),
#                                      'Right View': (-1, 0, 0, 0, 0, 1)}
#
# mprog.Front_View_Drawing(drafting_Coordinate_Position['Front View'][0], drafting_Coordinate_Position['Front View'][1], scale)
# mprog.Left_View_Drawing(drafting_Coordinate_Position['Left View'][0], drafting_Coordinate_Position['Left View'][1], scale)
# mprog.Right_View_Drawing(drafting_Coordinate_Position['Right View'][0], drafting_Coordinate_Position['Right View'][1], scale)