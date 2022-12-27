from PyQt5 import QtCore, QtGui, QtWidgets
import sys, datetime, os
from GUI import Ui_Dialog
import main_program as mprog
import win32com.client as win32
import N_Test as N
import drafting_tesst2 as DT


# 電子型錄規格
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

FRAME44_height = [588 , 665.5 , 733 , 773 , 890.5 , 1023 , 1173 , 1385.5 , 1455.5 , 1445.5]
FRAME_41_depth = [590 , 590 , 590 , 590 , 490 , 590 , 590 , 590 , 590 , 590]
FRAME_7_width = [176 , 182 , 197 , 208 , 228 , 255 , 295 , 299 , 305 , 305]
FRAME_7_15_width = [200.3 , 209.3 , 231.8 , 248.3 , 278.3 , 318.8 , 378.8 , 384.8 , 393.8 , 393.8]
FRAME_11_height = [1122 , 1184.5 , 1237 , 1362 , 1469.5 , 1587 , 1727 , 1884.5 , 2094.5 , 2304.5]
FRAME_11_width = [600 , 680 , 720 , 800 , 880 , 960 , 1040 , 1120 , 1180 , 1240]
FRAME_11_15_width = [765 , 885 , 945 , 1065 , 1185 , 1305 , 1425 , 1545 , 1635 , 1725]
FRAME_8_width = [164 , 170 , 185 , 196 , 216 , 243 , 283 , 288 , 293 , 358]
FRAME_8_15_width = [246 , 255 , 278 , 294 , 324 , 365 , 425 , 432 , 440 , 537]
FRAME_13_depth = [59.016 , 139.016 , 179.016 , 259.016 , 339.016 , 419.016 , 499.016 , 579.016 , 639.016 , 699.016]



projection = {"front view":(0, 1 , 0, 0 , 0, 1) , "Rear view":(0, -1, 0, 0 , 0, 1) , "top view":(0, 1 , 0, -1 , 0 , 0)
    , "bottom view":(0, 1, 0, 1, 0, 0), "Left view":(1, 0, 0, 0, 0, 1) , "right view":(-1, 0, 0, 0, 0, 1)
              , 'top view(left horizontal)':(-1 , 0 , 0 , 0 , -1 , 0)  , "top view(Y inverse)":(0 , -1 , 0 , -1 , 0 , 0) ,
              'bottom view(right horizontal)':(-1 , 0 , 0 , 0 , 1 , 0) , 'right view(left horizontal)':(0 , 0 , -1 , -1 , 0 , 0),
              'Rear view(left horizontal)':(0 , 0 , -1 , 0 , -1 , 0) , 'top view(right horizontal)':(1 , 0 , 0 , 0 , 1 , 0) ,
              'top view(180 degree)':(0 , -1 , 0 , 1 , 0 , 0) , 'bottom view(left horizontal)':(1 , 0 , 0 , 0 , -1 , 0)}
U = ["front view" , "Rear view" , "top view" , "bottom view" , "Left view" , "right view" , 'top view(left horizontal)' ,
     "top view(Y inverse)" , 'bottom view(right horizontal)' , 'right view(left horizontal)','Rear view(left horizontal)',
     'top view(right horizontal)' , 'top view(180 degree)' , 'bottom view(left horizontal)']
#圖框範圍
circle_gap = 400
ALL_range = []
gap = 5
drafting_min_Y = 44
drafting_max_Y = 810
drafting_min_X = 20
drafting_max_X = 1169
#第一虛擬方框位置
box_1_Xmax = 372
box_1_Ymax = 810
box_1_Xmin = 25
box_1_Ymin = 25
box_width_gap = 80+2*gap#虛擬方框一的寬度間隙
box_heigth_gap = 150+2*gap#虛擬方框一的高度間隙

def test(i, l):
    scale, box_1_center , box_1_range = DT.scale_Adjustment(5, 1)
    scale = 1 / 4
    ALL_range , scale = DT.drafting_parameter_calculation(5, 1 , scale , box_1_range)#ALL_range = [方寬編號[Xmax , Xmin , Ymax , Ymin]]
    # projection_file_name_list = ['FRAME1' , 'FRAME2' ,'FRAME30', 'FRAME44', 'FRAME41', 'FRAME34', 'FRAME9', 'FRAME45', 'FRAME43', 'FRAME7',
    #                              'MAIN_GEAR2', 'FRAME32' , 'FRAME47' , 'FRAME23' , 'FRAME31', 'FRAME24', 'FRAME38' , 'FRAME11', 'FRAME39', 'FRAME17',
    #                              'FRAME3', 'FRAME13' , 'FRAME14', 'FRAME22',
    #                              'FRAME37', 'FRAME8', 'FRAME29', 'FRAME20', 'FRAME33' , 'FRAME48' , 'FRAME49']
    projection_file_name_list = ['FRAME1' , 'FRAME2' , 'FRAME30' , 'FRAME44' , 'FRAME41' , 'FRAME34' , 'FRAME9' , 'FRAME45' , 'FRAME43' , 'FRAME7',
                                 'MAIN_GEAR2' , 'FRAME32']
    part_circle_position = {
        '1': ['1' , -B[i] / 2 , H[i] / 2 + circle_gap],
        '2': ['2' , -B[i] / 2 , H[i] / 2 + circle_gap],
        '3': ['3' , -(R[i]+180) / 2 , FRAME44_height[i] / 2 + circle_gap],
        '4': ['4' , -(R[i]+180) / 2 , FRAME44_height[i] / 2 + circle_gap],
        '5': ['5' , -(R[i]+180) / 2 , FRAME_41_depth[i] / 2 + circle_gap],
        '6': ['6' , -182 / 2 , 80 / 2 + circle_gap],
        '7': ['7' , -(R[i]+180) / 2 , 50 / 2 +circle_gap],
        '8': ['8' , -R[i] / 2 , 476 / 2 + circle_gap],
        '9': ['8-1' , -R[i] / 2 , 150 / 2 + circle_gap],
        '10': ['10' , -FRAME_7_width[i] / 2 , 300 / 2 + circle_gap],
        '11': ['11' , -125 / 2 , 270 / 2 + circle_gap],
        '12': ['12' , -R[i] / 2 , 429 / 2 + circle_gap],
    }
    for x in projection_file_name_list:
        mprog.import_part("C:\\Users\\USER\\Desktop\\stamping_press", x)
        if l == 1:
            if x == 'FRAME1':
                p = 4#左視圖投影
                p2 = 5#右視圖投影
                X1 = (ALL_range[0][0]+ALL_range[0][1])/4 + 10#第一個投影X中心+尺寸標註預留量
                X2 = (ALL_range[0][0]+ALL_range[0][1])*3/4#第二個投影X中心(X中心為將X範圍分成兩塊之中心)
                Y1 = (ALL_range[0][2]+ALL_range[0][3])*3/4#第一個投影Y中心
                Y2 = (ALL_range[0][2]+ALL_range[0][3])*3/4#第二個投影Y中心(Y中心為將Y範圍分成兩塊之上半部中心)
                N.Material_diagram_projection(p, X1, Y1 , x , scale)
                N.Material_diagram_projection(p2, X2, Y2 , x, scale)
                N.Material_diagram_balloons(x , part_circle_position['1'][0] , part_circle_position['1'][1] , part_circle_position['1'][2])
            elif x == 'FRAME2':
                p = 4#左視圖投影
                p2 = 5#右視圖投影
                X1 = (ALL_range[0][0]+ALL_range[0][1])/4 + 10#第一個投影X中心+尺寸標註預留量
                X2 = (ALL_range[0][0]+ALL_range[0][1])*3/4#第二個投影X中心(X中心為將X範圍分成兩塊之中心)
                Y1 = (ALL_range[0][2]+ALL_range[0][3])/4#第一個投影Y中心
                Y2 = (ALL_range[0][2]+ALL_range[0][3])/4#第二個投影Y中心(Y中心為將Y範圍分成兩塊之上半部中心)
                N.Material_diagram_projection(p, X1, Y1 , x, scale)
                N.Material_diagram_projection(p2, X2, Y2 , x, scale)
                N.Material_diagram_balloons(x, part_circle_position['2'][0], part_circle_position['2'][1],
                                            part_circle_position['2'][2])
            elif x == 'FRAME30':
                p = 0#前視圖投影
                p2 = 5#右視圖投影
                p3 = 3#下視圖投影
                X1 = ALL_range[1][1] + (R[i]+180)*(scale)/2 + gap#BOX_2_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[1][0] - (40)*(scale)/2 - gap#BOX_2_Xmax-FRAME30厚度一半-間隙
                X3 = ALL_range[1][1] + (R[i]+180)*(scale)/2 + gap#BOX_2_Xmin+前視圖寬度一半+間隙
                Y1 =ALL_range[1][2] - (FRAME44_height[i])*(scale)/2 - gap#BOX_2_Ymax-前視圖高度-間隙(取完整高度是為了使零件圖與右壁板高度相近)
                Y2 =ALL_range[1][2] - (FRAME44_height[i])*(scale)/2 - gap#BOX_2_Ymax-前視圖高度-間隙(取完整高度是為了使零件圖與右壁板高度相近)
                Y3 =ALL_range[1][3] + (40)*(scale)/2 + gap#BOX_2_Ymin+FRAME30厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1 , x, scale)
                N.Material_diagram_projection(p2, X2, Y2 , x, scale)
                N.Material_diagram_projection(p3, X3, Y3 , x, scale)
                N.Material_diagram_balloons(x, part_circle_position['3'][0], part_circle_position['3'][1],
                                            part_circle_position['3'][2])
            elif x == 'FRAME44':
                p = 0#前視圖投影
                p2 = 3#下視圖投影
                X1 = ALL_range[2][1] + (R[i]+180)*(scale)/2 + gap#box_3_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[2][1] + (R[i]+180)*(scale)/2 + gap#box_3_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[2][2] - (FRAME44_height[i])*(scale)/2 - gap#box_3_Ymax+FRAME44高度一半-間隙
                Y2 = ALL_range[2][3] + 40*(scale)/2 + gap#box_3_Ymin+FRAME44厚度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_balloons(x, part_circle_position['4'][0], part_circle_position['4'][1],
                                            part_circle_position['4'][2])

            elif x == 'FRAME41':
                p = 6#上視圖(左轉向)投影
                p2 = 5#右視圖投影
                X1 = ALL_range[3][1] + (R[i] + 180) * (scale) / 2 + gap#box_4_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[3][1] + (R[i] + 180) * (scale) / 2 + gap#box_4_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[3][2] - (FRAME_41_depth[i]) * (scale) / 2 - gap#box_3_Ymax+FRAME41高度一半-間隙
                Y2 = ALL_range[3][3] + 22 * (scale) / 2 + gap#box_3_Yin+FRAME41厚度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_balloons(x, part_circle_position['5'][0], part_circle_position['5'][1],
                                            part_circle_position['5'][2])
            elif x == 'FRAME34':
                p2 = 1#後視圖投影
                p = 5#右視圖投影
                p3 = 7#下視圖翻轉180度投影
                X2 = ALL_range[4][0] - 182 * scale / 2 - gap#box_5_Xmax-FRAME34寬度一半-間隙
                X1 = ALL_range[4][1] + 40 * scale / 2 + gap#box_5_Xmin+FRAME34厚度一半+間隙
                X3 = ALL_range[4][0] - 182 * scale / 2 - gap#box_5_Xmax-FRAME34寬度一半-間隙
                Y2 = ALL_range[4][2] - 80 * scale / 2 - gap#box_5_Ymax-FRAME34深度一半-間隙
                Y1 = ALL_range[4][2] - 80 * scale / 2 - gap#box_5_Ymax-FRAME34深度一半-間隙
                Y3 = ALL_range[4][3] + 40 * scale / 2 + gap#box_5_Ymax+FRAME34厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_projection(p3, X3, Y3, x, scale)
                N.Material_diagram_balloons(x, part_circle_position['6'][0], part_circle_position['6'][1],
                                            part_circle_position['6'][2])
            elif x == 'FRAME9':
                p = 2#上視圖投影
                p2 = 0#前視圖投影
                X1 = ALL_range[5][1] + (R[i] + 180) * scale / 2 + gap#box_6_Xmin+FRAME9寬度一半+間隙
                X2 = ALL_range[5][1] + (R[i] + 180) * scale / 2 + gap #box_6_Xmin+FRAME9寬度一半+間隙
                Y1 = ALL_range[5][2] - 50 * scale / 2 - gap#box_6_Ymax-FRAME9厚度一半-間隙
                Y2 = ALL_range[5][3] + (Z[i] - T[i] - 40) * scale / 2 + gap#box_6_Ymin+FRAME9高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_balloons(x, part_circle_position['7'][0], part_circle_position['7'][1],
                                            part_circle_position['7'][2])
            elif x =='FRAME45':
                p = 2#上視圖投影
                p2 = 9#右側視圖(左旋轉)投影
                X1 = ALL_range[6][1] + R[i] * scale / 2 + gap#box_7_Xmin+FRAME45寬度一半+間隙
                X2 = ALL_range[6][0] - 19 * scale / 2 - gap#box_7_Xmax+FRAME45厚度一半-間隙
                Y1 = ALL_range[6][2] - 476 * scale / 2 - gap#box_7_Ymax-FRAME45深度一半-間隙
                Y2 = ALL_range[6][2] - 476 * scale / 2 - gap  # box_7_Ymax-FRAME45深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_balloons(x, part_circle_position['8'][0], part_circle_position['8'][1],
                                            part_circle_position['8'][2])
            elif x == 'FRAME43':
                p = 0#前視圖投影
                X1 = ALL_range[7][1] + R[i] * scale / 2 + gap#box_8_Xmin+FRAME43寬度一半+間隙
                Y1 = ALL_range[7][2] - 150 * scale / 2 - gap#box_8_Ymax-FRAME45深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_balloons(x, part_circle_position['9'][0], part_circle_position['9'][1],
                                            part_circle_position['9'][2])
            elif x == 'FRAME7':
                p = 8#下視圖(右旋)投影
                X1 = ALL_range[8][1] + FRAME_7_width[i] * scale / 2 + gap#box_9_Xmin+FRAME7寬度一半+間隙
                Y1 = ALL_range[8][2] - 300 * scale / 2 - gap#box_9_Ymax-FRAME7深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_balloons(x, part_circle_position['10'][0], part_circle_position['10'][1],
                                            part_circle_position['10'][2])
            elif x == 'MAIN_GEAR2':
                p = 5#右視圖投影
                X1 = ALL_range[9][1] + 125 * scale / 2 + gap#box_10_Xmin+MAIN_GEAR2深度一半+間隙
                Y1 = ALL_range[9][2] - 270 * scale / 2 - gap#box_10_Ymax-MAIN_GEAR2高度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_balloons(x, part_circle_position['11'][0], part_circle_position['11'][1],
                                            part_circle_position['11'][2])
            elif x == 'FRAME32':
                p = 6#上視圖(左旋)投影
                p2 = 10#背視圖(左旋)投影
                X1 = ALL_range[10][1] + R[i] * scale / 2 + gap#box_11_Xmin+FRAME32寬度一半+間隙
                X2 = ALL_range[10][0] - 19 * scale / 2 - gap  # box_11_Xmax-FRAME32厚度一半-間隙
                Y1 = ALL_range[10][2] - 429 * scale / 2 - gap#box_11_Ymax-FRAME深度一半-間隙
                Y2 = ALL_range[10][2] - 429 * scale / 2 - gap  # box_11_Ymax-FRAME深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_balloons(x, part_circle_position['12'][0], part_circle_position['12'][1],
                                            part_circle_position['12'][2])
            elif x == 'FRAME47':
                p = 6#上視圖(左旋)投影
                p2 = 1#背視圖投影
                X1 = ALL_range[11][1] + 50 * scale / 2 + gap#box_12_Xmin+FRAME47寬度一半+間隙
                X2 = ALL_range[11][0] - 19 * scale / 2 - gap#box_12_Xmax-FRAME47厚度一半-間隙
                Y1 = ALL_range[11][2] - 74 * scale / 2 - gap#box_12_Ymax-FRAME47深度一半-間隙
                Y2 = ALL_range[11][2] - 74 * scale / 2 - gap  # box_12_Ymax-FRAME47深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x== 'FRAME23':
                p = 6#上視圖(左旋)投影
                p2 = 5#右視圖投影
                X1 = ALL_range[12][1] + 99.35 * scale / 2 + gap#box_13_Xmin+FRAME23深度一半+間隙
                X2 = ALL_range[12][1] + 99.35 * scale / 2 + gap  # box_13_Xmin+FRAME23深度一半+間隙
                Y1 = ALL_range[12][2] - 35 * scale / 2 - gap#box_13_Ymax-FRAME23寬度一半-間隙
                Y2 = ALL_range[12][3] + 150 * scale / 2 + gap#box_13_Ymin+FRAME23高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x== 'FRAME31':
                p = 0#前視圖投影
                p2 = 5#右視圖投影
                p3 = 3#下視圖投影
                p4 = 0#前視圖投影
                p5 = 3#下視圖投影
                X1 = ALL_range[13][1] + 40 * scale / 2 + gap#box_14_Xmin+FRAME31寬度一半+間隙
                X2 = (ALL_range[13][1]+ALL_range[13][0])/2#box_14_X中心
                X3 = ALL_range[13][1] + 40 * scale / 2 + gap#box_14_Xmin+FRAME31寬度一半+間隙
                X4 = ALL_range[13][0] - 40 * scale / 2 - gap#box_14_Xmax-FRAME31寬度一半-間隙
                X5 = ALL_range[13][0] - 40 * scale / 2 - gap#box_14_Xmax-FRAME31寬度一半-間隙
                Y1 = ALL_range[13][2] - 150 * scale / 2 - gap#box_14_Ymax-FRAME31高度一半-間隙
                Y2 = ALL_range[13][2] - 150 * scale / 2 - gap#box_14_Ymax-FRAME31高度一半-間隙
                Y3 = ALL_range[13][3] + 45 * scale / 2 + gap#box_14_Ymin+FRAME31厚度一半+間隙
                Y4 = ALL_range[13][2] - 150 * scale / 2 - gap#box_14_Ymax-FRAME31高度一半-間隙
                Y5 = ALL_range[13][3] + 45 * scale / 2 + gap#box_14_Ymin+FRAME31厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_projection(p3, X3, Y3, x, scale)
                N.Material_diagram_projection(p4, X4, Y4, x, scale)
                N.Material_diagram_projection(p5, X5, Y5, x, scale)
            elif x == 'FRAME24':
                p = 6#上視圖(左旋)投影
                p2 = 1#背視圖投影
                p3 = 5#右視圖投影
                X1 = ALL_range[14][1] + 99.35 * scale / 2 + gap#box_15Xmin+FRAME24深度一半+間隙
                X2 = ALL_range[14][0] - 35 * scale / 2 + gap#box_15Xmax-FRAME24寬度一半-間隙
                X3 = ALL_range[14][1] + 99.35 * scale / 2 + gap#box_15Xmin+FRAME24深度一半+間隙
                Y1 = ALL_range[14][2] - 35 * scale / 2 - gap#box_15Ymax-FRAME24寬度一半-間隙
                Y2 = ALL_range[14][3] + 150 * scale / 2 + gap#box_15Ymin+FRAME24高度一半+間隙
                Y3 = ALL_range[14][3] + 150 * scale / 2 + gap  # box_15Ymin+FRAME24高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_projection(p3, X3, Y3, x, scale)
            elif x == 'FRAME38':
                p = 11#上視圖(右旋)投影
                X1 = ALL_range[15][1] + 290 * scale / 2 + gap#box_16_Xmin+FRAME38寬度一半+間隙
                Y1 = ALL_range[15][2] - 145 * scale / 2 - gap#box_16_Ymax-FRAME38深度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
            elif x == 'FRAME11':
                p = 5#右視圖投影
                p2 = 6#上視圖(左旋)投影
                X1 = ALL_range[16][1] + FRAME_11_width[i] * scale / 2 + gap#box_17_Xmin+FRAME11寬度一半+間隙
                X2 = ALL_range[16][1] + FRAME_11_width[i] * scale / 2 + gap#box_17_Xmin+FRAME11寬度一半+間隙
                Y1 = ALL_range[16][3] + FRAME_11_height[i] * scale / 2 + gap#box_17_Ymin+FRAME11高度一半+間隙
                Y2 = ALL_range[16][2] - 90 * scale / 2 - gap#box_17_Ymax-FRAME11厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME39':
                p = 12#上視圖(翻轉180度)投影
                p2 =1#後視圖
                X1 = ALL_range[17][1] + 145 * scale / 2 + gap#box_18_Xmin+FRAME39深度一半+間隙
                X2 = ALL_range[17][1] + 145 * scale / 2 + gap  # box_18_Xmin+FRAME39深度一半+間隙
                Y1 = ALL_range[17][2] - 300 * scale / 2 - gap#box_18_Ymax-FRAME39寬度一半-間隙
                Y2 = ALL_range[17][3] + 19 * scale / 2 + gap  # box_18_Ymin+FRAME39厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME17':
                p = 12#上視圖(翻轉180度)投影
                p2 =4 #左視圖投影
                X1 = ALL_range[18][1] + 75 * scale / 2 + gap#box_19_Xmin+FRAME17寬度一半+間隙
                X2 = ALL_range[18][1] + 75 * scale / 2 + gap#box_19_Xmin+FRAME17寬度一半+間隙
                Y1 = ALL_range[18][2] - 75 * scale / 2 - gap#box_19_Ymax-FRAME17深度一半-間隙
                Y2 = ALL_range[18][3] + 48 * scale / 2 + gap#box_19_Ymin+FRAME17高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME3':
                p = 0#前視圖投影
                p2 = 2#上視圖投影
                X1 = ALL_range[19][1] + (R[i]+180) * scale / 2 + gap#box_20_Xmin+FRAME3寬度一半+間隙
                X2 = ALL_range[19][1] + (R[i]+180) * scale / 2 + gap#box_20_Xmin+FRAME3寬度一半+間隙
                Y1 = ALL_range[19][3] + (Z[i] - T[i] - 40) * scale / 2 + gap#box_20_Ymin+FRAME3高度一半+間隙
                Y2 = ALL_range[19][2] - 50 * scale / 2 - gap#box_20_Ymax-FRAME3厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME13':
                p = 2#上視圖投影
                p2 = 9#右視圖(左旋)投影
                X1 = ALL_range[20][1] + 85 * scale / 2 + gap#box_21_Xmin+FRAME13寬度一半+間隙
                X2 = ALL_range[20][0] - 55 * scale / 2 - gap#box_21_Xmax-FRAME13高度一半-間隙
                Y1 = ALL_range[20][2] - FRAME_13_depth[i] * scale / 2 - gap#box_21_Ymax-FRAME13深度一半-間隙
                Y2 = ALL_range[20][2] - FRAME_13_depth[i] * scale / 2 - gap  # box_21_Ymax-FRAME13深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME14':
                p = 6#上視圖(左旋)投影
                p2 = 4#左視圖投影
                X1 = ALL_range[21][1] + 75 * scale / 2 + gap#box_22_Xmin+FRAME14深度一半+間隙
                X2 = ALL_range[21][1] + 75 * scale / 2 + gap#box_22_Xmin+FRAME14深度一半+間隙
                Y1 = ALL_range[21][2] - 75 * scale / 2 - gap#box_22_Ymax-FRAME14寬度一半-間隙
                Y2 = ALL_range[21][3] + 70 * scale / 2 + gap#box_22_Ymin+FRAME14高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME22':
                p = 4#左視圖投影
                X1 = ALL_range[22][1] + 460 * scale / 2 + gap#box_23_Xmin+FRAME22寬度一半+間隙
                Y1 = ALL_range[22][2] - 280 * scale / 2 + gap#box_23_Ymax-FRAME22高度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
            elif x == 'FRAME37':
                p = 1#後視圖投影
                X1 = ALL_range[23][1] + 140 * scale / 2 + gap#box_24_Xmin+FRAME37寬度一半+間隙
                Y1 = ALL_range[23][2] - 80 * scale / 2 - gap#box_24_Ymax-FRAME37高度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
            elif x == 'FRAME8':
                p = 13#下視圖(左旋)投影
                p2 = 10#後視圖(左旋)投影
                X1 = ALL_range[24][0] - FRAME_8_width[i] * scale / 2 - gap#box_25_Xmax-FRAME8寬度一半-間隙
                X2 = ALL_range[24][1] + 40 * scale / 2 + gap#box_25_Xmin+FRAME8厚度一半+間隙
                Y1 = ALL_range[24][2] - 300 * scale / 2 - gap#box_25_Ymax-FRAME8深度一半-間隙
                Y2 = ALL_range[24][2] - 300 * scale / 2 - gap  # box_25_Ymax-FRAME8深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME29':
                p = 7#下視圖(180度)投影
                p2 = 9#右視圖(左旋)投影
                X1 = ALL_range[25][0] - (R[i]+180) * scale / 2 - gap#box_26_Xmax-FRAME29寬度一半-間隙
                X2 = ALL_range[25][1] + 65 * scale / 2 + gap#box_26_Xmin+FRAME29深度一半+間隙
                Y1 = ALL_range[25][2] - 30 * scale / 2 - gap#box_26_Ymax-FRAME29厚度一半-間隙
                Y2 = ALL_range[25][2] - 30 * scale / 2 - gap  # box_26_Ymax-FRAME29厚度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME20':
                p = 0#前視圖投影
                p2 = 2#上視圖投影
                X1 = ALL_range[26][1] + (R[i] + 180) * scale / 2 + gap#box_27_Xmin+FRAME20寬度一半+間隙
                X2 = ALL_range[26][1] + (R[i] + 180) * scale / 2 + gap#box_27_Xmin+FRAME20寬度一半+間隙
                Y2 = ALL_range[26][2] - 50 * scale / 2 - gap#box_27_Ymax-FRAME20厚度一半-間隙
                Y1 = ALL_range[26][3] + FRAME20_H[i] * scale / 2 + gap#box_27_Ymin+FRAME20高度一半+間隙
                a = Y1
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME33':
                p = 12#上視圖(轉180度)投影
                p2 = 1#後視圖投影
                X1 = ALL_range[27][1] + 50 * scale / 2 + gap#box_28_Xmin+FRAME33深度一半+間隙
                X2 = ALL_range[27][1] + 50 * scale / 2 + gap  # box_28_Xmin+FRAME33深度一半+間隙
                Y1 = ALL_range[27][2] - 180 * scale / 2 - gap#box_28_Ymax-FRAME33寬度一半-間隙
                Y2 = ALL_range[27][3] + 50 * scale / 2 + gap#box_28_Ymin+FRAME33高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME48':
                p = 4#左視圖投影
                p2 = 0#前視圖投影
                X1 = ALL_range[28][1] + 32 * scale / 2 + gap#box_29_Xmin+FRAME48寬度一半+間隙
                X2 = ALL_range[28][0] - 82 * scale / 2 - gap#box_29_Xmax-FRAME48深度一半-間隙
                Y1 = ALL_range[28][2] - 32 * scale / 2 - gap#box_29_Ymax-FRAME48高度一半-間隙
                Y2 = ALL_range[28][2] - 32 * scale / 2 - gap  # box_29_Ymax-FRAME48高度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME49':
                p = 0#前視圖投影
                p2 = 5#右視圖投影
                X1 = ALL_range[29][1] + 30 * scale / 2 + gap#box_30_Xmin+FRAME49寬度一半+間隙
                X2 = ALL_range[29][0] - 54 * scale / 2 - gap#box_30_Xmax-FRAME49深度一半-間隙
                Y1 = ALL_range[29][2] - 30 * scale / 2 - gap#box_30_Ymax-FRAME49高度一半-間隙
                Y2 = ALL_range[29][2] - 30 * scale / 2 - gap  # box_30_Ymax-FRAME49高度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            else :
                break
        elif l == 0:
            if x == 'FRAME1':
                p = 4  # 左視圖投影
                p2 = 5  # 右視圖投影
                X1 = (ALL_range[0][0] + ALL_range[0][1]) / 4 + 10  # 第一個投影X中心+尺寸標註預留量
                X2 = (ALL_range[0][0] + ALL_range[0][1]) * 3 / 4  # 第二個投影X中心(X中心為將X範圍分成兩塊之中心)
                Y1 = (ALL_range[0][2] + ALL_range[0][3]) * 3 / 4  # 第一個投影Y中心
                Y2 = (ALL_range[0][2] + ALL_range[0][3]) * 3 / 4  # 第二個投影Y中心(Y中心為將Y範圍分成兩塊之上半部中心)
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME2':
                p = 4  # 左視圖投影
                p2 = 5  # 右視圖投影
                X1 = (ALL_range[0][0] + ALL_range[0][1]) / 4 + 10  # 第一個投影X中心+尺寸標註預留量
                X2 = (ALL_range[0][0] + ALL_range[0][1]) * 3 / 4  # 第二個投影X中心(X中心為將X範圍分成兩塊之中心)
                Y1 = (ALL_range[0][2] + ALL_range[0][3]) / 4  # 第一個投影Y中心
                Y2 = (ALL_range[0][2] + ALL_range[0][3]) / 4  # 第二個投影Y中心(Y中心為將Y範圍分成兩塊之上半部中心)
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME30':
                p = 0  # 前視圖投影
                p2 = 5  # 右視圖投影
                p3 = 3  # 下視圖投影
                X1 = ALL_range[1][1] + (R_15[i] + 180) * (scale) / 2 + gap  # BOX_2_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[1][0] - (40) * (scale) / 2 - gap  # BOX_2_Xmax-FRAME30厚度一半-間隙
                X3 = ALL_range[1][1] + (R_15[i] + 180) * (scale) / 2 + gap  # BOX_2_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[1][2] - (FRAME44_height[i]) * (
                    scale) / 2 - gap  # BOX_2_Ymax-前視圖高度-間隙(取完整高度是為了使零件圖與右壁板高度相近)
                Y2 = ALL_range[1][2] - (FRAME44_height[i]) * (
                    scale) / 2 - gap  # BOX_2_Ymax-前視圖高度-間隙(取完整高度是為了使零件圖與右壁板高度相近)
                Y3 = ALL_range[1][3] + (40) * (scale) / 2 + gap  # BOX_2_Ymin+FRAME30厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_projection(p3, X3, Y3, x, scale)
            elif x == 'FRAME44':
                p = 0  # 前視圖投影
                p2 = 3  # 下視圖投影
                X1 = ALL_range[2][1] + (R_15[i] + 180) * (scale) / 2 + gap  # box_3_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[2][1] + (R_15[i] + 180) * (scale) / 2 + gap  # box_3_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[2][2] - (FRAME44_height[i]) * (scale) / 2 - gap  # box_3_Ymax+FRAME44高度一半-間隙
                Y2 = ALL_range[2][3] + 40 * (scale) / 2 + gap  # box_3_Ymin+FRAME44厚度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME41':
                p = 6  # 上視圖(左轉向)投影
                p2 = 5  # 右視圖投影
                X1 = ALL_range[3][1] + (R_15[i] + 180) * (scale) / 2 + gap  # box_4_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[3][1] + (R_15[i] + 180) * (scale) / 2 + gap  # box_4_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[3][2] - (FRAME_41_depth[i]) * (scale) / 2 - gap  # box_3_Ymax+FRAME41高度一半-間隙
                Y2 = ALL_range[3][3] + 22 * (scale) / 2 + gap  # box_3_Yin+FRAME41厚度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME34':
                p = 1  # 後視圖投影
                p2 = 5  # 右視圖投影
                p3 = 7  # 下視圖翻轉180度投影
                X1 = ALL_range[4][0] - 182 * scale / 2 - gap  # box_5_Xmax-FRAME34寬度一半-間隙
                X2 = ALL_range[4][1] + 40 * scale / 2 + gap  # box_5_Xmin+FRAME34厚度一半+間隙
                X3 = ALL_range[4][0] - 182 * scale / 2 - gap  # box_5_Xmax-FRAME34寬度一半-間隙
                Y1 = ALL_range[4][2] - 80 * scale / 2 - gap  # box_5_Ymax-FRAME34深度一半-間隙
                Y2 = ALL_range[4][2] - 80 * scale / 2 - gap  # box_5_Ymax-FRAME34深度一半-間隙
                Y3 = ALL_range[4][3] + 40 * scale / 2 + gap  # box_5_Ymax+FRAME34厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_projection(p3, X3, Y3, x, scale)
            elif x == 'FRAME9':
                p = 2  # 上視圖投影
                p2 = 0  # 前視圖投影
                X1 = ALL_range[5][1] + (R_15[i] + 180) * scale / 2 + gap  # box_6_Xmin+FRAME9寬度一半+間隙
                X2 = ALL_range[5][1] + (R_15[i] + 180) * scale / 2 + gap  # box_6_Xmin+FRAME9寬度一半+間隙
                Y1 = ALL_range[5][2] - 50 * scale / 2 - gap  # box_6_Ymax-FRAME9厚度一半-間隙
                Y2 = ALL_range[5][3] + (Z[i] - T[i] - 40) * scale / 2 + gap  # box_6_Ymin+FRAME9高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME45':
                p = 2  # 上視圖投影
                p2 = 9  # 右側視圖(左旋轉)投影
                X1 = ALL_range[6][1] + R_15[i] * scale / 2 + gap  # box_7_Xmin+FRAME45寬度一半+間隙
                X2 = ALL_range[6][0] - 19 * scale / 2 - gap  # box_7_Xmax+FRAME45厚度一半-間隙
                Y1 = ALL_range[6][2] - 476 * scale / 2 - gap  # box_7_Ymax-FRAME45深度一半-間隙
                Y2 = ALL_range[6][2] - 476 * scale / 2 - gap  # box_7_Ymax-FRAME45深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME43':
                p = 0  # 前視圖投影
                X1 = ALL_range[7][1] + R_15[i] * scale / 2 + gap  # box_8_Xmin+FRAME43寬度一半+間隙
                Y1 = ALL_range[7][2] - 150 * scale / 2 - gap  # box_8_Ymax-FRAME45深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
            elif x == 'FRAME7':
                p = 8  # 下視圖(右旋)投影
                X1 = ALL_range[8][1] + FRAME_7_15_width[i] * scale / 2 + gap  # box_9_Xmin+FRAME7寬度一半+間隙
                Y1 = ALL_range[8][2] - 300 * scale / 2 - gap  # box_9_Ymax-FRAME7深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
            elif x == 'MAIN_GEAR2':
                p = 5  # 右視圖投影
                X1 = ALL_range[9][1] + 125 * scale / 2 + gap  # box_10_Xmin+MAIN_GEAR2深度一半+間隙
                Y1 = ALL_range[9][2] - 270 * scale / 2 - gap  # box_10_Ymax-MAIN_GEAR2高度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
            elif x == 'FRAME32':
                p = 6  # 上視圖(左旋)投影
                p2 = 10  # 背視圖(左旋)投影
                X1 = ALL_range[10][1] + R_15[i] * scale / 2 + gap  # box_11_Xmin+FRAME32寬度一半+間隙
                X2 = ALL_range[10][0] - 19 * scale / 2 - gap  # box_11_Xmax-FRAME32厚度一半-間隙
                Y1 = ALL_range[10][2] - 429 * scale / 2 - gap  # box_11_Ymax-FRAME深度一半-間隙
                Y2 = ALL_range[10][2] - 429 * scale / 2 - gap  # box_11_Ymax-FRAME深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME47':
                p = 6  # 上視圖(左旋)投影
                p2 = 1  # 背視圖投影
                X1 = ALL_range[11][1] + 50 * scale / 2 + gap  # box_12_Xmin+FRAME47寬度一半+間隙
                X2 = ALL_range[11][0] - 19 * scale / 2 - gap  # box_12_Xmax-FRAME47厚度一半-間隙
                Y1 = ALL_range[11][2] - 74 * scale / 2 - gap  # box_12_Ymax-FRAME47深度一半-間隙
                Y2 = ALL_range[11][2] - 74 * scale / 2 - gap  # box_12_Ymax-FRAME47深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME23':
                p = 6  # 上視圖(左旋)投影
                p2 = 5  # 右視圖投影
                X1 = ALL_range[12][1] + 99.35 * scale / 2 + gap  # box_13_Xmin+FRAME23深度一半+間隙
                X2 = ALL_range[12][1] + 99.35 * scale / 2 + gap  # box_13_Xmin+FRAME23深度一半+間隙
                Y1 = ALL_range[12][2] - 35 * scale / 2 - gap  # box_13_Ymax-FRAME23寬度一半-間隙
                Y2 = ALL_range[12][3] + 150 * scale / 2 + gap  # box_13_Ymin+FRAME23高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME31':
                p = 0  # 前視圖投影
                p2 = 5  # 右視圖投影
                p3 = 3  # 下視圖投影
                p4 = 0  # 前視圖投影
                p5 = 3  # 下視圖投影
                X1 = ALL_range[13][1] + 40 * scale / 2 + gap  # box_14_Xmin+FRAME31寬度一半+間隙
                X2 = (ALL_range[13][1] + ALL_range[13][0]) / 2  # box_14_X中心
                X3 = ALL_range[13][1] + 40 * scale / 2 + gap  # box_14_Xmin+FRAME31寬度一半+間隙
                X4 = ALL_range[13][0] - 40 * scale / 2 - gap  # box_14_Xmax-FRAME31寬度一半-間隙
                X5 = ALL_range[13][0] - 40 * scale / 2 - gap  # box_14_Xmax-FRAME31寬度一半-間隙
                Y1 = ALL_range[13][2] - 150 * scale / 2 - gap  # box_14_Ymax-FRAME31高度一半-間隙
                Y2 = ALL_range[13][2] - 150 * scale / 2 - gap  # box_14_Ymax-FRAME31高度一半-間隙
                Y3 = ALL_range[13][3] + 45 * scale / 2 + gap  # box_14_Ymin+FRAME31厚度一半+間隙
                Y4 = ALL_range[13][2] - 150 * scale / 2 - gap  # box_14_Ymax-FRAME31高度一半-間隙
                Y5 = ALL_range[13][3] + 45 * scale / 2 + gap  # box_14_Ymin+FRAME31厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_projection(p3, X3, Y3, x, scale)
                N.Material_diagram_projection(p4, X4, Y4, x, scale)
                N.Material_diagram_projection(p5, X5, Y5, x, scale)
            elif x == 'FRAME24':
                p = 6  # 上視圖(左旋)投影
                p2 = 1  # 背視圖投影
                p3 = 5  # 右視圖投影
                X1 = ALL_range[14][1] + 99.35 * scale / 2 + gap  # box_15Xmin+FRAME24深度一半+間隙
                X2 = ALL_range[14][0] - 35 * scale / 2 + gap  # box_15Xmax-FRAME24寬度一半-間隙
                X3 = ALL_range[14][1] + 99.35 * scale / 2 + gap  # box_15Xmin+FRAME24深度一半+間隙
                Y1 = ALL_range[14][2] - 35 * scale / 2 - gap  # box_15Ymax-FRAME24寬度一半-間隙
                Y2 = ALL_range[14][3] + 150 * scale / 2 + gap  # box_15Ymin+FRAME24高度一半+間隙
                Y3 = ALL_range[14][3] + 150 * scale / 2 + gap  # box_15Ymin+FRAME24高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
                N.Material_diagram_projection(p3, X3, Y3, x, scale)
            elif x == 'FRAME38':
                p = 11  # 上視圖(右旋)投影
                X1 = ALL_range[15][1] + 290 * scale / 2 + gap  # box_16_Xmin+FRAME38寬度一半+間隙
                Y1 = ALL_range[15][2] - 145 * scale / 2 - gap  # box_16_Ymax-FRAME38深度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
            elif x == 'FRAME11':
                p = 5  # 右視圖投影
                p2 = 6  # 上視圖(左旋)投影
                X1 = ALL_range[16][1] + FRAME_11_15_width[i] * scale / 2 + gap  # box_17_Xmin+FRAME11寬度一半+間隙
                X2 = ALL_range[16][1] + FRAME_11_15_width[i] * scale / 2 + gap  # box_17_Xmin+FRAME11寬度一半+間隙
                Y1 = ALL_range[16][3] + FRAME_11_height[i] * scale / 2 + gap  # box_17_Ymin+FRAME11高度一半+間隙
                Y2 = ALL_range[16][2] - 90 * scale / 2 - gap  # box_17_Ymax-FRAME11厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME39':
                p = 12  # 上視圖(翻轉180度)投影
                p2 = 1  # 後視圖
                X1 = ALL_range[17][1] + 145 * scale / 2 + gap  # box_18_Xmin+FRAME39深度一半+間隙
                X2 = ALL_range[17][1] + 145 * scale / 2 + gap  # box_18_Xmin+FRAME39深度一半+間隙
                Y1 = ALL_range[17][2] - 300 * scale / 2 - gap  # box_18_Ymax-FRAME39寬度一半-間隙
                Y2 = ALL_range[17][3] + 19 * scale / 2 + gap  # box_18_Ymin+FRAME39厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME17':
                p = 12  # 上視圖(翻轉180度)投影
                p2 = 4  # 左視圖投影
                X1 = ALL_range[18][1] + 75 * scale / 2 + gap  # box_19_Xmin+FRAME17寬度一半+間隙
                X2 = ALL_range[18][1] + 75 * scale / 2 + gap  # box_19_Xmin+FRAME17寬度一半+間隙
                Y1 = ALL_range[18][2] - 75 * scale / 2 - gap  # box_19_Ymax-FRAME17深度一半-間隙
                Y2 = ALL_range[18][3] + 48 * scale / 2 + gap  # box_19_Ymin+FRAME17高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME3':
                p = 0  # 前視圖投影
                p2 = 2  # 上視圖投影
                X1 = ALL_range[19][1] + (R_15[i] + 180) * scale / 2 + gap  # box_20_Xmin+FRAME3寬度一半+間隙
                X2 = ALL_range[19][1] + (R_15[i] + 180) * scale / 2 + gap  # box_20_Xmin+FRAME3寬度一半+間隙
                Y1 = ALL_range[19][3] + (Z[i] - T[i] - 40) * scale / 2 + gap  # box_20_Ymin+FRAME3高度一半+間隙
                Y2 = ALL_range[19][2] - 50 * scale / 2 - gap  # box_20_Ymax-FRAME3厚度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME13':
                p = 2  # 上視圖投影
                p2 = 9  # 右視圖(左旋)投影
                X1 = ALL_range[20][1] + 85 * scale / 2 + gap  # box_21_Xmin+FRAME13寬度一半+間隙
                X2 = ALL_range[20][0] - 55 * scale / 2 - gap  # box_21_Xmax-FRAME13高度一半-間隙
                Y1 = ALL_range[20][2] - FRAME_13_depth[i] * scale / 2 - gap  # box_21_Ymax-FRAME13深度一半-間隙
                Y2 = ALL_range[20][2] - FRAME_13_depth[i] * scale / 2 - gap  # box_21_Ymax-FRAME13深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME14':
                p = 6  # 上視圖(左旋)投影
                p2 = 4  # 左視圖投影
                X1 = ALL_range[21][1] + 75 * scale / 2 + gap  # box_22_Xmin+FRAME14深度一半+間隙
                X2 = ALL_range[21][1] + 75 * scale / 2 + gap  # box_22_Xmin+FRAME14深度一半+間隙
                Y1 = ALL_range[21][2] - 75 * scale / 2 - gap  # box_22_Ymax-FRAME14寬度一半-間隙
                Y2 = ALL_range[21][3] + 70 * scale / 2 + gap  # box_22_Ymin+FRAME14高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME22':
                p = 4  # 左視圖投影
                X1 = ALL_range[22][1] + 460 * scale / 2 + gap  # box_23_Xmin+FRAME22寬度一半+間隙
                Y1 = ALL_range[22][2] - 280 * scale / 2 + gap  # box_23_Ymax-FRAME22高度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
            elif x == 'FRAME37':
                p = 1  # 後視圖投影
                X1 = ALL_range[23][1] + 140 * scale / 2 + gap  # box_24_Xmin+FRAME37寬度一半+間隙
                Y1 = ALL_range[23][2] - 80 * scale / 2 - gap  # box_24_Ymax-FRAME37高度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
            elif x == 'FRAME8':
                p = 13  # 下視圖(左旋)投影
                p2 = 10  # 後視圖(左旋)投影
                X1 = ALL_range[24][0] - FRAME_8_width[i] * scale / 2 - gap  # box_25_Xmax-FRAME8寬度一半-間隙
                X2 = ALL_range[24][1] + 40 * scale / 2 + gap  # box_25_Xmin+FRAME8厚度一半+間隙
                Y1 = ALL_range[24][2] - 300 * scale / 2 - gap  # box_25_Ymax-FRAME8深度一半-間隙
                Y2 = ALL_range[24][2] - 300 * scale / 2 - gap  # box_25_Ymax-FRAME8深度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME29':
                p = 7  # 下視圖(180度)投影
                p2 = 9  # 右視圖(左旋)投影
                X1 = ALL_range[25][0] - (R[i] + 180) * scale / 2 - gap  # box_26_Xmax-FRAME29寬度一半-間隙
                X2 = ALL_range[25][1] + 65 * scale / 2 + gap  # box_26_Xmin+FRAME29深度一半+間隙
                Y1 = ALL_range[25][2] - 30 * scale / 2 - gap  # box_26_Ymax-FRAME29厚度一半-間隙
                Y2 = ALL_range[25][2] - 30 * scale / 2 - gap  # box_26_Ymax-FRAME29厚度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME20':
                p = 0  # 前視圖投影
                p2 = 2  # 上視圖投影
                X1 = ALL_range[26][1] + (R[i] + 180) * scale / 2 + gap  # box_27_Xmin+FRAME20寬度一半+間隙
                X2 = ALL_range[26][1] + (R[i] + 180) * scale / 2 + gap  # box_27_Xmin+FRAME20寬度一半+間隙
                Y2 = ALL_range[26][2] - 50 * scale / 2 - gap  # box_27_Ymax-FRAME20厚度一半-間隙
                Y1 = ALL_range[26][3] + FRAME20_H[i] * scale / 2 + gap  # box_27_Ymin+FRAME20高度一半+間隙
                a = Y1
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME33':
                p = 12  # 上視圖(轉180度)投影
                p2 = 1  # 後視圖投影
                X1 = ALL_range[27][1] + 50 * scale / 2 + gap  # box_28_Xmin+FRAME33深度一半+間隙
                X2 = ALL_range[27][1] + 50 * scale / 2 + gap  # box_28_Xmin+FRAME33深度一半+間隙
                Y1 = ALL_range[27][2] - 180 * scale / 2 - gap  # box_28_Ymax-FRAME33寬度一半-間隙
                Y2 = ALL_range[27][3] + 50 * scale / 2 + gap  # box_28_Ymin+FRAME33高度一半+間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME48':
                p = 4  # 左視圖投影
                p2 = 0  # 前視圖投影
                X1 = ALL_range[28][1] + 32 * scale / 2 + gap  # box_29_Xmin+FRAME48寬度一半+間隙
                X2 = ALL_range[28][0] - 82 * scale / 2 - gap  # box_29_Xmax-FRAME48深度一半-間隙
                Y1 = ALL_range[28][2] - 32 * scale / 2 - gap  # box_29_Ymax-FRAME48高度一半-間隙
                Y2 = ALL_range[28][2] - 32 * scale / 2 - gap  # box_29_Ymax-FRAME48高度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            elif x == 'FRAME49':
                p = 0  # 前視圖投影
                p2 = 5  # 右視圖投影
                X1 = ALL_range[29][1] + 30 * scale / 2 + gap  # box_30_Xmin+FRAME49寬度一半+間隙
                X2 = ALL_range[29][0] - 54 * scale / 2 - gap  # box_30_Xmax-FRAME49深度一半-間隙
                Y1 = ALL_range[29][2] - 30 * scale / 2 - gap  # box_30_Ymax-FRAME49高度一半-間隙
                Y2 = ALL_range[29][2] - 30 * scale / 2 - gap  # box_30_Ymax-FRAME49高度一半-間隙
                N.Material_diagram_projection(p, X1, Y1, x, scale)
                N.Material_diagram_projection(p2, X2, Y2, x, scale)
            else :
                break

# mprog.set_CATIA_workbench_env()
mprog.OPEN_Drawing()
test(5 , 1)
# N.Material_diagram_balloons('FRAME1' , -200 , (A[5]/2+45)*15)