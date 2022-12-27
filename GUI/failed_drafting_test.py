

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


#圖框範圍
ALL_range = []
gap = 5
box_gap = 30#虛擬方框間間隙
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

def scale_Adjustment3(i , j):
    scale_p = 0
    while True:
        scale_p += 1
        print(scale_p)
        if scale_p % 2 != 0 and scale_p % 5 != 0 and scale_p % 10 != 0:  # 比例要為2,5,10的倍數
            continue
        #----------------------------基本數值-----------------------------------------
        scale = 1 / scale_p
        box_1_range = [box_1_Xmax, box_1_Xmin, box_1_Ymax, box_1_Ymin]
        box_1_rangeX = [box_1_Xmax - box_1_Xmin]#預設範圍
        box_1_rangeY = [box_1_Ymax - box_1_Ymin]#預設範圍
        box_1_center = [(box_1_Xmax + box_1_Xmin) / 2 , (box_1_Ymax + box_1_Ymin) / 2]
        w_scale = A[i] * scale
        h_scale = H[i] * scale
        d_scale = B[i] * scale
        # --------------------------------box_1----------------------------
        part_1_width = [2 * d_scale + box_width_gap]
        part_1_heigth = [h_scale + box_heigth_gap]
        if part_1_width > box_1_rangeX:  # 元件寬度判斷
            continue
        elif part_1_heigth > box_1_rangeY:  # 元件高度判斷
            continue
        break
    return scale_p , box_1_center


def drafting_parameter_calculation3(i, l , scale_p , ):
    #迴圈次數
    F = 0
    #各方框迴圈X，Y方向平移次數次數
    F2X = 0
    F2Y = 0
    F3X = 0
    F3Y = 0
    F4X = 0
    F4Y = 0
    F5X = 0
    F5Y = 0
    while True:
        F += 1
        scale = 1 / scale_p
        box_1_range = [round(box_1_Xmax , 2), round(box_1_Xmin , 2), round(box_1_Ymax , 2), round(box_1_Ymin , 2)]#方框1的範圍[Xmax , Xmin , Ymax , Ymin]
        if F == 1 :
            ALL_range.append(box_1_range)
    #-------------------------------box_2範圍建立--------------------------------------------------------------------------
        box_2_rangeX = (R[i]+180 + 40)*scale+2*gap+50#FRAME33 的寬+厚+間隙
        box_2_rangeY = (FRAME44_height[i]+40)*scale+2*gap +50#FRAME33 的高+厚+間隙
        box_2_center = [ALL_range[0][0] + F2X*1 , ALL_range[0][2] - F2Y*1]#初始中心為BOX1的右上，依照迴圈次數慢慢往右往下移動
        if box_2_center[0] < box_1_Xmax + box_2_rangeX/2 + box_gap:#BOX2寬度判斷(BOX1寬+BOX2邊緣至中心點距離)
            F2X += 1
            continue
        elif box_2_center[1] + box_2_rangeY > box_1_Ymax:#BOX2高度判斷(BOX2中心+BOX2Y範圍須<BOX1的Ymax)
            F2Y += 1
            continue
        box_2_range = [round(box_2_center[0] + box_2_rangeX/2 , 2) , round(box_2_center[0] - box_2_rangeX/2 , 2) ,
                       round(box_2_center[1] + box_2_rangeY/2 , 2) , round(box_2_center[1] - box_2_rangeY/2 , 2)]#方框2的範圍[Xmax , Xmin , Ymax , Ymin]
        if F3X == 0:
            ALL_range.append(box_2_range)
#--------------------------------------------------box_3範圍建立---------------------------------------------------------
        box_3_rangeX = (R[i]+180)*scale+2*gap
        box_3_rangeY = (FRAME44_height[i]+40)*scale+2*gap + 50
        box_3_center = [ALL_range[1][1] + F3X * 1, ALL_range[1][3] - F3Y * 1]#初始中心為BOX2的左下，依照迴圈次數慢慢往右下移動
        if box_3_center[0] < box_1_Xmax + box_3_rangeX/2 + box_gap:#X方向平移
            F3X += 1
            continue
        elif box_3_center[1] + box_3_rangeY/2 + box_gap > ALL_range[1][3]:#Y方向平移
            F3Y += 1
            continue
        box_3_range = [round(box_3_center[0] + box_3_rangeX/2 , 2) , round(box_3_center[0] - box_3_rangeX/2 , 2) ,
                       round(box_3_center[1] + box_3_rangeY/2 , 2) , round(box_3_center[1] - box_3_rangeY/2 , 2)]#方框3的範圍[Xmax , Xmin , Ymax , Ymin]
        if F4X == 0:
            ALL_range.append(box_3_range)
#-----------------------------------------------box_4範圍建立----------------------------------------------------------
        box_4_rangeX = (R[i]+180)*scale+2*gap
        box_4_rangeY = (FRAME_41_depth[i]+22)*scale + 2*gap + 50
        box_4_center = [ALL_range[2][1] + F4X*1 , ALL_range[2][3] + F4Y*1]#初始中心為BOX3的左下，依照迴圈次數慢慢往右下移動
        if box_4_center[0] < box_1_Xmax + box_4_rangeX/2 + box_gap:#X方向平移
            F4X += 1
            continue
        elif box_4_center[1] + box_4_rangeY/2 + box_gap > ALL_range[2][3]:#Y方向平移
            F4Y -= 1
            continue
        box_4_range = [round(box_4_center[0] + box_4_rangeX/2 , 2) , round(box_4_center[0] - box_4_rangeX/2 , 2) ,
                       round(box_4_center[1] + box_4_rangeY/2 , 2) , round(box_4_center[1] - box_4_rangeY/2 , 2)]#方框4的範圍[Xmax , Xmin , Ymax , Ymin]
        if F5X == 0:
            ALL_range.append(box_4_range)

        break
    return ALL_range