import win32com.client as win32
i = 0#型號順序
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
#圖框範圍
ALL_range = []
ALL_center = []
temporary_range = []
box_1_range = []
scale = 1
gap = 5
box_gap = 22#虛擬方框內間隙
drafting_min_Y = 44
drafting_max_Y = 810
drafting_min_X = 20
drafting_max_X = 1169
#第一虛擬方框位置
box_1_Xmax = 490
box_1_Ymax = 810
box_1_Xmin = 25
box_1_Ymin = 25
box_width_gap = 80+2*gap#虛擬方框一的寬度間隙
box_heigth_gap = 150+2*gap#虛擬方框一的高度間隙

projection = {"front view":(0, 1 , 0, 0 , 0, 1) , "Rear view":(0, -1, 0, 0 , 0, 1) , "top view":(0, 1 , 0, -1 , 0 , 0)
    , "bottom view":(0, 1, 0, 1, 0, 0), "Left view":(1, 0, 0, 0, 0, 1) , "right view":(-1, 0, 0, 0, 0, 1)
              , 'top view(left horizontal)':(-1 , 0 , 0 , 0 , -1 , 0)  , "top view(Y inverse)":(0 , -1 , 0 , -1 , 0 , 0) ,
              'bottom view(right horizontal)':(-1 , 0 , 0 , 0 , 1 , 0) , 'right view(left horizontal)':(0 , 0 , -1 , -1 , 0 , 0),
              'Rear view(left horizontal)':(0 , 0 , -1 , 0 , -1 , 0) , 'top view(right horizontal)':(1 , 0 , 0 , 0 , 1 , 0) ,
              'top view(180 degree)':(0 , -1 , 0 , 1 , 0 , 0) , 'bottom view(left horizontal)':(1 , 0 , 0 , 0 , -1 , 0)}
U = ["front view" , "Rear view" , "top view" , "bottom view" , "Left view" , "right view" , 'top view(left horizontal)' ,
     "top view(Y inverse)" , 'bottom view(right horizontal)' , 'right view(left horizontal)','Rear view(left horizontal)',
     'top view(right horizontal)' , 'top view(180 degree)' , 'bottom view(left horizontal)']

projection_file_name_list = ['FRAME30', 'FRAME44', 'FRAME41', 'FRAME34', 'FRAME9', 'FRAME45', 'FRAME43', 'FRAME7',
                                 'MAIN_GEAR2', 'FRAME32' , 'FRAME47' , 'FRAME23' , 'FRAME31', 'FRAME24', 'FRAME38' , 'FRAME11', 'FRAME39', 'FRAME17',
                                 'FRAME3', 'FRAME14', 'FRAME22', 'FRAME13' ,
                                 'FRAME37', 'FRAME8', 'FRAME29', 'FRAME20', 'FRAME33' , 'FRAME48' , 'FRAME49']
file_list_range = {
    'FRAME30': [(R[i] + 180 + 40) * scale + 2 * gap + box_gap,
                (FRAME44_height[i] + 40) * scale + 2 * gap + box_gap],
    'FRAME44': [(R[i] + 180) * scale + 2 * gap + box_gap,
                (FRAME44_height[i] + 40) * scale + 2 * gap + box_gap],
    'FRAME41': [(R[i] + 180) * scale + 2 * gap + box_gap,
                (FRAME_41_depth[i] + 22) * scale + 2 * gap + box_gap],
    'FRAME34': [(182 + 40) * scale + 2 * gap + box_gap, (80 + 40) * scale + 2 * gap + box_gap],
    'FRAME9': [(R[i] + 180) * scale + 2 * gap + box_gap,
               (((Z[i] - T[i] - 40) + 50) * scale + 2 * gap + box_gap)],
    'FRAME45': [(R[i] + 19) * scale + 2 * gap + box_gap, 476 * scale + 2 * gap + box_gap],
    'FRAME43': [R[i] * scale + 2 * gap + box_gap, 150 * scale + 2 * gap + box_gap / 2],
    'FRAME7': [FRAME_7_width[i] * scale + 2 * gap + box_gap, 300 * scale + 2 * gap + box_gap / 2],
    'MAIN_GEAR2': [125 * scale + 2 * gap + box_gap, 270 * scale + 2 * gap + box_gap / 2],
    'FRAME32': [(R[i] + 19) * scale + 2 * gap + box_gap, 429 * scale + 2 * gap + box_gap / 2],
    'FRAME31': [(40 + 45) * scale + 2 * gap + box_gap + 50, (150 + 45) * scale + 2 * gap + box_gap / 2],
    'FRAME38': [290 * scale + 2 * gap + box_gap, 145 * scale + 2 * gap + box_gap / 2],
    'FRAME11': [FRAME_11_width[i] * scale + 2 * gap + box_gap, (FRAME_11_height[i] + 90) * scale + 2 * scale + box_gap],
    'FRAME39': [145 * scale + 2 * gap + box_gap, (300 + 19) * scale + 2 * gap + box_gap],
    'FRAME17': [75 * scale + 2 * gap + box_gap, (75 + 48) * scale + 2 * gap + box_gap / 2],
    'FRAME3': [(R[i] + 180) * scale + 2 * gap + box_gap, ((Z[i] - T[i] - 40 + 50) * scale + 2 * gap + box_gap)],
    'FRAME14': [75 * scale + 2 * gap + box_gap, (75 + 70) * scale + 2 * gap + box_gap],
    'FRAME22': [460 * scale + 2 * gap + box_gap, 280 * scale + 2 * gap + box_gap / 2 ],
    'FRAME37': [140 * scale + 2 * gap + box_gap, 80 * scale + 2 * gap + box_gap / 2],
    'FRAME8': [(FRAME_8_width[i] + 40) * scale + 2 * gap + box_gap, 300 * scale + 2 * gap + box_gap],
    'FRAME29': [(65 + R[i] + 180) * scale + 2 * gap + box_gap, 30 * scale + 2 * gap + box_gap],
    'FRAME20': [(R[i] + 180) * scale + 2 * gap + box_gap, (FRAME20_H[i] + 50) * scale + 2 * gap + box_gap],
    'FRAME33': [50 * scale + 2 * gap + box_gap, (180 + 50) * scale + 2 * gap + box_gap],
    'FRAME13': [(85+55) * scale + 2 * gap + box_gap , FRAME_13_depth[i] * scale + box_gap + 2 * gap],
    'FRAME23': [99.35 * scale + 2 * gap + box_gap , (150+35) * scale + 2 * gap + box_gap],
    'FRAME24': [(99.35 + 35) * scale + 2 * gap + box_gap, (150 + 35) * scale + 2 * gap + box_gap],
    'FRAME47': [(50 + 19) * scale + 2 * gap + box_gap , 74 * scale + 2 * gap + box_gap],
    'FRAME48': [(32 + 82) * scale + 2 * gap + box_gap , 32 * scale + 2 * gap + box_gap],
    'FRAME49': [(30 + 54) * scale + 2 * gap + box_gap , 30 * scale + 2 * gap + box_gap]
}#正常情況下之零件清單(FRAME24X範圍有額外尺寸)

def scale_Adjustment(i , l):
    scale_p = 0
    while True:
        scale_p += 1
        print(scale_p)
        # if scale_p % 2 != 0 or scale_p % 5 != 0 or scale_p % 10 != 0:  # 比例要為2,5,10的倍數
        #     continue
        if scale_p % 2 == 0:
            pass
        elif scale_p % 5 == 0:
            pass
        elif scale_p % 10 == 0:  # 比例要為2,5,10的倍數
            pass
        else:
            continue
        #----------------------------基本數值-----------------------------------------
        scale = 1 / scale_p
        box_1_range = [box_1_Xmax, box_1_Xmin, box_1_Ymax, box_1_Ymin]
        box_1_rangeX = [box_1_Xmax - box_1_Xmin]#預設範圍
        box_1_rangeY = [box_1_Ymax - box_1_Ymin]#預設範圍
        box_1_center = [(box_1_Xmax + box_1_Xmin) / 2 , (box_1_Ymax + box_1_Ymin) / 2]
        if l ==1:
            w_scale = A[i] * scale
            h_scale = H[i] * scale
            d_scale = B[i] * scale
        else:#1.5寬深之規個
            w_scale = A_15[i] * scale
            h_scale = H[i] * scale
            d_scale = B_15[i] * scale
        # --------------------------------box_1----------------------------
        part_1_width = [2 * d_scale + box_width_gap]
        part_1_heigth = [h_scale + box_heigth_gap]
        if part_1_width > box_1_rangeX:  # 元件寬度判斷
            continue
        elif part_1_heigth > box_1_rangeY:  # 元件高度判斷
            continue
        box_1_range = [round(box_1_Xmax, 2), round(box_1_Xmin, 2), round(box_1_Ymax, 2),
                       round(box_1_Ymin, 2)]  # 方框1的範圍[Xmax , Xmin , Ymax , Ymin]
        break
    return scale , box_1_center , box_1_range

def drafting_parameter_calculation(i, l , scale , box_1_range):
    list = projection_file_name_list#零件清單匯入
    part = 0
    temporary_range = []#X , Y範圍暫存代號
    new_line = ['FRAME30']
    new_line_X = []
    file_list_range = {
                           'FRAME30': [(R[i] + 180 + 40) * scale + 2 * gap + box_gap,
                                        (FRAME44_height[i] + 40) * scale + 2 * gap + box_gap],
                            'FRAME44': [(R[i] + 180) * scale + 2 * gap + box_gap,
                                        (FRAME44_height[i] + 40) * scale + 2 * gap + box_gap],
                            'FRAME41': [(R[i] + 180) * scale + 2 * gap + box_gap,
                                        (FRAME_41_depth[i] + 22) * scale + 2 * gap + box_gap],
                            'FRAME34': [(182 + 40) * scale + 2 * gap + box_gap, (80 + 40) * scale + 2 * gap + box_gap],
                            'FRAME9': [(R[i] + 180) * scale + 2 * gap + box_gap,
                                       (((Z[i] - T[i] - 40) + 50) * scale + 2 * gap + box_gap)],
                            'FRAME45': [(R[i] + 19) * scale + 2 * gap + box_gap, 476 * scale + 2 * gap + box_gap],
                            'FRAME43': [R[i] * scale + 2 * gap + box_gap, 150 * scale + 2 * gap + box_gap / 2],
                            'FRAME7': [FRAME_7_width[i] * scale + 2 * gap + box_gap, 300 * scale + 2 * gap + box_gap / 2],
                            'MAIN_GEAR2': [125 * scale + 2 * gap + box_gap, 270 * scale + 2 * gap + box_gap / 2],
                            'FRAME32': [(R[i] + 19) * scale + 2 * gap + box_gap, 429 * scale + 2 * gap + box_gap / 2],
                            'FRAME31': [(40 + 45) * scale + 2 * gap + box_gap + 50, (150 + 45) * scale + 2 * gap + box_gap / 2],
                            'FRAME38': [290 * scale + 2 * gap + box_gap, 145 * scale + 2 * gap + box_gap / 2],
                            'FRAME11': [FRAME_11_width[i] * scale + 2 * gap + box_gap, (FRAME_11_height[i]+90) * scale + 2 * scale + box_gap],
                            'FRAME39': [145 * scale + 2 * gap + box_gap, (300 + 19) * scale + 2 * gap + box_gap],
                            'FRAME17': [75 * scale + 2 * gap + box_gap, (75 + 48) * scale + 2 * gap + box_gap/2],
                            'FRAME3': [(R[i] + 180) * scale + 2 * gap + box_gap, ((Z[i] - T[i] - 40 + 50) * scale + 2 * gap + box_gap)],
                            'FRAME14': [75 * scale + 2 * gap + box_gap, (75 + 70) * scale + 2 * gap + box_gap],
                            'FRAME22': [460 * scale + 2 * gap + box_gap, 280 * scale + 2 * gap + box_gap/2],
                            'FRAME37': [140 * scale + 2 * gap + box_gap, 80 * scale + 2 * gap + box_gap/2],
                            'FRAME8': [(FRAME_8_width[i]+40) * scale + 2 * gap + box_gap, 300 * scale + 2 * gap + box_gap],
                            'FRAME29': [(65 + R[i] + 180) * scale + 2 * gap + box_gap, 30 * scale + 2 * gap + box_gap],
                            'FRAME20': [(R[i] + 180) * scale + 2 * gap + box_gap, (FRAME20_H[i] + 50) * scale + 2 * gap + box_gap],
                            'FRAME33': [50 * scale + 2 * gap + box_gap, (180 + 50) * scale + 2 * gap + box_gap],
                            'FRAME13': [(85+55) * scale + 2 * gap + box_gap , FRAME_13_depth[i] * scale + box_gap/2 + 2 * gap],
                            'FRAME23': [99.35 * scale + 2 * gap + box_gap, (150+35) * scale + 2 * gap + box_gap],
                            'FRAME24': [(99.35 + 35) * scale + 2 * gap + box_gap, (150 + 35) * scale + 2 * gap + box_gap],
                            'FRAME47': [(50 + 19) * scale + 2 * gap + box_gap , 74 * scale + 2 * gap + box_gap/2],
                            'FRAME48': [(32+82) * scale + 2 * gap + box_gap , 32 * scale + 2 * gap + box_gap/2],
                            'FRAME49': [(30+54) * scale + 2 * gap + box_gap , 30 * scale + 2 * gap + box_gap/2]
                            }
    file_list_range_update:{}
    file_list_15_range = {
                            'FRAME30': [(R_15[i] + 180 + 40) * scale + 2 * gap + box_gap,
                                        (FRAME44_height[i] + 40) * scale + 2 * gap + box_gap],
                            'FRAME44': [(R_15[i] + 180) * scale + 2 * gap + box_gap,
                                        (FRAME44_height[i] + 40) * scale + 2 * gap + box_gap],
                            'FRAME41': [(R_15[i] + 180) * scale + 2 * gap + box_gap,
                                        (FRAME_41_depth[i] + 22) * scale + 2 * gap + box_gap],
                            'FRAME34': [(182 + 40) * scale + 2 * gap + box_gap, (80 + 40) * scale + 2 * gap + box_gap],
                            'FRAME9': [(R_15[i] + 180) * scale + 2 * gap + box_gap,
                                       ((Z[i] - T[i] - 40 + 50) * scale + 2 * gap + box_gap)],
                            'FRAME45': [(R_15[i] + 19) * scale + 2 * gap + box_gap, 476 * scale + 2 * gap + box_gap],
                          'FRAME43': [R_15[i] * scale + 2 * gap + box_gap, 150 * scale + 2 * gap + box_gap],
                          'FRAME7': [FRAME_7_15_width[i] * scale + 2 * gap + box_gap, 300 * scale + 2 * gap + box_gap],
                          'MAIN_GEAR2': [125 * scale + 2 * gap + box_gap, 270 * scale + 2 * gap + box_gap],
                          'FRAME32': [(R_15[i] + 19) * scale + 2 * gap + box_gap, 429 * scale + 2 * gap + box_gap],
                          'FRAME31': [(40 + 45) * scale + 2 * gap + box_gap + 50, (150 + 45) * scale + 2 * gap + box_gap],
                          'FRAME38': [290 * scale + 2 * gap + box_gap, 145 * scale + 2 * gap + box_gap / 2],
                          'FRAME11': [FRAME_11_15_width[i] * scale + 2 * gap + box_gap, (FRAME_11_height[i]+90) * scale + 2 * scale + box_gap],
                          'FRAME39': [145 * scale + 2 * gap + box_gap, (300 + 19) * scale + 2 * gap + box_gap],
                          'FRAME17': [75 * scale + 2 * gap + box_gap, (75 + 48) * scale + 2 * gap + box_gap/2],
                          'FRAME3': [(R_15[i] + 180) * scale + 2 * gap + box_gap, ((Z[i] - T[i] - 40 + 50) * scale + 2 * gap + box_gap)],
                          'FRAME14': [75 * scale + 2 * gap + box_gap, (75 + 70) * scale + 2 * gap + box_gap],
                          'FRAME22': [460 * scale + 2 * gap + box_gap, 280 * scale + 2 * gap + box_gap/2],
                          'FRAME37': [140 * scale + 2 * gap + box_gap, 80 * scale + 2 * gap + box_gap/2],
                          'FRAME8': [(FRAME_8_15_width[i]+40) * scale + 2 * gap + box_gap, 300 * scale + 2 * gap + box_gap],
                          'FRAME29': [(65 + R_15[i] + 180) * scale + 2 * gap + box_gap, 30 * scale + 2 * gap + box_gap],
                          'FRAME20': [(R_15[i] + 180) * scale + 2 * gap + box_gap, (FRAME20_H[i] + 50) * scale + 2 * gap + box_gap],
                          'FRAME33': [50 * scale + 2 * gap + box_gap, (180 + 50) * scale + 2 * gap + box_gap],
                            'FRAME13': [(85 + 55) * scale + 2 * gap + box_gap, FRAME_13_depth[i] * scale + box_gap + 2 * gap/2],
                            'FRAME23': [99.35 * scale + 2 * gap + box_gap, (150 + 35) * scale + 2 * gap + box_gap],
                            'FRAME24': [(99.35 + 35) * scale + 2 * gap + box_gap, (150 + 35) * scale + 2 * gap + box_gap],
                            'FRAME47': [(50 + 19) * scale + 2 * gap + box_gap, 74 * scale + 2 * gap + box_gap/2],
                            'FRAME48': [(32 + 82) * scale + 2 * gap + box_gap, 32 * scale + 2 * gap + box_gap/2],
                            'FRAME49': [(30 + 54) * scale + 2 * gap + box_gap, 30 * scale + 2 * gap + box_gap/2]
                            }
    file_list_15_range_update: {}
    while True:
            if temporary_range == []:#第一次執行
                temporary_range = [box_1_range[1] , box_1_range[2]]#第一方框之Xmin , 第一方框之Ymax
                ALL_range = [box_1_range]#加入範圍清單
                continue
            else:
                while part < 30:
                    if part == 29:
                        break
                    x = list[part]
                    while True:
                        if l ==1:#正常零件規格
                            X_center = temporary_range[0]#X訂為上一方框的Xmin
                            Y_center = temporary_range[1]#Y訂為上一方框Ymax
                            if X_center < ALL_range[0][0] + file_list_range[x][0]/2 + box_gap and len(new_line) == 1:#方框2為始的第一行X方向虛擬方框判斷(X中心+方框一半寬)
                                temporary_range[0] += 1
                                continue
                            if len(new_line) != 1 and X_center < max(new_line_X) + file_list_range[x][0]/2 + box_gap:#後續X方向虛擬方框判斷(X中心+方框一半寬)
                                temporary_range[0] += 1
                                continue
                            if  X_center + file_list_range[x][0]/2 + box_gap > 409 and Y_center + file_list_range[x][1]/2 + box_gap > 821:
                                temporary_range[1] -= 1
                                continue
                            if Y_center + file_list_range[x][1]/2 + box_gap > box_1_range[2] and x == new_line[-1]:#換行後第一個零件的Y方向虛擬方框判斷(Y中心+方框一半高)
                                temporary_range[1] -=1
                                continue
                            if Y_center + file_list_range[x][1]/2 + box_gap > ALL_range[-1][3] and x != new_line[-1]:#後續其他零件的Y方向虛擬方框判斷(Y中心+方框一半高)
                                temporary_range[1] -=1
                                continue
                            if Y_center - file_list_range[x][1] / 2 - box_gap < 25 or X_center + file_list_range[x][0]/2 > 760 and Y_center - file_list_range[x][1] / 2 + box_gap < 45 or X_center + file_list_range[x][0] / 2 > 960 and Y_center - file_list_range[x][1] / 2 - box_gap < 570:#當方框Ymin超出下界限時，Y中心重製為方框1的Ymax，X中心重製為上一行零件的Xmax
                                new_line.append(x)#將超出下限的零件加入換行零件串列
                                temporary_range[1] = box_1_range[2]#Y中心重製為方框1的Ymax
                                for p in range(len(ALL_range)) :#利用for迴圈取出目前方框的Xmax
                                    new_line_X.append(ALL_range[p][0])
                                    new_line.append(x)
                                    p += 1
                                temporary_range[0] = max(new_line_X)#將X中心訂為找出的X最大值
                                continue
                            if X_center + file_list_range[x][0]/2 > 1164:#當X超出圖框範圍時，退出迴圈增加比例
                                part = 0#零件序號重製
                                scale_p = 1 / scale#將比例倒數回整數
                                scale_p = int(scale_p)
                                scale_p += 1
                                if scale_p % 2 != 0 and scale_p % 5 != 0 and scale_p % 10 != 0:  # 比例要為2,5,10的倍數
                                    scale_p += 1
                                print(scale_p)
                                scale = 1 / scale_p#將暫存比例整數倒數回去
                                #將各變數重製
                                temporary_range[0] = box_1_range[1]
                                temporary_range[1] = box_1_range[2]
                                new_line = ['FRAME30']
                                new_line_X.clear()
                                ALL_range = [box_1_range]
                                file_list_range_update = {
                           'FRAME30': [(R[i] + 180 + 40) * scale + 2 * gap + box_gap,
                                        (FRAME44_height[i] + 40) * scale + 2 * gap + box_gap],
                            'FRAME44': [(R[i] + 180) * scale + 2 * gap + box_gap,
                                        (FRAME44_height[i] + 40) * scale + 2 * gap + box_gap],
                            'FRAME41': [(R[i] + 180) * scale + 2 * gap + box_gap,
                                        (FRAME_41_depth[i] + 22) * scale + 2 * gap + box_gap],
                            'FRAME34': [(182 + 40) * scale + 2 * gap + box_gap, (80 + 40) * scale + 2 * gap + box_gap],
                            'FRAME9': [(R[i] + 180) * scale + 2 * gap + box_gap,
                                       (((Z[i] - T[i] - 40) + 50) * scale + 2 * gap + box_gap)],
                            'FRAME45': [(R[i] + 19) * scale + 2 * gap + box_gap, 476 * scale + 2 * gap + box_gap],
                            'FRAME43': [R[i] * scale + 2 * gap + box_gap, 150 * scale + 2 * gap + box_gap / 2],
                            'FRAME7': [FRAME_7_width[i] * scale + 2 * gap + box_gap, 300 * scale + 2 * gap + box_gap / 2],
                            'MAIN_GEAR2': [125 * scale + 2 * gap + box_gap, 270 * scale + 2 * gap + box_gap / 2],
                            'FRAME32': [(R[i] + 19) * scale + 2 * gap + box_gap, 429 * scale + 2 * gap + box_gap / 2],
                            'FRAME31': [(40 + 45) * scale + 2 * gap + box_gap + 50, (150 + 45) * scale + 2 * gap + box_gap / 2],
                            'FRAME38': [290 * scale + 2 * gap + box_gap, 145 * scale + 2 * gap + box_gap / 2],
                            'FRAME11': [FRAME_11_width[i] * scale + 2 * gap + box_gap, (FRAME_11_height[i]+90) * scale + 2 * scale + box_gap],
                            'FRAME39': [145 * scale + 2 * gap + box_gap, (300 + 19) * scale + 2 * gap + box_gap],
                            'FRAME17': [75 * scale + 2 * gap + box_gap, (75 + 48) * scale + 2 * gap + box_gap/2],
                            'FRAME3': [(R[i] + 180) * scale + 2 * gap + box_gap, ((Z[i] - T[i] - 40 + 50) * scale + 2 * gap + box_gap)],
                            'FRAME14': [75 * scale + 2 * gap + box_gap, (75 + 70) * scale + 2 * gap + box_gap],
                            'FRAME22': [460 * scale + 2 * gap + box_gap, 280 * scale + 2 * gap + box_gap/2],
                            'FRAME37': [140 * scale + 2 * gap + box_gap, 80 * scale + 2 * gap + box_gap/2],
                            'FRAME8': [(FRAME_8_width[i]+40) * scale + 2 * gap + box_gap, 300 * scale + 2 * gap + box_gap],
                            'FRAME29': [(65 + R[i] + 180) * scale + 2 * gap + box_gap, 30 * scale + 2 * gap + box_gap],
                            'FRAME20': [(R[i] + 180) * scale + 2 * gap + box_gap, (FRAME20_H[i] + 50) * scale + 2 * gap + box_gap],
                            'FRAME33': [50 * scale + 2 * gap + box_gap, (180 + 50) * scale + 2 * gap + box_gap],
                            'FRAME13': [(85 + 55) * scale + 2 * gap + box_gap,
                                                FRAME_13_depth[i] * scale + box_gap/2 + 2 * gap],
                            'FRAME23': [99.35 * scale + 2 * gap + box_gap,
                                                (150 + 35) * scale + 2 * gap + box_gap],
                            'FRAME24': [(99.35 + 35) * scale + 2 * gap + box_gap,
                                                (150 + 35) * scale + 2 * gap + box_gap],
                            'FRAME47': [(50 + 19) * scale + 2 * gap + box_gap, 74 * scale + 2 * gap + box_gap/2],
                            'FRAME48': [(32 + 82) * scale + 2 * gap + box_gap, 32 * scale + 2 * gap + box_gap/2],
                            'FRAME49': [(30 + 54) * scale + 2 * gap + box_gap, 30 * scale + 2 * gap + box_gap/2]
                            }
                                file_list_range = file_list_range_update
                                break
                            part_range = [round(X_center + file_list_range[x][0]/2 , 2) , round(X_center - file_list_range[x][0]/2 , 2) ,
                                          round(Y_center + file_list_range[x][1]/2 , 2) , round(Y_center - file_list_range[x][1]/2 , 2)]#將計算完成之方框範圍儲存
                            temporary_range[1] = part_range[2]#更換Y中心
                            part += 1
                            ALL_range.append(part_range)#將方框範圍輸入範圍清單
                            break
                        else:#1.5寬深之規格
                            X_center = temporary_range[0]  # X訂為上一方框的Xmin
                            Y_center = temporary_range[1]  # Y訂為上一方框Ymax
                            if X_center < ALL_range[0][0] + file_list_15_range[x][0] / 2 + box_gap and len(
                                    new_line) == 1:  # 方框2為始的第一行X方向虛擬方框判斷(X中心+方框一半寬)
                                temporary_range[0] += 1
                                continue
                            if len(new_line) != 1 and X_center < max(new_line_X) + file_list_15_range[x][
                                0] / 2 + box_gap:  # 後續X方向虛擬方框判斷(X中心+方框一半寬)
                                temporary_range[0] += 1
                                continue
                            if Y_center + file_list_15_range[x][1] / 2 + box_gap > box_1_range[2] and x == new_line[
                                -1]:  # 換行後第一個零件的Y方向虛擬方框判斷(Y中心+方框一半高)
                                temporary_range[1] -= 1
                                continue
                            if Y_center + file_list_15_range[x][1] / 2 + box_gap > ALL_range[-1][3] and x != new_line[
                                -1]:  # 後續其他零件的Y方向虛擬方框判斷(Y中心+方框一半高)
                                temporary_range[1] -= 1
                                continue
                            if Y_center - file_list_15_range[x][1] / 2 - box_gap < 25 or X_center + file_list_15_range[x][0]/2 > 760 and Y_center - file_list_15_range[x][1] / 2 + box_gap < 45 or X_center + file_list_range[x][0] / 2 > 960 and Y_center - file_list_range[x][1] / 2 - box_gap < 570:#當方框Ymin超出下界限時，Y中心重製為方框1的Ymax，X中心重製為上一行零件的Xmax
                                new_line.append(x)#將超出下限的零件加入換行零件串列
                                temporary_range[1] = box_1_range[2]#Y中心重製為方框1的Ymax
                                for p in range(len(ALL_range)) :#利用for迴圈取出目前方框的Xmax
                                    new_line_X.append(ALL_range[p][0])
                                    new_line.append(x)
                                    p += 1
                                temporary_range[0] = max(new_line_X)#將X中心訂為找出的X最大值
                                continue
                            if X_center + file_list_15_range[x][0]/2 > 1164:#當X超出圖框範圍時，退出迴圈增加比例
                                part = 0  # 零件序號重製
                                scale_p = 1 / scale  # 將比例倒數回整數
                                scale_p = int(scale_p)
                                scale_p += 1
                                if scale_p % 2 != 0 and scale_p % 5 != 0 and scale_p % 10 != 0:  # 比例要為2,5,10的倍數
                                    scale_p += 1
                                print(scale_p)
                                scale = 1 / scale_p# 將暫存比例整數倒數回去
                                temporary_range[0] = box_1_range[1]
                                temporary_range[1] = box_1_range[2]
                                new_line = ['FRAME30']
                                new_line_X = []
                                ALL_range = [box_1_range]
                                file_list_15_range_update = {
                            'FRAME30': [(R_15[i] + 180 + 40) * scale + 2 * gap + box_gap,
                                        (FRAME44_height[i] + 40) * scale + 2 * gap + box_gap],
                            'FRAME44': [(R_15[i] + 180) * scale + 2 * gap + box_gap,
                                        (FRAME44_height[i] + 40) * scale + 2 * gap + box_gap],
                            'FRAME41': [(R_15[i] + 180) * scale + 2 * gap + box_gap,
                                        (FRAME_41_depth[i] + 22) * scale + 2 * gap + box_gap],
                            'FRAME34': [(182 + 40) * scale + 2 * gap + box_gap, (80 + 40) * scale + 2 * gap + box_gap],
                            'FRAME9': [(R_15[i] + 180) * scale + 2 * gap + box_gap,
                                       ((Z[i] - T[i] - 40 + 50) * scale + 2 * gap + box_gap)],
                            'FRAME45': [(R_15[i] + 19) * scale + 2 * gap + box_gap, 476 * scale + 2 * gap + box_gap],
                          'FRAME43': [R_15[i] * scale + 2 * gap + box_gap, 150 * scale + 2 * gap + box_gap],
                          'FRAME7': [FRAME_7_15_width[i] * scale + 2 * gap + box_gap, 300 * scale + 2 * gap + box_gap],
                          'MAIN_GEAR2': [125 * scale + 2 * gap + box_gap, 270 * scale + 2 * gap + box_gap],
                          'FRAME32': [(R_15[i] + 19) * scale + 2 * gap + box_gap, 429 * scale + 2 * gap + box_gap],
                          'FRAME31': [(40 + 45) * scale + 2 * gap + box_gap + 50, (150 + 45) * scale + 2 * gap + box_gap],
                          'FRAME38': [290 * scale + 2 * gap + box_gap, 145 * scale + 2 * gap + box_gap / 2],
                          'FRAME11': [FRAME_11_15_width[i] * scale + 2 * gap + box_gap, (FRAME_11_height[i]+90) * scale + 2 * scale + box_gap],
                          'FRAME39': [145 * scale + 2 * gap + box_gap, (300 + 19) * scale + 2 * gap + box_gap],
                          'FRAME17': [75 * scale + 2 * gap + box_gap, (75 + 48) * scale + 2 * gap + box_gap/2],
                          'FRAME3': [(R_15[i] + 180) * scale + 2 * gap + box_gap, ((Z[i] - T[i] - 40 + 50) * scale + 2 * gap + box_gap)],
                          'FRAME14': [75 * scale + 2 * gap + box_gap, (75 + 70) * scale + 2 * gap + box_gap],
                          'FRAME22': [460 * scale + 2 * gap + box_gap, 280 * scale + 2 * gap + box_gap/2],
                          'FRAME37': [140 * scale + 2 * gap + box_gap, 80 * scale + 2 * gap + box_gap/2],
                          'FRAME8': [(FRAME_8_15_width[i]+40) * scale + 2 * gap + box_gap, 300 * scale + 2 * gap + box_gap],
                          'FRAME29': [(65 + R_15[i] + 180) * scale + 2 * gap + box_gap, 30 * scale + 2 * gap + box_gap],
                          'FRAME20': [(R_15[i] + 180) * scale + 2 * gap + box_gap, (FRAME20_H[i] + 50) * scale + 2 * gap + box_gap],
                          'FRAME33': [50 * scale + 2 * gap + box_gap, (180 + 50) * scale + 2 * gap + box_gap],
                        'FRAME13': [(85 + 55) * scale + 2 * gap + box_gap,
                                                FRAME_13_depth[i] * scale + box_gap/2 + 2 * gap],
                        'FRAME23': [99.35 * scale + 2 * gap + box_gap,
                                                (150 + 35) * scale + 2 * gap + box_gap],
                        'FRAME24': [(99.35 + 35) * scale + 2 * gap + box_gap,
                                                (150 + 35) * scale + 2 * gap + box_gap],
                        'FRAME47': [(50 + 19) * scale + 2 * gap + box_gap, 74 * scale + 2 * gap + box_gap/2],
                        'FRAME48': [(32 + 82) * scale + 2 * gap + box_gap, 32 * scale + 2 * gap + box_gap/2],
                        'FRAME49': [(30 + 54) * scale + 2 * gap + box_gap, 30 * scale + 2 * gap + box_gap/2]
                            }
                                file_list_15_range = file_list_15_range_update
                                break
                            part_range = [round(X_center + file_list_15_range[x][0] / 2, 2),
                                          round(X_center - file_list_15_range[x][0] / 2, 2),
                                          round(Y_center + file_list_15_range[x][1] / 2, 2),
                                          round(Y_center - file_list_15_range[x][1] / 2, 2)]  # 將計算完成之方框範圍儲存
                            temporary_range[1] = part_range[2]  # 更換Y中心
                            part += 1
                            ALL_range.append(part_range)  # 將方框範圍輸入範圍清單
                            break
            break
    return ALL_range , scale

def Material_diagram_projection(surface , XLocation , YLocation , part_name , scale , part_view_number , new_part):#零件圖佈圖主程式
    catapp = win32.Dispatch('CATIA.Application')
    ActWin = catapp.Windows.item("detail_drawing.CATDrawing")
    ActWin.Activate()
    drawingDocument = catapp.ActiveDocument
    drawingSheets = drawingDocument.Sheets
    drawingSheet = drawingSheets.Item("Sheet.1")
    drawingViews1 = drawingSheet.Views
    drawingView1 = drawingViews1.Add(part_name + "_" + part_view_number)
    drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    documents1 = catapp.Documents
    partDocument1 = documents1.Item(part_name + ".CATPart")
    # partDocument1 = documents1.Item('FRAME1' + ".CATPart")
    product1 = partDocument1.GetItem(part_name)
    drawingViewGenerativeBehavior1.Document = product1
    drawingViewGenerativeBehavior1.DefineFrontView(projection[U[surface]][0], projection[U[surface]][1]
                                                           , projection[U[surface]][2], projection[U[surface]][3]
                                                           , projection[U[surface]][4], projection[U[surface]][5])
    drawingView1.X = XLocation
    drawingView1.Y = YLocation
    drawingView1.Scale = scale
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    drawingViewGenerativeBehavior1.update()
    drawingView1.Activate()
    # print(x + "_" + part_view_number)
    # drawingview2 = drawingViews1.Item(x + "_" + part_view_number)
    # drawingtexts1 = drawingview2.Texts
    # drawingtext1 = drawingtexts1.Item(1)
    # # drawingtexts1 = drawingtext1.Parent
    # selection = partDocument1.Selection
    # selection.Add(drawingtext1)
    # selection.Delete()
    selection = drawingDocument.Selection
    selection.Clear()
    # --
    try:
        drawingview1 =  drawingViews1.Item(part_name + "_" + part_view_number)
        # drawingview1 =  drawingViews1.Item('FRAME1_1')
        drawingtexts1 = drawingview1.Texts
        drawingtext1 = drawingtexts1.Item(1)
        drawingtexts1 = drawingtext1.Parent
        selection = drawingDocument.Selection
        selection.Add(drawingtext1)
        selection.Delete()
        selection.Clear()
    except:
        drawingview1 = drawingViews1.Item(part_name + "_" + part_view_number)
        # drawingview1 =  drawingViews1.Item('FRAME1_1')
        drawingtexts1 = drawingview1.Texts
        drawingtext1 = drawingtexts1.Item(1)
        drawingtexts1 = drawingtext1.Parent
        selection = drawingDocument.Selection
        selection.Add(drawingtext1)
        selection.Delete()
        selection.Clear()
    # drawingViews1.update()
    # drawingView1.FrameVisualization = False


def Parts_drafting_balloons(view , name , XLocation , YLocation , part_view_number , scale):#圈碼圖建立
        catapp = win32.Dispatch("CATIA.Application")
        partdoc = catapp.ActiveDocument
        catapp = win32.Dispatch('CATIA.Application')
        drawingdocument = catapp.ActiveDocument
        drawingsheets = drawingdocument.Sheets
        drawingsheet = drawingsheets.Item('Sheet.1')
        drawingviews = drawingsheet.Views
        drawingview = drawingviews.Item(view + "_" + part_view_number)
        drawingview.Activate()
        DrawTexts_balloons = drawingview.Texts
        DrawText = DrawTexts_balloons.Add(name, XLocation , YLocation)
        DrawText.SetFontName(0, 0, 'Arial Unicode MS (TrueType)')
        DrawText.SetFontSize(0, 0, 8.4)  # 調整字體位置和大小(x, y, 字體大小)
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
BOM_list = {'1': ['01 ', '50', '1680x2910', 'SS400', '1', '1628.6', '如圖'],
            '2': ['02 ', '50', '1680x2910', 'SS400', '1', '1628.6', '如圖'],
            '3': ['03 ', '40', '850x1027', 'SS400', '1', '1626.6', '如圖'],
            '4': ['04 ', '32', '850x1050', 'SS400', '1', '207.0', '如圖'],
            '5': ['05 ', '22', '850x590', 'SS400', '1', '84.2', '如圖'],
            '6': ['06 ', '40', '80x182', 'SS400', '2', '3.9', '如圖'],
            '7': ['07 ', '50', '850x710', 'SS400', '1', '127.0', '如圖'],
            '8': ['08 ', '22', '670x476', 'SS400', '1', '53.8', '如圖'],
            '9': ['8-1', '22', '670x150', 'SS400', '2', '17.4', '如圖'],
            '10': [' 9 ', ' 9 ', '協中:250x90x850(待加)L', 'SS400', '1', '28.5', '槽型鋼'],
            '11': [' 9 ', ' 9 ', '協台:250x90x850(待加)L', 'SS400', '1', '29.4', '槽型鋼'],
            '12': ['10 ', '45', '255x300', 'SS400', '2', '24.5', '如圖'],
            '13': ['11 ', '125', 'φ270xφ170', 'SS400', '1', '33.9', '如圖'],
            '14': ['12 ', '19', '670x429', 'SS400', '1', '42.0', '如圖'],
            '15': ['13 ', '19', '48x74', 'SS400', '2', '0.4', '如圖'],
            '16': ['14 ', '40', '105x150', 'SS400', '2', '4.9', '如圖'],
            '17': ['15A ', '50', '40x150', 'SS400', '2', '2.1', '如圖'],
            '18': ['15 ', '50', '40x150', 'SS400', '2', '2.3', '如圖'],
            '19': ['16 ', '40', '105x150', 'SS400', '2', '4.8', '如圖'],
            '20': ['18 ', '19', '145x290', 'SS400', '1', '6.2', '如圖'],
            '21': ['19 ', '90', '960x1588', 'SS400', '2', '609.4', '如圖'],
            '22': ['20 ', '19', '145x300', 'SS400', '1', '5.8', '如圖'],
            '23': ['21 ', '53', '75x75', 'SS400', '1', '2.4', '如圖'],
            '24': ['22 ', '50', '850x710', 'SS400', '1', '156.4', '如圖'],
            '25': ['24 ', '60', '85x420', 'SS400', '2', '15.9', '如圖'],
            '26': ['25 ', '75', '75x75', 'SS400', '5', '2.4', '如圖'],
            '27': ['26 ', '35', '50x50', 'SS400', '2', '0.7', ''],
            '28': ['27 ', '25', '460x280', 'SS400', '2', '25.2', '如圖'],
            '29': ['29 ', '19', '140x80', 'SS400', '4', '1.2', '如圖'],
            '30': ['30 ', '45', '243x300', 'SS400', '2', '0.9', ''],
            '31': ['33 ', '6', 'L65x30x850', 'SS400', '1', '3.6', '如圖'],
            '32': ['34 ', '60', 'φ50', 'SS400', '2', '0.9', ''],
            '33': ['35 ', '50', '850x680', 'SS400', '1', '197.2', '如圖'],
            '34': ['40 ', '22', '850x165', 'SS400', '1', '24.2', ''],
            '35': ['41 ', '6','角鐵:50x50x180L', 'SS400', '1', '197.2', '如圖'],
            '36': ['43 ', '32', '32x85', 'SS400', '1', '0.7', '如圖'],
            '37': ['44 ', '62', 'φ30', 'SS400', '1', '0.3', '如圖']}


BOM_list_name = []
for i in BOM_list:
    BOM_list_name.append(i)

BOM_position = [918, 114]
BOM_position_x = [865 + 99, 865 + 116, 865 + 130, 865 + 212.5, 865 + 242, 865 + 255, 865 + 280]
BOM_position_Y = 12

def bom_text_create():
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingview = drawingsheet.Views.Item('Background View')
    drawingview = drawingview.Texts
    # drawingview.Activate()
    for i in BOM_list_name:
        for j in range(0, 7):
            DrawText = drawingview.Add(BOM_list[i][j], BOM_position_x[j], 120 + 12 * int(i)) # 線段長度
            DrawText.SetFontName(0, 0, 'Arial Unicode MS (TrueType)')
            DrawText.SetFontSize(0, 0, 5)  # 調整字體位置和大小(x, y, 字體大小).
            DrawText.AnchorPosition = 2  # 文字框對其位置(TopL=1,MidL=2,BottomL=3,TopM=4,MidM=5,BottomM=6,TopR=7,MidR=8,BottomR=9)

# bom_text_create()

def create_center_line(x_value_1, y_value_1, x_value_2, y_value_2):
    catapp = win32.Dispatch('CATIA.Application')
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingviews = drawingsheet.Views
    drawingview = drawingviews.Item('Background View')
    drawingview.Activate()
    factory2d = drawingview.Factory2D
    dotted_line = factory2d.CreateLine(x_value_1, y_value_1, x_value_2, y_value_2)
    dotted_line.Name = 'chain_line'
    object = catapp.ActiveDocument.Selection
    # object.Search('Name=chain_line,all')
    vis = catapp.ActiveDocument.Selection.VisProperties
    vis.SetRealLineType(2, 0.3)
    vis.SetRealWidth(1, 0.13)

# for i in BOM_list_name:
#     create_center_line(960, 102 + 12 * int(i), 1169, 102 + 12 * int(i))


