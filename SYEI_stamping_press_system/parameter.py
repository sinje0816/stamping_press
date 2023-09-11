import math

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
FRAME20_H = [288, 363, 428, 445.5, 558, 688, 833, 998, 1128, 1178]
FRAME2_lower_depth = [166.016, 246.016, 286.016, 366.016, 446.016, 526.016, 606.016, 686.016, 746.016, 806.016]
FRAME2_lower_depth_15 = [331.016, 451.016, 511.016, 631.016, 751.016, 871.016, 991.016, 1111.016, 1201.016, 1291.016]
FRAME1_lower_high = [1330, 1335, 1340, 1445, 1455, 1460, 1470, 1575, 1595, 1695]
FRAME20_FRAME2_YZ = [805, 805, 979, 979, 979, 979, 979, 979, 979, 979]
BALANCER1_XZ = [204, 253, 268, 282, 317, 345, 375, 460, 575, 690]  # (R+180mm)/2+80mm
FRAME_10_H = [558, 620.5, 673, 798, 905.5, 1023, 1163, 1320.5, 1530.5, 1740.5]
FRAME_32_XY = [0, 0, 0, 1722, 1819.5, 1922, 2052, 2294.5, 2504.5, 2794.5]
FRAME44_height = [588, 665.5, 733, 773, 890.5, 1023, 1173, 1385.5, 1455.5, 1445.5]
FRAME_41_depth = [590, 590, 590, 590, 490, 590, 590, 590, 590, 590]
FRAME_7_width = [176, 182, 197, 208, 228, 255, 295, 299, 305, 305]
FRAME_7_15_width = [200.3, 209.3, 231.8, 248.3, 278.3, 318.8, 378.8, 384.8, 393.8, 393.8]
FRAME_11_height = [1122, 1184.5, 1237, 1362, 1469.5, 1587, 1727, 1884.5, 2094.5, 2304.5]
FRAME_11_width = [600, 680, 720, 800, 880, 960, 1040, 1120, 1180, 1240]
FRAME_11_15_width = [765, 885, 945, 1065, 1185, 1305, 1425, 1545, 1635, 1725]
FRAME_8_width = [164, 170, 185, 196, 216, 243, 283, 288, 293, 358]
FRAME_8_15_width = [246, 255, 278, 294, 324, 365, 425, 432, 440, 537]
FRAME_13_depth = [59.016, 139.016, 179.016, 259.016, 339.016, 419.016, 499.016, 579.016, 639.016, 699.016]
FRAME_6_7_width = [48.5, 54.5, 69.5, 80.5, 100.5, 127.5, 167.5, 171.5, 177.5, 183.5]
pocket_1_upper_hole = [370, 415, 450, 577.5, 670, 770, 895, 1080, 1210, 1340]
FRAME_10_11_center_to_Y_1 = ['', '', '', '', -560.5, -552, -600, -715.5, -719.5, -614]
FRAME_10_11_center_to_Y_2 = ['', '', '', '', -587.5, -583, -650 + 19, -765.5 + 19, -739.5, -644.5]

# 三角函數
sin45 = math.sin(math.radians(45))
cos45 = math.cos(math.radians(45))
sin30 = math.sin(math.radians(30))
cos30 = math.cos(math.radians(30))
sin60 = math.sin(math.radians(60))
cos60 = math.cos(math.radians(60))
cos35_267 = math.cos(math.radians(35.267))
sin35_267 = math.sin(math.radians(35.267))

# -------------開啟零件檔-------------
file_name_list = ['BOLSTER1', 'BOLSTER2', 'BOLSTER3', 'Fixture', 'FRAME1', 'FRAME2', 'FRAME3', 'FRAME4',
                  'FRAME5', 'FRAME6', 'FRAME7', 'FRAME8', 'FRAME9', 'FRAME10', 'FRAME11', 'FRAME12', 'FRAME13',
                  'FRAME14', 'FRAME15', 'FRAME16', 'FRAME17', 'FRAME18', 'FRAME19', 'FRAME20', 'FRAME21',
                  'FRAME22', 'FRAME23', 'FRAME24', 'FRAME25', 'FRAME26', 'FRAME27', 'FRAME28', 'FRAME29',
                  'FRAME30', 'FRAME31', 'FRAME32', 'FRAME33', 'FRAME34', 'FRAME35', 'FRAME36', 'FRAME36',
                  'FRAME37', 'FRAME38', 'FRAME39', 'FRAME40', 'FRAME41', 'FRAME42', 'FRAME43', 'FRAME44',
                  'FRAME45', 'FRAME46', 'GIB1', 'GIB2', 'BALANCER_LEFT_All', 'BALANCER_RIGHT_ALL',
                  'CRANK_SHAFT_CLOCK',
                  'CLUCTH_ASSEMBLY_All', 'SLIDE_UNIT_All', 'CRANK_SHAFT', 'JOINT_All', 'MAIN_GEAR1',
                  'MAIN_GEAR2', 'MAIN_GEAR3', 'MAIN_GEAR4', 'JOINT1', 'FRAME47', 'FRAME48', 'FRAME49',
                  'FRAME50', 'FRAME51', 'FRAME52', 'MOTOR',
                  ]

# -------------焊接圖---------------
#隱藏part
hide_part_name = ['BOLSTER1', 'BOLSTER2', 'BOLSTER3', 'BALANCER_LEFT_All', 'BALANCER_RIGHT_All', 'CRANK_SHAFT_CLOCK',
                  'CLUCTH_ASSEMBLY_All', 'SLIDE_UNIT_All', 'CRANK_SHAFT.1', 'JOINT_All', 'MAIN_GEAR1',
                  'MAIN_GEAR3', 'MAIN_GEAR4', 'JOINT1', 'GIB1', 'GIB2', 'FRAME35', 'FRAME40', 'Fixture']
part_name_Section_view_F_F = ['FRAME1', 'FRAME10']
part_name_Section_view_D_D = ['FRAME21', 'FRAME22', 'FRAME32']
# 前視圖中心 & 上視圖中心
drafting_front_area_centerX = 435
# 下圖
drafting_down_area_centerY = 460 / 2 + 65
# 上圖
drafting_up_area_centerY = (821 - 65 - 460) / 2 + 65 + 420
# 圖面最大範圍框
drafting_view_min_Y = 43
drafting_view_max_Y = 842
drafting_view_min_X = 20
drafting_view_max_X = 1169
# 圖框虛線間隔
draft_X_clearence = 10
draft_Y_clearence = 25
# 下圖Y最小位置
drafting_down_min_Y = 65
# 焊接符號標註文字
# 註:第一格無值
drafting_Welding_text = {'A-A Top': ['', '15', '', '6處\n頂面焊道要用砂輪磨平', '', '', '', '', '', '', '35°'],
                         'test': ['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11']}
# 圈碼座標


# -------------平板相關參數---------------
# T型槽相關參數
total_pierce = []
total_clearance_hole = []
total_t_direction = []
total_t_dimension = []
total_unpierce = []
# T型槽外型尺寸
t_all_dimension = []
t_all_dimension = [22, 38, 23, 16]
t_all_dimension_name = ['tw2', 'tw1', 'th2', 'th1']
# 平板外型尺寸
plate_dimension = []
# 除料孔(待補齊)
plate_hole_type = []
plate_hole_position = []
plate_hole_dimension = []
# 平板變數大全
plate_all_parameter = {}
plate_parameter_name = []
plate_parameter_value = []