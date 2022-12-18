import math
import os
import win32com.client as win32
import datetime, time
import main_program as mprog
import drafting as draft

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

type = str('SN1-110')
i = 6
#
# #
mprog.switch_window()
mprog.constaint_value_change(167, 0, 0)
# --------------------------------------- 生成爆炸圖--------------------------------------------
mprog.constaint_value_change(159, 312.5, 0)  # 支架
mprog.constaint_value_change(186, 510, 1)  # 離合器
mprog.constaint_value_change(203, 1010, 1)  # JOINT_All.1
mprog.constaint_value_change(195, 592.5 + 150, 1)  # 大齒輪MAIN_GEA1
mprog.constaint_value_change(201, -84, 1)  # Joint1

# 重新定義拘束尺寸
BOLSTER1_Offset_value = 1750
GIB_Offset_value = -3000
Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
CLOCK_Offset_value = 2850
CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
BOLSTER1_XY_Z = -T[i]
BOLSTER1_XY_T = Z[i] - T[i] + BOLSTER1_XY_Z
BOLSTER1_XY = BOLSTER1_XY_T + DH_S[i]
BALANCER = -1700

mprog.constaint_value_change(3, -B[i] - BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(6, -B[i] - BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(36, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
mprog.constaint_value_change(39, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
mprog.constaint_value_change(63, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(66, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(168, BOLSTER1_XY, 0)

mprog.update()  # 更新
mprog.switch_window()
drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale = draft.drafting_parameter_calculation(A[i], B[i], H[i], T[i])  # 計算爆炸圖比例及位置
draft.change_Drawing_scale(1 / scale)  # 圖面比例
draft.exploded_Drawing_1(type, drafting_isometric_Coordinate_Position['exploded_1'][0], drafting_isometric_Coordinate_Position['exploded_1'][1], scale)  # 爆炸圖1
mprog.switch_window()  # 開啟3D圖視窗
# 還原零件初始位置
BOLSTER1_Offset_value = 80 - F[i] / 2
GIB_Offset_value = 334.5 - 45
CLOCK_Offset_value = 35
CLOCK_SHAFT_Offset_value = 45
Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
BOLSTER1_XY_Z = - Z[i]
BOLSTER1_XY_T = -T[i]
BOLSTER1_XY = DH_S[i]
BALANCER = -32

mprog.constaint_value_change(3, -B[i] - BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(6, -B[i] - BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(36, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
mprog.constaint_value_change(39, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
mprog.constaint_value_change(63, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(66, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
mprog.constaint_value_change(168, BOLSTER1_XY, 0)
mprog.update()
# # --------------------爆炸圖右圖------------------
Fixture_offset_value = 2850
CLUCTH_ASSEMBLY_offset_value = B[i] + 1650
JOINT_ALL_offset_value = CLUCTH_ASSEMBLY_offset_value
MAIN_GEAR_offset_value = B[i] + 850

mprog.constaint_value_change(159, -312.5 - Fixture_offset_value, 0)  # 支架
mprog.constaint_value_change(186, CLUCTH_ASSEMBLY_offset_value, 1)  # 離合器
mprog.constaint_value_change(203, JOINT_ALL_offset_value, 1)  # JOINT_All.1
mprog.constaint_value_change(195, MAIN_GEAR_offset_value, 1)  # 大齒輪MAIN_GEA1
mprog.constaint_value_change(201, -375, 1)  # Joint1
mprog.update()
mprog.switch_window()
draft.exploded_Drawing_2(type, drafting_isometric_Coordinate_Position['exploded_2'][0], drafting_isometric_Coordinate_Position['exploded_2'][1], scale)
mprog.switch_window()
# 復原位置
mprog.constaint_value_change(159, 312.5, 0)  # 支架
mprog.constaint_value_change(186, 510, 1)  # 離合器
mprog.constaint_value_change(203, 1010, 1)  # JOINT_All.1
mprog.constaint_value_change(195, 592.5 + 150, 1)  # 大齒輪MAIN_GEA1
mprog.constaint_value_change(201, -84, 1)  # Joint1
mprog.update()
mprog.switch_window()
# # --------------------爆炸圖下(前、左、右)----------------
draft.Front_View_Drawing(type, drafting_Coordinate_Position['Front View'][0],
                         drafting_Coordinate_Position['Front View'][1], scale)
draft.Left_View_Drawing(type, drafting_Coordinate_Position['Left View'][0],
                        drafting_Coordinate_Position['Left View'][1], scale)
draft.Right_View_Drawing(type, drafting_Coordinate_Position['Right View'][0],
                         drafting_Coordinate_Position['Right View'][1], scale)

# ------------圈碼圖------------
sin45 = math.sin(math.radians(45))
cos45 = math.cos(math.radians(45))
sin30 = math.sin(math.radians(30))
cos30 = math.cos(math.radians(30))
sin60 = math.sin(math.radians(60))
cos60 = math.cos(math.radians(60))

BOLSTER1_Offset_value = 1750
BLOSTER1_center_Offset_Value = 1775
GIB_Offset_value = -3000
Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 +109.85
CLOCK_Offset_value = 2850
CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
BOLSTER1_XY_Z = -T[i]
BOLSTER1_XY_T = Z[i] - T[i] + BOLSTER1_XY_Z
BOLSTER1_XY = BOLSTER1_XY_T + DH_S[i]

balloons_posotion = [750 * cos30, 750 * sin30]
CLOCK_Pointer = 3325
FRAME_TOP_LEFT_X = (R[i] + 90 + 50) * cos45 + 50
FRAME_TOP_LEFT_X_1 = FRAME_TOP_LEFT_X * sin30 / cos30


FRAME_TOP_LEFT_X_Y = H[i] - S[i] - Z[i] - 12  # 因X產生微小高度
FRAME_TOP_LEFT_Y = FRAME_TOP_LEFT_X * math.sin(math.radians(57.24)) / math.cos(math.radians(57.24)) + 112

scale = 1 / scale

# (將3D圖面尺寸轉換為2D尺寸, 將圖面x座標尺寸轉為R再將R轉為Y)
point_position = {'2': [-CLOCK_Offset_value * cos45,
                        -CLOCK_Offset_value * cos45 * sin30 / cos30],
                  '3': [(-1250) * cos45,
                        (-1250) * cos45 * sin30 / cos30],
                  '4': [(-BOLSTER1_Offset_value - R[i] / 2) * cos45,
                        (-BOLSTER1_Offset_value - R[i] / 2) * cos45 * sin30 / cos30],
                  '5': [(-BOLSTER1_Offset_value - R[i]) * cos45,
                        (-BOLSTER1_Offset_value - R[i]) * cos45 * sin30 / cos30 - S[i] + T[i]],
                  '6': [(GIB_Offset_value - (R[i] + 145) / 2) * cos45,
                        (GIB_Offset_value - (R[i] + 145) / 2) * cos45 * sin30 / cos30]}

circle_position = {
    '2': ['2', point_position['2'][0] - FRAME_TOP_LEFT_X, point_position['2'][1] + FRAME_TOP_LEFT_X_1],
    '3': ['3', point_position['3'][0] - FRAME_TOP_LEFT_X, point_position['3'][1] + FRAME_TOP_LEFT_X_1],
    '4': ['4', point_position['4'][0] - FRAME_TOP_LEFT_X, point_position['4'][1] + FRAME_TOP_LEFT_X_1],
    '5': ['5', point_position['5'][0] * 2 * 0.9 - FRAME_TOP_LEFT_X, point_position['5'][0] * 2 * 0.9 * sin30 / cos30 + FRAME_TOP_LEFT_X_1],
    '6': ['6', point_position['6'][0] - FRAME_TOP_LEFT_X, point_position['6'][1] + FRAME_TOP_LEFT_X_1]}

draft.balloons('Isometric view1', circle_position['6'][0], circle_position['6'][1], circle_position['6'][2],
               point_position['6'][0], point_position['6'][1], circle_position['6'][1])
draft.balloons('Isometric view1', circle_position['2'][0], circle_position['2'][1], circle_position['2'][2],
               point_position['2'][0], point_position['2'][1], circle_position['2'][1])
draft.balloons('Isometric view1', circle_position['4'][0], circle_position['4'][1], circle_position['4'][2],
               point_position['4'][0], point_position['4'][1], circle_position['4'][1])
draft.balloons('Isometric view1', circle_position['5'][0], circle_position['5'][1], circle_position['5'][2],
               point_position['5'][0], point_position['5'][1], circle_position['5'][1])
draft.balloons('Isometric view1', circle_position['3'][0], circle_position['3'][1], circle_position['3'][2],
               point_position['3'][0], point_position['3'][1], circle_position['3'][1])

# ------------中心線-------------
draft.create_center_line('Isometric view1', 0, 0, -CLOCK_Pointer * cos45, -CLOCK_Pointer * cos45 * sin30 / cos30)
draft.create_center_line('Isometric view1', -BLOSTER1_center_Offset_Value * cos45, -BLOSTER1_center_Offset_Value * cos45 * sin30 / cos30,
                         -BLOSTER1_center_Offset_Value * cos45, -S[i] - Z[i] - T[i])
draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30, -JOINT_ALL_offset_value * cos45, -JOINT_ALL_offset_value * cos45 * sin30 / cos30 - 50)
draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30 - 300,
                         -(Fixture_offset_value + B[i]) * cos45, -(Fixture_offset_value + B[i]) * cos45 * sin30 / cos30 - 300)



