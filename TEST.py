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

i = 5
# # 計算爆炸圖比例及位置
# drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale = draft.drafting_parameter_calculation(A[i],
#                                                                                                                    B[i],
#                                                                                                                    H[i],
#                                                                                                                    E[i])
#                                                                                                                    T[i])
#
# sin45 = math.sin(math.radians(45))
# cos45 = math.cos(math.radians(45))
# sin30 = math.sin(math.radians(30))
# cos30 = math.cos(math.radians(30))
# sin60 = math.sin(math.radians(60))
# cos60 = math.cos(math.radians(60))
#
# BOLSTER1_Offset_value = 1750
# BLOSTER1_center_Offset_Value = 1775
# GIB_Offset_value = -3000
# Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 +109.85
# CLOCK_Offset_value = 2850
# CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
# BOLSTER1_XY_Z = -T[i]
# BOLSTER1_XY_T = Z[i] - T[i] + BOLSTER1_XY_Z
# BOLSTER1_XY = BOLSTER1_XY_T + DH_S[i]
#
# balloons_posotion = [750 * cos30, 750 * sin30]
# CLOCK_Pointer = 3325
# FRAME_TOP_LEFT_X = (R[i] + 90 + 50) * cos45 + 50
# FRAME_TOP_LEFT_X_1 = FRAME_TOP_LEFT_X * sin30 / cos30
#
#
# FRAME_TOP_LEFT_X_Y = H[i] - S[i] - Z[i] - 12  # 因X產生微小高度
# FRAME_TOP_LEFT_Y = FRAME_TOP_LEFT_X * math.sin(math.radians(57.24)) / math.cos(math.radians(57.24)) + 112
#
# scale = 1 / scale
#
# # (將3D圖面尺寸轉換為2D尺寸, 將圖面x座標尺寸轉為R再將R轉為Y)
# point_position = {'2': [-CLOCK_Offset_value * cos45,
#                         -CLOCK_Offset_value * cos45 * sin30 / cos30],
#                   '3': [(-1250) * cos45,
#                         (-1250) * cos45 * sin30 / cos30],
#                   '4': [(-BOLSTER1_Offset_value - R[i] / 2) * cos45,
#                         (-BOLSTER1_Offset_value - R[i] / 2) * cos45 * sin30 / cos30],
#                   '5': [(-BOLSTER1_Offset_value - R[i]) * cos45,
#                         (-BOLSTER1_Offset_value - R[i]) * cos45 * sin30 / cos30 - S[i] + T[i]],
#                   '6': [(GIB_Offset_value - (R[i] + 145) / 2) * cos45,
#                         (GIB_Offset_value - (R[i] + 145) / 2) * cos45 * sin30 / cos30]}
#
# circle_position = {
#     '2': ['2', point_position['2'][0] - FRAME_TOP_LEFT_X, point_position['2'][1] + FRAME_TOP_LEFT_X_1],
#     '3': ['3', point_position['3'][0] - FRAME_TOP_LEFT_X, point_position['3'][1] + FRAME_TOP_LEFT_X_1],
#     '4': ['4', point_position['4'][0] - FRAME_TOP_LEFT_X, point_position['4'][1] + FRAME_TOP_LEFT_X_1],
#     '5': ['5', point_position['5'][0] * 2 * 0.9 - FRAME_TOP_LEFT_X, point_position['5'][0] * 2 * 0.9 * sin30 / cos30 + FRAME_TOP_LEFT_X_1],
#     '6': ['6', point_position['6'][0] - FRAME_TOP_LEFT_X, point_position['6'][1] + FRAME_TOP_LEFT_X_1]}
#
# draft.balloons('Isometric view', circle_position['6'][0], circle_position['6'][1], circle_position['6'][2],
#                point_position['6'][0], point_position['6'][1], circle_position['6'][1])
# draft.balloons('Isometric view', circle_position['2'][0], circle_position['2'][1], circle_position['2'][2],
#                point_position['2'][0], point_position['2'][1], circle_position['2'][1])
# draft.balloons('Isometric view', circle_position['4'][0], circle_position['4'][1], circle_position['4'][2],
#                point_position['4'][0], point_position['4'][1], circle_position['4'][1])
# draft.balloons('Isometric view', circle_position['5'][0], circle_position['5'][1], circle_position['5'][2],
#                point_position['5'][0], point_position['5'][1], circle_position['5'][1])
# draft.balloons('Isometric view', circle_position['3'][0], circle_position['3'][1], circle_position['3'][2],
#                point_position['3'][0], point_position['3'][1], circle_position['3'][1])
#
# draft.create_center_line(0, 0, -CLOCK_Pointer * cos45, -CLOCK_Pointer * cos45 * sin30 / cos30)
# draft.create_center_line(-BLOSTER1_center_Offset_Value * cos45, -BLOSTER1_center_Offset_Value * cos45 * sin30 / cos30,
#                          -BLOSTER1_center_Offset_Value * cos45, -S[i] - Z[i] - T[i])
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'XZ.PLANE', 0,
#                                           173)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', H[i] - 12 - S[i] - Z[i], 'YZ.PLANE', 0,
#                                   174)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', -422.052, 'XY.PLANE',
#                                   0, 175)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 0, 'XZ.PLANE',
#                                   0, 187)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 35, 'YZ.PLANE',
#                                   1, 188)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 0, 'XY.PLANE',
#                                   1, 189)


# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'XZ.PLANE', 0,
#                                           173)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1',  33, 'YZ.PLANE', 0,
#                                   174)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', -(H[i] - 34 - S[i] - Z[i]), 'XY.PLANE',
#                                   0, 175)

l = 1
h = 1
i = 5
# for i in range(0, 10):
#     print('S = ' + str(S[i]) + 'mm')


# mprog.base_lock('FRAME20.1', 'FRAME20.1', 0)  # 基準零件(定海神針)
#
# # (0表示SAME, 1表示OPPOSITE)
# # 平板-四底座
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', A_15[i] / 2, 'XZ.PLANE', 1, 1)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', A[i] / 2, 'XZ.PLANE', 1, 1)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -Z[i], 'XY.PLANE', 1, 2)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - B_15[i] + F[i] / 2, 'YZ.PLANE', 1, 3)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - B[i] + F[i] / 2, 'YZ.PLANE', 1, 3)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -A_15[i] / 2, 'XZ.PLANE', 0, 4)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -A[i] / 2, 'XZ.PLANE', 0, 4)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -Z[i], 'XY.PLANE', 0, 5)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - B_15[i] + F[i] / 2, 'YZ.PLANE', 1, 6)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - B[i] + F[i] / 2, 'YZ.PLANE', 1, 6)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -A_15[i] / 2, 'XZ.PLANE', 0, 7)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -A[i] / 2, 'XZ.PLANE', 0, 7)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -Z[i], 'XY.PLANE', 1, 8)
# if l == 0:
#     mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -B_15[i], 'YZ.PLANE', 1, 9)
# else:
#     mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -B[i], 'YZ.PLANE', 1, 9)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', A_15[i] / 2, 'XZ.PLANE', 0, 10)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', A[i] / 2, 'XZ.PLANE', 0, 10)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', -Z[i], 'XY.PLANE', 0, 11)
# if l == 0:
#     mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -B_15[i], 'YZ.PLANE', 0, 12)
# else:
#     mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -B[i], 'YZ.PLANE', 0, 12)
# # 左右側板
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R_15[i] / 2 + 140, 'XZ.PLANE', 1, 13)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R[i] / 2 + 140, 'XZ.PLANE', 1, 13)
# mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', H[i] / 2, 'XY.PLANE', 0, 14)
# if l == 0:
#     mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', B_15[i] / 2, 'YZ.PLANE', 0, 15)
# else:
#     mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', B[i] / 2, 'YZ.PLANE', 0, 15)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -R_15[i] / 2 - 140, 'XZ.PLANE', 1, 16)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -R[i] / 2 - 140, 'XZ.PLANE', 1, 16)
# mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -H[i] / 2, 'XY.PLANE', 1, 17)
# if l == 0:
#     mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -B_15[i] / 2, 'YZ.PLANE', 1, 18)
# else:
#     mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -B[i] / 2, 'YZ.PLANE', 1, 18)
# # 底部前、中板
# if l == 0:
#     mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', A_15[i] / 2, 'XZ.PLANE', 1, 19)
# else:
#     mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', A[i] / 2, 'XZ.PLANE', 1, 19)
# mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'XY.PLANE', 0, 20)
# mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'YZ.PLANE', 1, 21)
# mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XZ.PLANE', 0, 22)
# mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XY.PLANE', 0, 23)
# if l == 0:
#     mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -FRAME2_lower_depth_15[i] + 5, 'YZ.PLANE', 0, 24)
# else:
#     mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -FRAME2_lower_depth[i] + 5, 'YZ.PLANE', 0, 24)
# # 中間左右側板
# if l == 0:
#     mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', R_15[i] / 2, 'XZ.PLANE', 0, 25)
# else:
#     mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', R[i] / 2, 'XZ.PLANE', 0, 25)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME11.1', -T[i], 'XY.PLANE', 1, 26)
# if l == 0:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', FRAME2_lower_depth_15[i], 'YZ.PLANE', 1, 27)
# else:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', FRAME2_lower_depth[i], 'YZ.PLANE', 1, 27)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 0, 28)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R[i] / 2 - 90, 'XZ.PLANE', 0, 28)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -T[i], 'XY.PLANE', 0, 29)
# if l == 0:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', FRAME2_lower_depth_15[i], 'YZ.PLANE', 1, 30)
# else:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', FRAME2_lower_depth[i], 'YZ.PLANE', 1, 30)
# # 底部後面ㄇ形角鐵
# mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XZ.PLANE', 1, 31)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XY.PLANE', 0, 32)
# if l == 0:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', B_15[i], 'YZ.PLANE', 0, 33)
# else:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', B[i], 'YZ.PLANE', 0, 33)
# # 平板底部零件
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -R_15[i] / 2, 'XZ.PLANE', 1, 34)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -R[i] / 2, 'XZ.PLANE', 1, 34)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -T[i], 'XY.PLANE', 0, 35)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', 0, 'YZ.PLANE', 1, 36)
# # 平板底部零件
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', R_15[i] / 2, 'XZ.PLANE', 1, 37)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', R[i] / 2, 'XZ.PLANE', 1, 37)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', -T[i], 'XY.PLANE', 0, 38)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', 0, 'YZ.PLANE', 1, 39)
# # 前中上軸承板&角鐵
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME20.1', 0, 'XZ.PLANE', 1, 40)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', -H[i] + 40, 'XY.PLANE', 0, 41)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', 0, 'YZ.PLANE', 0, 42)
# mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', 0, 'XZ.PLANE', 0, 43)
# mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -5, 'XY.PLANE', 0, 44)
# mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -550, 'YZ.PLANE', 0, 45)
# mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'XZ.PLANE', 1, 46)
# mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', -FRAME20_H[i], 'XY.PLANE', 0, 47)
# mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'YZ.PLANE', 1, 48)
# # 氣壓缸鎖固板左右
# if i == 4:
#     mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', 183, 'XZ.PLANE', 1, 49)
# else:
#     mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', 242, 'XZ.PLANE', 1, 49)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME21.1', -H[i] + 40, 'XY.PLANE', 1, 50)
# mprog.add_offset_assembly('FRAME20.1', 'FRAME21.1', 280, 'YZ.PLANE', 1, 51)
# if i == 4:
#     mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', -183, 'XZ.PLANE', 1, 52)
# else:
#     mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', -242, 'XZ.PLANE', 1, 52)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME22.1', -H[i] + 40, 'XY.PLANE', 0, 53)
# mprog.add_offset_assembly('FRAME20.1', 'FRAME22.1', 280, 'YZ.PLANE', 1, 54)
# # 鎖固平板六兄弟
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', 0, 'XZ.PLANE', 0, 55)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', -T[i], 'XY.PLANE', 0, 56)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME14.1', -75, 'YZ.PLANE', 1, 57)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 58)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', R[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 58)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -T[i], 'XY.PLANE', 0, 59)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -F[i] / 2 + 80 + 37.5, 'YZ.PLANE', 1, 60)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 61)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', R[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 61)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', -T[i], 'XY.PLANE', 0, 62)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', F[i] / 2 - 80 - 37.5, 'YZ.PLANE', 1, 63)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 64)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -R[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 64)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -T[i], 'XY.PLANE', 0, 65)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', F[i] / 2 - 80 - 37.5, 'YZ.PLANE', 0, 66)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 67)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -R[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 67)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -T[i], 'XY.PLANE', 0, 68)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -F[i] / 2 + 80 + 37.5, 'YZ.PLANE', 0, 69)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', 0, 'XZ.PLANE', 1, 70)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', -T[i], 'XY.PLANE', 0, 71)
# mprog.add_offset_assembly('FRAME9.1', 'FRAME17.1', 0, 'YZ.PLANE', 0, 72)
# # 左右側板前GIB
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME3.1', FRAME1_lower_high[i] + 40 - 34.5, 'XY.PLANE', 0, 73)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME3.1', FRAME1_lower_high[i] + 40, 'XY.PLANE', 0, 73)
# mprog.add_offset_assembly('GIB1.1', 'FRAME1.1', 72.5, 'XZ.PLANE', 0, 74)
# mprog.add_offset_assembly('GIB1.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0, 75)
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME3.1', FRAME1_lower_high[i] + 40 - 34.5, 'XY.PLANE', 0, 76)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME3.1', FRAME1_lower_high[i] + 40, 'XY.PLANE', 0, 76)
# mprog.add_offset_assembly('GIB2.1', 'FRAME2.1', -72.5, 'XZ.PLANE', 0, 77)
# mprog.add_offset_assembly('GIB2.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0, 78)
# # 左GIB後鎖固用方塊
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -690 - 34.5, 'XY.PLANE', 1, 79)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -690, 'XY.PLANE', 1, 79)
# mprog.add_offset_assembly('FRAME2.1', 'FRAME23.1', 50 + 35, 'XZ.PLANE', 1, 80)
# mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', 0, 'YZ.PLANE', 0, 81)
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', -34.5, 'XY.PLANE', 1, 82)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'XY.PLANE', 1, 82)
# mprog.add_offset_assembly('FRAME2.1', 'FRAME24.1', 50 + 35, 'XZ.PLANE', 0, 83)
# mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'YZ.PLANE', 1, 84)
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -690 / 2 - 34.5, 'XY.PLANE', 1, 85)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -690 / 2, 'XY.PLANE', 1, 85)
# mprog.add_offset_assembly('FRAME2.1', 'FRAME27.1', 50, 'XZ.PLANE', 1, 86)
# mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', 0, 'YZ.PLANE', 1, 87)
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -690 - 34.5, 'XY.PLANE', 1, 88)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -690, 'XY.PLANE', 1, 88)
# mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -50, 'XZ.PLANE', 0, 89)
# mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -130.35, 'YZ.PLANE', 0, 90)
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150 - 34.5, 'XY.PLANE', 1, 91)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150, 'XY.PLANE', 1, 91)
# mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -50, 'YZ.PLANE', 0, 92)
# mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -62.5, 'XZ.PLANE', 0, 93)
# # 右GIB後鎖固用方塊
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -690 - 34.5, 'XY.PLANE', 1, 94)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -690, 'XY.PLANE', 1, 94)
# mprog.add_offset_assembly('FRAME1.1', 'FRAME25.1', -50 - 35, 'XZ.PLANE', 0, 95)
# mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', 0, 'YZ.PLANE', 1, 96)
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', -34.5, 'XY.PLANE', 1, 97)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'XY.PLANE', 1, 97)
# mprog.add_offset_assembly('FRAME1.1', 'FRAME26.1', -50 - 35, 'XZ.PLANE', 0, 98)
# mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'YZ.PLANE', 0, 99)
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -690 / 2 - 34.5, 'XY.PLANE', 1, 100)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -690 / 2, 'XY.PLANE', 1, 100)
# mprog.add_offset_assembly('FRAME1.1', 'FRAME28.1', -50, 'XZ.PLANE', 0, 101)
# mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', 0, 'YZ.PLANE', 1, 102)
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -690 - 34.5, 'XY.PLANE', 1, 103)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -690, 'XY.PLANE', 1, 103)
# mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', -50, 'XZ.PLANE', 1, 104)
# mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', 130.35, 'YZ.PLANE', 1, 105)
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME31.4', -150 - 34.5, 'XY.PLANE', 1, 106)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME31.4', -150, 'XY.PLANE', 1, 106)
# mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'YZ.PLANE', 0, 107)
# mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'XZ.PLANE', 0, 108)
# # 中間與平板鎖固方塊下板子
# mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', -48, 'XY.PLANE', 1, 109)
# mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'XZ.PLANE', 0, 110)
# mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'YZ.PLANE', 1, 111)
# # 後面安裝馬達下板子
# mprog.add_offset_assembly('FRAME41.1', 'FRAME20.1', -1340, 'XY.PLANE', 0, 112)
# mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'XZ.PLANE', 0, 113)
# mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'YZ.PLANE', 0, 114)
# # FRAME2中間下板子
# mprog.add_offset_assembly('FRAME32.1', 'BOLSTER1.1', 0, 'XZ.PLANE', 0, 115)
# mprog.add_offset_assembly('FRAME32.1', 'FRAME3.1', FRAME_32_XY[i] + 19, 'XY.PLANE', 0, 116)
# mprog.add_offset_assembly('FRAME32.1', 'FRAME30.1', 0, 'YZ.PLANE', 0, 117)
# # 後方大軸承支架
# mprog.add_offset_assembly('FRAME10.1', 'FRAME34.2', -485.543, 'YZ.PLANE', 0, 118)
# mprog.add_offset_assembly('FRAME34.2', 'FRAME10.1', FRAME_10_H[i] + 40, 'XY.PLANE', 1, 119)
# if l == 0:
#     mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 0, 120)
# else:
#     mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -R[i] / 2 - 90, 'XZ.PLANE', 0, 120)
# if l == 0:
#     mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -R_15[i] - 180, 'XZ.PLANE', 1, 121)
# else:
#     mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -R[i] - 180, 'XZ.PLANE', 1, 121)
# mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'XY.PLANE', 1, 122)
# mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'YZ.PLANE', 0, 123)
# # 後方大軸承
# mprog.add_offset_assembly('FRAME35.1', 'FRAME3.1', 0, 'XZ.PLANE', 1, 124)
# mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -868, 'XY.PLANE', 0, 125)
# mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -FRAME20_FRAME2_YZ[i] - 32, 'YZ.PLANE', 0, 126)
# # 後方馬達下支撐板
# mprog.add_offset_assembly('FRAME39.1', 'FRAME1.1', 50, 'XZ.PLANE', 0, 127)
# mprog.add_offset_assembly('FRAME39.1', 'FRAME20.1', -360, 'XY.PLANE', 0, 128)
# mprog.add_offset_assembly('FRAME39.1', 'FRAME7.1', 320, 'YZ.PLANE', 0, 129)
# mprog.add_offset_assembly('FRAME38.1', 'FRAME2.1', -50, 'XZ.PLANE', 0, 130)
# mprog.add_offset_assembly('FRAME38.1', 'FRAME20.1', 360, 'XY.PLANE', 1, 131)
# mprog.add_offset_assembly('FRAME38.1', 'FRAME6.1', -320, 'YZ.PLANE', 1, 132)
# mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 0, 'XZ.PLANE', 0, 133)
# mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 19, 'XY.PLANE', 1, 134)
# mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 125, 'YZ.PLANE', 1, 135)
# mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', 0, 'XZ.PLANE', 0, 136)
# mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', 19, 'XY.PLANE', 1, 137)
# mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', -125, 'YZ.PLANE', 1, 138)
# mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', 0, 'XZ.PLANE', 1, 139)
# mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', 0, 'XY.PLANE', 0, 140)
# mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', -125, 'YZ.PLANE', 1, 141)
# mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 0, 'XZ.PLANE', 1, 142)
# mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 0, 'XY.PLANE', 0, 143)
# mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 125, 'YZ.PLANE', 1, 144)
# # FRAME2右邊角鐵
# mprog.add_offset_assembly('FRAME33.1', 'FRAME2.1', 140, 'XZ.PLANE', 1, 145)
# mprog.add_offset_assembly('FRAME33.1', 'FRAME15.1', -708, 'XY.PLANE', 0, 146)
# mprog.add_offset_assembly('FRAME33.1', 'FRAME11.1', 313.984, 'YZ.PLANE', 0, 147)
# # FRAME2中間半圓形零件
# mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', 130.5, 'XZ.PLANE', 0, 148)
# mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', -320, 'XY.PLANE', 0, 149)
# mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', FRAME20_FRAME2_YZ[i], 'YZ.PLANE', 1, 150)
# # FRAME35上兩圓管
# mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XZ.PLANE', 1, 151)
# mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XY.PLANE', 1, 152)
# mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', 88, 'YZ.PLANE', 1, 153)
# mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 272.236, 'XZ.PLANE', 1, 154)
# mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', -272.236, 'XY.PLANE', 1, 155)
# mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 88, 'YZ.PLANE', 1, 156)
# # 後方馬達下支撐板上治具
# mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 95, 'XZ.PLANE', 0, 157)
# mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 22 + 18, 'XY.PLANE', 0, 158)
# mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 312.5, 'YZ.PLANE', 0, 159)
# # 馬達下方板與後方橫板
# mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', 0, 'XZ.PLANE', 1, 160)
# mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', 71.5, 'XY.PLANE', 1, 161)
# mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', -19, 'YZ.PLANE', 1, 162)
# # 平板2, 3組立
# mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0, 163)
# mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XY.PLANE', 0, 164)
# mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0, 165)
# # 平板1, 3組立
# mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0, 166)
# mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0, 167)
# # if k == 0:
# if h == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_S[i], 'XY.PLANE', 0, 168)
# elif h == 1:
#     mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_H[i], 'XY.PLANE', 0, 168)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', DH_P[i], 'XY.PLANE', 0, 168)
# # SLIDE跟平板3組立
# if i == 4:
#     mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 267.38, 'XY.PLANE',
#                                       1, 169)
# else:
#     mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 85, 'XY.PLANE', 1,
#                                       170)
# mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 1, 171)
# if i == 4:
#     mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 138.171,
#                                       'YZ.PLANE', 0, 172)
# else:
#     mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0,
#                                       172)
# # 時鐘跟FRAME20組立
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'XZ.PLANE', 0,
#                                   173)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1',  37, 'YZ.PLANE', 0,
#                                   174)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', -(H[i] - 12 - S[i] - Z[i]), 'XY.PLANE',
#                                   0, 175)
# # 氣壓缸跟FRAME20組立
# if i == 4:
#     mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -87.5,
#                                       'XY.PLANE', 1, 176)
# else:
#     mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE',
#                                       1, 176)
# if i == 4:
#     if l == 0:
#         mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -500.5,
#                                           'XZ.PLANE', 0, 177)
#     else:
#         mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -347,
#                                           'XZ.PLANE', 0, 177)
# else:
#     if l == 0:
#         mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                           -R_15[i] / 2 - 13, 'XZ.PLANE', 0, 178)
#     else:
#         mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                           -R[i] / 2 - 13, 'XZ.PLANE', 0, 178)
# if i == 4:
#     mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -261.248,
#                                       'YZ.PLANE', 0, 179)
# else:
#     mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -260,
#                                       'YZ.PLANE', 0, 179)
# if i == 4:
#     mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -21.8,
#                                       'XY.PLANE', 1, 180)
# else:
#     mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE',
#                                       1, 180)
# if i == 4:
#     if l == 0:
#         mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -500.5,
#                                           'XZ.PLANE', 1, 181)
#     else:
#         mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -347,
#                                           'XZ.PLANE', 1, 181)
# else:
#     if l == 0:
#         mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                           -R_15[i] / 2 - 13, 'XZ.PLANE', 1, 182)
#     else:
#         mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                           -R[i] / 2 - 13, 'XZ.PLANE', 1, 182)
# if i == 4:
#     mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 261.248,
#                                       'YZ.PLANE', 1, 183)
# else:
#     mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 260, 'YZ.PLANE',
#                                       1, 183)
# # 離合器與FRAME20結合
# if i == 4:
#     mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 108.109,
#                                       'XY.PLANE', 0, 184)
# else:
#     mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XY.PLANE',
#                                       0, 184)
# mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE', 1,
#                                   185)
# mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 510, 'YZ.PLANE', 1,
#                                   186)
# # 曲軸與時鐘結合
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 0, 'XZ.PLANE',
#                                   0, 187)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 35, 'YZ.PLANE',
#                                   1, 188)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 0, 'XY.PLANE',
#                                   1, 189)
# # 大齒輪結合
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 0, 'XZ.PLANE', 0, 190)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 77.172, 'XY.PLANE', 0, 191)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 0, 'YZ.PLANE', 1, 192)
# # 大齒輪結合曲軸
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'CRANK_SHAFT.1.1', 0, 'XZ.PLANE', 1, 193)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'CRANK_SHAFT.1.1', 0, 'XY.PLANE', 0, 194)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'FRAME20.1', 592.5 + 150, 'YZ.PLANE', 1, 195)
# # 機架20結合曲軸旁圓管
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', 0, 'XZ.PLANE', 0, 196)
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', 420, 'XY.PLANE', 1, 197)
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', H[i] - Z[i] - S[i] - 34, 'YZ.PLANE', 1, 198)
# # 大齒輪結合長棒
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', 0, 'XZ.PLANE', 0, 199)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', 0, 'XY.PLANE', 1, 200)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', -84, 'YZ.PLANE', 1, 201)
# #
# mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE',
#                                   0, 202)
# mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', 1010, 'YZ.PLANE',
#                                   1, 203)
# mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', -(H[i] - Z[i] - S[i] - 34), 'XY.PLANE',
#                                   0, 204)
# # 大齒輪內套環
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 0, 'XZ.PLANE', 1, 205)
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 0, 'XY.PLANE', 0, 206)
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 11, 'YZ.PLANE', 1, 207)
# # BOLSTER後板
# mprog.add_offset_assembly('FRAME45.1', 'FRAME17.1', 0, 'XZ.PLANE', 0, 208)
# mprog.add_offset_assembly('FRAME45.1', 'FRAME17.1', -48, 'XY.PLANE', 1, 209)
# mprog.add_offset_assembly('FRAME45.1', 'FRAME17.1', 0, 'YZ.PLANE', 1, 210)
#
# mprog.add_offset_assembly('FRAME46.1', 'FRAME17.1', -48, 'XY.PLANE', 1, 211)
# mprog.add_offset_assembly('FRAME46.1', 'FRAME17.1', 0, 'XZ.PLANE', 0, 212)
# mprog.add_offset_assembly('FRAME46.1', 'FRAME17.1', 0, 'YZ.PLANE', 1, 213)
#
# mprog.add_offset_assembly('FRAME20.1', 'FRAME44.1', 5, 'XY.PLANE', 0, 214)
# mprog.add_offset_assembly('FRAME20.1', 'FRAME44.1', 0, 'XZ.PLANE', 0, 215)
# mprog.add_offset_assembly('FRAME20.1', 'FRAME44.1', 979, 'YZ.PLANE', 1, 216)

# # --------------------------------------- 生成爆炸圖-------------------------------------------
# # 重新定義拘束尺寸
# BOLSTER1_Offset_value = 1750
# GIB_Offset_value = -3000
# Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
# CLOCK_Offset_value = 2850
# CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
# BOLSTER1_XY_Z = -T[i]
# BOLSTER1_XY_T = Z[i] - T[i] + BOLSTER1_XY_Z
# BOLSTER1_XY = BOLSTER1_XY_T + DH_S[i]
# BALANCER = -1700
#
# mprog.constaint_value_change(3, -B[i] - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(6, -B[i] - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(36, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
# mprog.constaint_value_change(39, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
# mprog.constaint_value_change(63, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(66, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
# mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
# mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
# mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
# mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
# mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
# mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
# mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
# mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
# mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
# mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
# mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
# mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
# mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
# mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
# mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(168, BOLSTER1_XY, 0)
#
# mprog.update()  # 更新
# mprog.OPEN_Drawing()
# drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale = draft.drafting_parameter_calculation(
#     A[i], B[i], H[i], T[i])  # 計算爆炸圖比例及位置
# print(drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale)
# draft.change_Drawing_scale(1 / scale)  # 圖面比例
# draft.exploded_Drawing_1(type, drafting_isometric_Coordinate_Position['exploded_1'][0],
#                          drafting_isometric_Coordinate_Position['exploded_1'][1], scale)  # 爆炸圖1
# mprog.switch_window()  # 開啟3D圖視窗
# # # 還原零件初始位置
# BOLSTER1_Offset_value = 80 - F[i] / 2
# GIB_Offset_value = 334.5 - 45
# CLOCK_Offset_value = 35
# CLOCK_SHAFT_Offset_value = 45
# Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
# BOLSTER1_XY_Z = - Z[i]
# BOLSTER1_XY_T = -T[i]
# BOLSTER1_XY = DH_S[i]
# BALANCER = -32
#
# mprog.constaint_value_change(3, -B[i] - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(6, -B[i] - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(36, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
# mprog.constaint_value_change(39, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
# mprog.constaint_value_change(63, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(66, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
# mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
# mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
# mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
# mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
# mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
# mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
# mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
# mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
# mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
# mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
# mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
# mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
# mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
# mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
# mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(168, BOLSTER1_XY, 0)
# mprog.update()
# # # # --------------------爆炸圖右圖------------------
# mprog.constaint_value_change(159, -3000, 0)  # 支架
# mprog.constaint_value_change(186, 3500, 1)  # 離合器
# mprog.constaint_value_change(203, 3400, 1)  # JOINT_All.1
# mprog.constaint_value_change(195, 2500, 1)  # 大齒輪MAIN_GEA1
# mprog.constaint_value_change(201, -500, 1)  # Joint1
# mprog.update()
# mprog.switch_window()
# draft.exploded_Drawing_2(type, drafting_isometric_Coordinate_Position['exploded_2'][0],
#                          drafting_isometric_Coordinate_Position['exploded_2'][1], scale)
# mprog.switch_window()
# # 復原位置
# mprog.constaint_value_change(159, 312.5, 0)  # 支架
# mprog.constaint_value_change(186, 510, 1)  # 離合器
# mprog.constaint_value_change(203, 1010, 1)  # JOINT_All.1
# mprog.constaint_value_change(195, 592.5 + 150, 1)  # 大齒輪MAIN_GEA1
# mprog.constaint_value_change(201, -84, 1)  # Joint1
# mprog.update()
# mprog.switch_window()
# # # --------------------爆炸圖下(前、左、右)----------------
# draft.Front_View_Drawing(type, drafting_Coordinate_Position['Front View'][0],
#                          drafting_Coordinate_Position['Front View'][1], scale)
# draft.Left_View_Drawing(type, drafting_Coordinate_Position['Left View'][0],
#                         drafting_Coordinate_Position['Left View'][1], scale)
# draft.Right_View_Drawing(type, drafting_Coordinate_Position['Right View'][0],
#                          drafting_Coordinate_Position['Right View'][1], scale)
#
# # ------------圈碼圖------------
sin45 = math.sin(math.radians(45))
cos45 = math.cos(math.radians(45))
sin30 = math.sin(math.radians(30))
cos30 = math.cos(math.radians(30))
cos35_267 = math.cos(math.radians(35.267))
sin35_267 = math.sin(math.radians(35.267))
#
# BOLSTER1_Offset_value = 1750
# BLOSTER1_center_Offset_Value = 1775
# GIB_Offset_value = -3000
# Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
# CLOCK_Offset_value = 2850
# CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
# BOLSTER1_XY_Z = -T[i]
# BOLSTER1_XY_T = Z[i] - T[i] + BOLSTER1_XY_Z
# BOLSTER1_XY = BOLSTER1_XY_T + DH_S[i]
#
# balloons_posotion = [750 * cos30, 750 * sin30]
# CLOCK_Pointer = 3325
# FRAME_TOP_LEFT_X = (R[i] + 90 + 50) * cos45 + 50
# FRAME_TOP_LEFT_X_1 = FRAME_TOP_LEFT_X * sin30 / cos30
#
# FRAME_TOP_LEFT_X_Y = H[i] - S[i] - Z[i] - 12  # 因X產生微小高度
# FRAME_TOP_LEFT_Y = FRAME_TOP_LEFT_X * math.sin(math.radians(57.24)) / math.cos(math.radians(57.24)) + 112
#
# scale = 1 / scale
#
# # (將3D圖面尺寸轉換為2D尺寸, 將圖面x座標尺寸轉為R再將R轉為Y)
# point_position = {'2': [-CLOCK_Offset_value * cos45,
#                         -CLOCK_Offset_value * cos45 * sin30 / cos30],
#                   '3': [(-1250) * cos45,
#                         (-1250) * cos45 * sin30 / cos30],
#                   '4': [(-BOLSTER1_Offset_value - R[i] / 2) * cos45,
#                         (-BOLSTER1_Offset_value - R[i] / 2) * cos45 * sin30 / cos30],
#                   '5': [(-BOLSTER1_Offset_value - R[i]) * cos45,
#                         (-BOLSTER1_Offset_value - R[i]) * cos45 * sin30 / cos30 - S[i] + T[i]],
#                   '6': [(GIB_Offset_value - (R[i] + 145) / 2) * cos45,
#                         (GIB_Offset_value - (R[i] + 145) / 2) * cos45 * sin30 / cos30]}
#
# circle_position = {
#     '2': ['2', point_position['2'][0] - FRAME_TOP_LEFT_X, point_position['2'][1] + FRAME_TOP_LEFT_X_1],
#     '3': ['3', point_position['3'][0] - FRAME_TOP_LEFT_X, point_position['3'][1] + FRAME_TOP_LEFT_X_1],
#     '4': ['4', point_position['4'][0] - FRAME_TOP_LEFT_X, point_position['4'][1] + FRAME_TOP_LEFT_X_1],
#     '5': ['5', point_position['5'][0] * 2 * 0.9 - FRAME_TOP_LEFT_X,
#           point_position['5'][0] * 2 * 0.9 * sin30 / cos30 + FRAME_TOP_LEFT_X_1],
#     '6': ['6', point_position['6'][0] - FRAME_TOP_LEFT_X, point_position['6'][1] + FRAME_TOP_LEFT_X_1]}
#
# draft.balloons('Isometric view', circle_position['6'][0], circle_position['6'][1], circle_position['6'][2],
#                point_position['6'][0], point_position['6'][1], circle_position['6'][1])
# draft.balloons('Isometric view', circle_position['2'][0], circle_position['2'][1], circle_position['2'][2],
#                point_position['2'][0], point_position['2'][1], circle_position['2'][1])
# draft.balloons('Isometric view', circle_position['4'][0], circle_position['4'][1], circle_position['4'][2],
#                point_position['4'][0], point_position['4'][1], circle_position['4'][1])
# draft.balloons('Isometric view', circle_position['5'][0], circle_position['5'][1], circle_position['5'][2],
#                point_position['5'][0], point_position['5'][1], circle_position['5'][1])
# draft.balloons('Isometric view', circle_position['3'][0], circle_position['3'][1], circle_position['3'][2],
#                point_position['3'][0], point_position['3'][1], circle_position['3'][1])
#
# # ------------中心線-------------
# draft.create_center_line(0, 0, -CLOCK_Pointer * cos45, -CLOCK_Pointer * cos45 * sin30 / cos30)
# draft.create_center_line(-BLOSTER1_center_Offset_Value * cos45,
#                          -BLOSTER1_center_Offset_Value * cos45 * sin30 / cos30,
#                          -BLOSTER1_center_Offset_Value * cos45, -S[i] - Z[i] - T[i])

# Fixture_offset_value = 2850
# CLUCTH_ASSEMBLY_offset_value = B[i] + 1650
# JOINT_ALL_offset_value = CLUCTH_ASSEMBLY_offset_value
# MAIN_GEAR_offset_value = B[i] + 850
#
# draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30, -JOINT_ALL_offset_value * cos45 - 60, -JOINT_ALL_offset_value * cos45 * sin30 / cos30 - 60)
# draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30 - 350,
#                          -(Fixture_offset_value + B[i] - 742.5) * cos45, -(Fixture_offset_value + B[i] - 742.5) * cos45 * sin30 / cos30 - 350)
# draft.add_dimension_to_view('Isometric view1', 0, -(R[i] / 2 + 140 + 25) * cos45,
#                             (H[i] - Z[i] - S[i] - 34 - 25) * cos35_267 + (R[i] / 2 + 140 + 25) * cos45 * sin30 / cos30,
#                             -(A[i] / 2) * cos45, -(S[i] + Z[i]) * sin35_267 / cos35_267 - (R[i] / 2 + 140 + 25) * cos45 * sin30 / cos30, 90)
# draft.add_dimension_to_view('Left view', 0, 0, (H[i] - Z[i] - S[i] - 34), 0, -(S[i] + Z[i]), 90)
# # (R[i] + 90) * cos45, H[i] - Z[i] - S[i] - 34, (R[i] + 90) * cos45, -S[i] - Z[i]


# Left_view = {'H': [0, 0, (H[i] - Z[i] - S[i] - 34), 0, -(S[i] + Z[i]), 90],
#              'D': [0, 0, (H[i] - Z[i] - S[i] - 34), -B[i], (H[i] - Z[i] - S[i] - 34) - 110, 0]}
#
# Left_view_list = []
# for key in Left_view:
#     Left_view_list.append(key)
# print(Left_view_list)
#
# for i in Left_view_list:
#     print(Left_view[i][0])
#     draft.add_dimension_to_view('Left view', str(i), Left_view[i][0], Left_view[i][1], Left_view[i][2], Left_view[i][3], Left_view[i][4], Left_view[i][5])
#
# view_name = ['Isometric view1', 'Isometric view2', 'Front view', 'Left view', 'Right view']
# for i in view_name:
#     draft.close_broken_line_block_diagram(i)

# mprog.add_offset_assembly('FRAME26.1', 'FRAME47.1', 29, 'XY.PLANE', 1, 217)
# mprog.add_offset_assembly('FRAME26.1', 'FRAME47.1', 35, 'XZ.PLANE', 0, 218)
# mprog.add_offset_assembly('FRAME26.1', 'FRAME47.1', -99.35, 'YZ.PLANE', 1, 219)
# mprog.add_offset_assembly('FRAME24.1', 'FRAME47.2', 29, 'XY.PLANE', 0, 220)
# mprog.add_offset_assembly('FRAME24.1', 'FRAME47.2', -35, 'XZ.PLANE', 1, 221)
# mprog.add_offset_assembly('FRAME24.1', 'FRAME47.2', 099.35, 'YZ.PLANE', 0, 222)
file_Assembly_name_list = ['BOLSTER1', 'BOLSTER2', 'BOLSTER3', 'Fixture', 'FRAME1', 'FRAME2', 'FRAME3',
                                   'FRAME4', 'FRAME5', 'FRAME6', 'FRAME7', 'FRAME8', 'FRAME9', 'FRAME10', 'FRAME11',
                                   'FRAME12', 'FRAME13', 'FRAME14', 'FRAME15', 'FRAME16', 'FRAME17', 'FRAME18',
                                   'FRAME19', 'FRAME20', 'FRAME21', 'FRAME22', 'FRAME23', 'FRAME24', 'FRAME25',
                                   'FRAME26', 'FRAME27', 'FRAME28', 'FRAME29', 'FRAME30', 'FRAME31', 'FRAME31',
                                   'FRAME31', 'FRAME31', 'FRAME32', 'FRAME33', 'FRAME34', 'FRAME34', 'FRAME35',
                                   'FRAME36', 'FRAME36', 'FRAME37', 'FRAME37', 'FRAME37', 'FRAME37', 'FRAME38',
                                   'FRAME39', 'FRAME40', 'FRAME41', 'FRAME42', 'FRAME43', 'FRAME44', 'FRAME45',
                                   'FRAME46', 'FRAME47', 'GIB1', 'GIB2', 'BALANCER_LEFT_All', 'BALANCER_RIGHT_ALL',
                                   'CRANK_SHAFT_CLOCK', 'CLUCTH_ASSEMBLY_All', 'SLIDE_UNIT_ALL', 'CRANK_SHAFT',
                                   'JOINT_All', 'MAIN_GEAR1', 'MAIN_GEAR2', 'MAIN_GEAR3', 'MAIN_GEAR4', 'JOINT1']
# print(len(file_Assembly_name_list))

type = 'SN1-110'
# --------------------------------------- 生成爆炸圖-------------------------------------------
# 重新定義拘束尺寸
# BOLSTER1_Offset_value = 1750
# GIB_Offset_value = -3000
# Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
# CLOCK_Offset_value = 2850
# CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
# BOLSTER1_XY_Z = -T[i]
# BOLSTER1_XY_T = Z[i] - T[i] + BOLSTER1_XY_Z
# BOLSTER1_XY = BOLSTER1_XY_T + DH_S[i]
# BALANCER = -1700
#
# mprog.constaint_value_change(3, -B[i] - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(6, -B[i] - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(36, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
# mprog.constaint_value_change(39, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
# mprog.constaint_value_change(63, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(66, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
# mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
# mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
# mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
# mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
# mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
# mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
# mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
# mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
# mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
# mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
# mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
# mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
# mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
# mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
# mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(168, BOLSTER1_XY, 0)
#
# mprog.update()  # 更新
# mprog.OPEN_Drawing()
# drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale = draft.drafting_parameter_calculation(
#     A[i], B[i], H[i], T[i])  # 計算爆炸圖比例及位置
# print(drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale)
# draft.change_Drawing_scale(1 / scale)  # 圖面比例
# draft.exploded_Drawing_1(type, drafting_isometric_Coordinate_Position['exploded_1'][0],
#                          drafting_isometric_Coordinate_Position['exploded_1'][1], scale)  # 爆炸圖1
# mprog.switch_window()  # 開啟3D圖視窗
# # # 還原零件初始位置
# BOLSTER1_Offset_value = 80 - F[i] / 2
# GIB_Offset_value = 334.5 - 45
# CLOCK_Offset_value = 35
# CLOCK_SHAFT_Offset_value = 45
# Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
# BOLSTER1_XY_Z = - Z[i]
# BOLSTER1_XY_T = -T[i]
# BOLSTER1_XY = DH_S[i]
# BALANCER = -32
#
# mprog.constaint_value_change(3, -B[i] - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(6, -B[i] - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(36, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
# mprog.constaint_value_change(39, -BOLSTER1_Offset_value - F[i] / 2 + 80, 1)
# mprog.constaint_value_change(63, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(66, -F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
# mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
# mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
# mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
# mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
# mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
# mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
# mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
# mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
# mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
# mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
# mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
# mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
# mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
# mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
# mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
# mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
# mprog.constaint_value_change(168, BOLSTER1_XY, 0)
# mprog.update()
# # # # --------------------爆炸圖右圖------------------
#
# Fixture_offset_value = 2850
# CLUCTH_ASSEMBLY_offset_value = B[i] + 1650
# JOINT_ALL_offset_value = CLUCTH_ASSEMBLY_offset_value
# MAIN_GEAR_offset_value = B[i] + 850
#
# mprog.constaint_value_change(159, -312.5 - Fixture_offset_value, 0)  # 支架
# mprog.constaint_value_change(186, CLUCTH_ASSEMBLY_offset_value, 1)  # 離合器
# mprog.constaint_value_change(203, JOINT_ALL_offset_value, 1)  # JOINT_All.1
# mprog.constaint_value_change(195, MAIN_GEAR_offset_value, 1)  # 大齒輪MAIN_GEA1
# mprog.constaint_value_change(201, -375, 1)  # Joint1
# mprog.update()
# mprog.switch_window()
# draft.exploded_Drawing_2(type, drafting_isometric_Coordinate_Position['exploded_2'][0],
#                                  drafting_isometric_Coordinate_Position['exploded_2'][1], scale)  # 爆炸圖2
# # 復原位置
# mprog.switch_window()
# mprog.constaint_value_change(159, 312.5, 0)  # 支架
# mprog.constaint_value_change(186, 510, 1)  # 離合器
# mprog.constaint_value_change(203, 1010, 1)  # JOINT_All.1
# mprog.constaint_value_change(195, 592.5 + 150, 1)  # 大齒輪MAIN_GEA1
# mprog.constaint_value_change(201, -84, 1)  # Joint1
# mprog.update()
# mprog.switch_window()
# # # --------------------爆炸圖下(前、左、右)----------------
# draft.Front_View_Drawing(type, drafting_Coordinate_Position['Front View'][0],
#                          drafting_Coordinate_Position['Front View'][1], scale)
# draft.Left_View_Drawing(type, drafting_Coordinate_Position['Left View'][0],
#                         drafting_Coordinate_Position['Left View'][1], scale)
# draft.Right_View_Drawing(type, drafting_Coordinate_Position['Right View'][0],
#                          drafting_Coordinate_Position['Right View'][1], scale)
#
# # ------------圈碼圖------------
# sin45 = math.sin(math.radians(45))
# cos45 = math.cos(math.radians(45))
# sin30 = math.sin(math.radians(30))
# cos30 = math.cos(math.radians(30))
# sin60 = math.sin(math.radians(60))
# cos60 = math.cos(math.radians(60))
# cos35_267 = math.cos(math.radians(35.267))
# sin35_267 = math.sin(math.radians(35.267))
#
# BOLSTER1_Offset_value = 1750
# BLOSTER1_center_Offset_Value = 1775
# GIB_Offset_value = -3000
# Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
# CLOCK_Offset_value = 2850
# CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
# BOLSTER1_XY_Z = -T[i]
# BOLSTER1_XY_T = Z[i] - T[i] + BOLSTER1_XY_Z
# BOLSTER1_XY = BOLSTER1_XY_T + DH_S[i]
#
# balloons_posotion = [750 * cos30, 750 * sin30]
# CLOCK_Pointer = 3325
# FRAME_TOP_LEFT_X = (R[i] + 90 + 50) * cos45 + 50
# FRAME_TOP_LEFT_X_1 = FRAME_TOP_LEFT_X * sin30 / cos30
#
# scale = 1 / scale
#
# # (將3D圖面尺寸轉換為2D尺寸, 將圖面x座標尺寸轉為R再將R轉為Y)
# # 引線點座標
# point_position = {'2': [-CLOCK_Offset_value * cos45,
#                         -CLOCK_Offset_value * cos45 * sin30 / cos30],
#                   '3': [(-1250) * cos45,
#                         (-1250) * cos45 * sin30 / cos30],
#                   '4': [(-BOLSTER1_Offset_value - R[i] / 2) * cos45,
#                         (-BOLSTER1_Offset_value - R[i] / 2) * cos45 * sin30 / cos30],
#                   '5': [(-BOLSTER1_Offset_value - R[i]) * cos45,
#                         (-BOLSTER1_Offset_value - R[i]) * cos45 * sin30 / cos30 - S[i] + T[i]],
#                   '6': [(GIB_Offset_value - (R[i] + 145) / 2) * cos45,
#                         (GIB_Offset_value - (R[i] + 145) / 2) * cos45 * sin30 / cos30]}
# # 圈圈座標
# circle_position = {
#     '2': ['2', point_position['2'][0] - FRAME_TOP_LEFT_X, point_position['2'][1] + FRAME_TOP_LEFT_X_1],
#     '3': ['3', point_position['3'][0] - FRAME_TOP_LEFT_X, point_position['3'][1] + FRAME_TOP_LEFT_X_1],
#     '4': ['4', point_position['4'][0] - FRAME_TOP_LEFT_X, point_position['4'][1] + FRAME_TOP_LEFT_X_1],
#     '5': ['5', point_position['5'][0] * 2 * 0.9 - FRAME_TOP_LEFT_X,
#           point_position['5'][0] * 2 * 0.9 * sin30 / cos30 + FRAME_TOP_LEFT_X_1],
#     '6': ['6', point_position['6'][0] - FRAME_TOP_LEFT_X, point_position['6'][1] + FRAME_TOP_LEFT_X_1]}
#
# draft.balloons('Isometric view1', circle_position['6'][0], circle_position['6'][1], circle_position['6'][2],
#                point_position['6'][0], point_position['6'][1], circle_position['6'][1])
# draft.balloons('Isometric view1', circle_position['2'][0], circle_position['2'][1], circle_position['2'][2],
#                point_position['2'][0], point_position['2'][1], circle_position['2'][1])
# draft.balloons('Isometric view1', circle_position['4'][0], circle_position['4'][1], circle_position['4'][2],
#                point_position['4'][0], point_position['4'][1], circle_position['4'][1])
# draft.balloons('Isometric view1', circle_position['5'][0], circle_position['5'][1], circle_position['5'][2],
#                point_position['5'][0], point_position['5'][1], circle_position['5'][1])
# draft.balloons('Isometric view1', circle_position['3'][0], circle_position['3'][1], circle_position['3'][2],
#                point_position['3'][0], point_position['3'][1], circle_position['3'][1])
#
# # ------------中心線-------------
# draft.create_center_line('Isometric view1', 0, 0, -CLOCK_Pointer * cos45,
#                          -CLOCK_Pointer * cos45 * sin30 / cos30)
# draft.create_center_line('Isometric view1', -BLOSTER1_center_Offset_Value * cos45,
#                          -BLOSTER1_center_Offset_Value * cos45 * sin30 / cos30,
#                          -BLOSTER1_center_Offset_Value * cos45, -S[i] - Z[i] - T[i])
# draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30,
#                          -JOINT_ALL_offset_value * cos45, -JOINT_ALL_offset_value * cos45 * sin30 / cos30)
# draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30 - 300,
#                          -(Fixture_offset_value + B[i]) * cos45,
#                          -(Fixture_offset_value + B[i]) * cos45 * sin30 / cos30 - 300)
#
#
# Left_view_list = []
# for key in Left_view:
#     Left_view_list.append(key)
#
# for i in Left_view_list:
#     draft.add_dimension_to_view('Left view', str(i), Left_view[i][0], Left_view[i][1], Left_view[i][2],
#                                 Left_view[i][3], Left_view[i][4], Left_view[i][5])
# # -------------關閉虛框---------------
# view_name = ['Isometric view1', 'Isometric view2', 'Front view', 'Left view', 'Right view']
# for i in view_name:
#     draft.close_broken_line_block_diagram(i)

Fixture_offset_value = 2500
CLUCTH_ASSEMBLY_offset_value = B[i] + 1650
JOINT_ALL_offset_value = CLUCTH_ASSEMBLY_offset_value
MAIN_GEAR_offset_value = B[i] + 850
CLOCK_Pointer = 3325
FRAME_TOP_LEFT_X = (R[i] + 90 + 50) * cos45 + 50
FRAME_TOP_LEFT_X_1 = FRAME_TOP_LEFT_X * sin30 / cos30
BOLSTER1_Offset_value = 1750
BLOSTER1_center_Offset_Value = 1775
GIB_Offset_value = -3000
Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
CLOCK_Offset_value = 2850
CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
BOLSTER1_XY_Z = -T[i]
BOLSTER1_XY_T = Z[i] - T[i] + BOLSTER1_XY_Z
BOLSTER1_XY = BOLSTER1_XY_T + DH_S[i]

#
mprog.constaint_value_change(159, -312.5 - Fixture_offset_value, 0)  # 支架
mprog.constaint_value_change(186, CLUCTH_ASSEMBLY_offset_value, 1)  # 離合器
mprog.constaint_value_change(203, JOINT_ALL_offset_value, 1)  # JOINT_All.1
mprog.constaint_value_change(195, MAIN_GEAR_offset_value, 1)  # 大齒輪MAIN_GEA1
mprog.constaint_value_change(201, -375, 1)  # Joint1
mprog.update()

draft.create_center_line('Isometric view1', 0, 0, -CLOCK_Pointer * cos45,
                         -CLOCK_Pointer * cos45 * sin30 / cos30)
draft.create_center_line('Isometric view1', -BLOSTER1_center_Offset_Value * cos45,
                         -BLOSTER1_center_Offset_Value * cos45 * sin30 / cos30,
                         -BLOSTER1_center_Offset_Value * cos45, -S[i] - Z[i] - T[i])
draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30,
                         -JOINT_ALL_offset_value * cos45, -JOINT_ALL_offset_value * cos45 * sin30 / cos30)
draft.create_center_line('Isometric view2', -742.5 * cos45, -742.5 * cos45 * sin30 / cos30,
                         -(Fixture_offset_value + B[i]) * cos45,
                         -(Fixture_offset_value + B[i]) * cos45 * sin30 / cos30)