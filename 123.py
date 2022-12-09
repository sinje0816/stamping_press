import os
import win32com.client as win32
import datetime, time
import main_program as mprog
import drafting as draft

A = [720, 830, 890, 940, 1050, 1160, 1300, 1480, 1560]
B = [1058, 1125, 1210, 1315, 1480, 1680, 1985, 2113, 2400]
H = [2060, 2185, 2290, 2540, 2755, 2990, 3270, 3725, 4005]
R = [388, 486, 516, 544, 614, 670, 730, 900, 970]
E = [700, 780, 840, 900, 1050, 1150, 1250, 1400, 1500]
D_DH = [250, 280, 330, 350, 380, 430, 490, 550, 580]
S = [983, 1068, 1158, 1285, 1445, 1630, 1809, 2067, 2262]
H_Z = [1260, 1385, 1490, 1640, 1855, 2086.933, 2370, 2725, 3005]
O = [1045, 1075, 1125, 1145, 1175, 1225, 1285, 1345, 1375]
hole_type = [0, 1, 2]
P = [330, 380, 430, 480, 560, 650, 720, 860, 960]
Q = [250, 300, 350, 400, 460, 520, 580, 650, 720]
T = [85, 100, 115, 130, 140, 155, 165, 180, 180]
Z = [800, 800, 800, 900, 900, 900, 900, 1000, 1000]
FRAME1_cutout_bottom = [197, 277, 317, 397, 477, 557, 637, 717, 797]
F = [320, 400, 440, 520, 600, 680, 760, 840, 900]
FRAME1_cutout = [655 + 370, 675 + 415, 695 + 450, 715 + 577.5, 735 + 670, 755 + 770, 775 + 895, 795 + 1080, 815 + 1210]
FRAME20_H = [280, 355, 420, 437.5, 550, 680, 825, 990, 1120]
FRAME2_lower_depth = [165, 245, 285, 365, 445, 525, 605, 685, 745]
FRAME1_lower_high = [1330, 1335, 1340, 1445, 1455, 1460, 1470, 1575, 1595]
FRAME20_FRAME2_YZ = [805, 805, 979, 979, 979, 979, 979, 979, 979]
BALANCER1_XZ = [204, 253, 268, 282, 317, 345, 375, 460, 575]  #
i = 5





# mprog.constaint_value_change(3, -80 - B[i] + F[i] / 2 - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(6, -80 - B[i] + F[i] / 2 - BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(36, -BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(39, -BOLSTER1_Offset_value, 1)
# mprog.constaint_value_change(63, F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value - 80, 1)
# mprog.constaint_value_change(66, F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value - 80, 0)
# mprog.constaint_value_change(60, -F[i] / 2 + 80 + 37.5 - BOLSTER1_Offset_value + 80, 1)
# mprog.constaint_value_change(69, -F[i] / 2 + 80 + 37.5 - BOLSTER1_Offset_value + 80, 0)
# mprog.constaint_value_change(167, 0, 0)  # BLOSTER1 & 3
# mprog.constaint_value_change(75, GIB_Offset_value, 0)
# mprog.constaint_value_change(78, GIB_Offset_value, 0)
# mprog.constaint_value_change(84, GIB_Offset_value - 334.65, 1)
# mprog.constaint_value_change(87, GIB_Offset_value - 334.65, 1)
# mprog.constaint_value_change(81, GIB_Offset_value - 334.65, 0)
# mprog.constaint_value_change(96, GIB_Offset_value - 334.65, 1)
# mprog.constaint_value_change(102, GIB_Offset_value - 334.65, 1)
# mprog.constaint_value_change(99, GIB_Offset_value - 334.65, 0)
# mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
# mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
# mprog.constaint_value_change(195, 765, 0)  # 大齒輪與FRAME20
# mprog.constaint_value_change(176, -21.8, 1)  # 右氣壓缸
# mprog.constaint_value_change(180, -21.8, 1)  # 左氣壓缸
# mprog.update()
# # --------------------爆炸圖右圖------------------
# mprog.constaint_value_change(159, -3000, 0)  # 支架
# mprog.constaint_value_change(186, 3500, 1)  # 離合器
# mprog.constaint_value_change(203, 3400, 1)  # JOINT_All.1
# mprog.constaint_value_change(195, 2500, 0)  # 大齒輪MAIN_GEA1
# mprog.constaint_value_change(201, -500, 1)  # Joint1
# mprog.update()

type = str('SN1-110')
i = 5
drafting_Coordinate_Position, scale = draft.drafting_parameter_calculation(A[i], B[i], H[i])  # 計算爆炸圖比例及位置

BOLSTER1_Offset_value = 2500
GIB_Offset_value = -3200
CLOCK_Offset_value = 2800
CLOCK_SHAFT_Offset_value = 1500

mprog.constaint_value_change(3, -B[i] - BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(6, -B[i] - BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(36, -BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(39, -BOLSTER1_Offset_value, 1)
mprog.constaint_value_change(63, F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value - 80, 1)
mprog.constaint_value_change(66, F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value - 80, 0)
mprog.constaint_value_change(60, -F[i] / 2 + 80 + 37.5 - BOLSTER1_Offset_value + 80, 1)
mprog.constaint_value_change(69, -F[i] / 2 + 80 + 37.5 - BOLSTER1_Offset_value + 80, 0)
mprog.constaint_value_change(167, -250, 0)  # BLOSTER1 & 3
mprog.constaint_value_change(75, GIB_Offset_value, 0)
mprog.constaint_value_change(78, GIB_Offset_value, 0)
mprog.constaint_value_change(84, GIB_Offset_value - 334.65, 1)
mprog.constaint_value_change(87, GIB_Offset_value - 334.65, 1)
mprog.constaint_value_change(81, GIB_Offset_value - 334.65, 0)
mprog.constaint_value_change(96, GIB_Offset_value - 334.65, 1)
mprog.constaint_value_change(102, GIB_Offset_value - 334.65, 1)
mprog.constaint_value_change(99, GIB_Offset_value - 334.65, 0)
mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
mprog.constaint_value_change(195, 1892.5 + 150, 0)  # 大齒輪與FRAME20
mprog.constaint_value_change(176, -1700, 1)  # 右氣壓缸
mprog.constaint_value_change(180, -1700, 1)  # 左氣壓缸
mprog.update()
mprog.switch_window()  # 切換頁面
draft.change_Drawing_scale(1 / scale)  # 圖面比例
draft.exploded_Drawing_1(scale)  # 爆炸圖1

# 計算爆炸圖比例及位置
# drafting_Coordinate_Position, scale = draft.drafting_parameter_calculation(A[i], B[i], H[i])  # 計算爆炸圖比例及位置
# draft.change_Drawing_scale(1 / scale)  # 圖面比例
# draft.Front_View_Drawing(drafting_Coordinate_Position['Front View'][0], drafting_Coordinate_Position['Front View'][1], scale, type)
# draft.Left_View_Drawing(drafting_Coordinate_Position['Left View'][0], drafting_Coordinate_Position['Left View'][1], scale, type)
# draft.Right_View_Drawing(drafting_Coordinate_Position['Right View'][0], drafting_Coordinate_Position['Right View'][1], scale, type)

# mprog.OPEN_Drawing_window()#開啟3D圖視窗


# mprog.update()
# mprog.hide_ass_all_Constraint()