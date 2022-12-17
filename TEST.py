import math
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
BALANCER1_XZ = [204, 253, 268, 282, 317, 345, 375, 460, 575]

i = 5
# 計算爆炸圖比例及位置
drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale = draft.drafting_parameter_calculation(A[i],
                                                                                                                   B[i],
                                                                                                                   H[i],
                                                                                                                   S[i],
                                                                                                                   T[i])
# 43.6

scale = 1 / scale
# (將3D圖面尺寸轉換為2D尺寸, 將圖面x座標尺寸轉為R再將R轉為Y)
point_position = {'6': [-4250 * math.cos(math.radians(45)) - (R[i] + 145) / 2 * math.cos(math.radians(45)),
                        (-4250 * math.cos(math.radians(45)) - (R[i] + 145) / 2 * math.cos(math.radians(45))) * math.sin(
                            math.radians(30)) / math.cos(math.radians(30)) - 690 * 1 / 3]}

circle_position = {'6': ['6', point_position['6'][0] - 300,
                         point_position['6'][1] + 300]}

print(point_position)
print(circle_position)

draft.balloons('Isometric view', circle_position['6'][0], circle_position['6'][1], circle_position['6'][2],
               point_position['6'][0], point_position['6'][1], circle_position['6'][1])
