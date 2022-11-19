
import main_program as mprog


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
i = 7

# 離合器跟FRAME3組立
# mprog.add_offset_product_assembly('CRANK_SHAFT_All.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'XY.PLANE', 0)
# mprog.add_offset_product_assembly('CRANK_SHAFT_All.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'XZ.PLANE', 0)
# mprog.add_offset_product_assembly('CRANK_SHAFT_All.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'YZ.PLANE', 0)
# # 氣壓缸跟FRAME3組立
# mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',  32, 'XY.PLANE',1)
# mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -460, 'XZ.PLANE', 0)
# mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -260, 'YZ.PLANE', 0)
# mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',  -32, 'XY.PLANE',1)
# mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -460, 'XZ.PLANE', 1)
# mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 260, 'YZ.PLANE', 1)
# #離合器與FRAME3結合
# mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XY.PLANE',0)
# mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE', 1)
# mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 510, 'YZ.PLANE', 1)

# for y in file_name_list:
#     # print('save' + str(y))
#     mprog.import_part("C:\\Users\\USER\\Desktop\\stamping_press", y)
#     # print('open' + str(y))
#     if i <= 3:
#         if y == 'BALANCER_LEFT_All' or y == 'BALANCER_RIGHT_ALL' or y == 'CRANK_SHAFT_All' or  y == 'CLUCTH_ASSEMBLY_All' or y == 'SLIDE_UNIT_All':
#             mprog.axis_system()
#             mprog.scaling(0.5)
#             if y == 'FRAME1' or y == 'FRAME2' or y == 'FRAME20' or y == 'FRAME30':  # 更改零件變數H
#                 mprog.param_change(y, 'H', H[i])
#                 mprog.save_file_part(path, y)
#             elif y == 'FRAME3' or y == 'FRAME4' or y == 'FRAME9' or y == 'FRAME32' or y == 'FRAME41' or y == 'FRAME43':  # 更改零件變數R
#                 mprog.param_change(y, 'R', R[i])
#                 mprog.save_file_part(path, y)
#             elif y == 'FRAME10' or y == 'FRAME11' or y == 'FRAME12' or y == 'FRAME13' or y == 'BOLSTER1':  # 更改零件變數E
#                 mprog.param_change(y, 'E', E[i])
#                 if y == 'BOLSTER1':
#                     mprog.param_change('BOLSTER1', "hole_type", hole_type[j])
#                 mprog.save_file_part(path, y)
#             elif y == 'FRAME29' or y == 'FRAME8' or y == 'FRAME5':  # 更改零件變數A
#                 mprog.param_change(y, 'A', A[i])
#                 mprog.save_file_part(path, y)
#             elif y == 'SLIDE':  # 更改零件變數P
#                 mprog.param_change(y, 'P', P[i])
#                 mprog.save_file_part(path, y)
#             elif y == 'BOLSTER3':  # 更改零件變數Q
#                 mprog.param_change(y, 'Q', Q[i])
#                 mprog.save_file_part(path, y)
#             else:
#                 mprog.save_file_part(path, y)
#         else:
#             if y == 'FRAME1' or y == 'FRAME2' or y == 'FRAME20' or y == 'FRAME30':  # 更改零件變數H
#                 mprog.param_change(y, 'H', H[i])
#                 mprog.save_file_part(path, y)
#             elif y == 'FRAME3' or y == 'FRAME4' or y == 'FRAME9' or y == 'FRAME32' or y == 'FRAME41' or y == 'FRAME43':  # 更改零件變數R
#                 mprog.param_change(y, 'R', R[i])
#                 mprog.save_file_part(path, y)
#             elif y == 'FRAME10' or y == 'FRAME11' or y == 'FRAME12' or y == 'FRAME13' or y == 'BOLSTER1':  # 更改零件變數E
#                 mprog.param_change(y, 'E', E[i])
#                 if y == 'BOLSTER1':
#                     mprog.param_change('BOLSTER1', "hole_type", hole_type[j])
#                 mprog.save_file_part(path, y)
#             elif y == 'FRAME29' or y == 'FRAME8' or y == 'FRAME5':  # 更改零件變數A
#                 mprog.param_change(y, 'A', A[i])
#                 mprog.save_file_part(path, y)
#             elif y == 'BOLSTER2' or y == 'SLIDE':  # 更改零件變數P
#                 mprog.param_change(y, 'P', P[i])
#                 mprog.save_file_part(path, y)
#             elif y == 'BOLSTER3':  # 更改零件變數Q
#                 mprog.param_change(y, 'Q', Q[i])
#                 mprog.save_file_part(path, y)
#             else:
#                 mprog.save_file_part(path, y)
#     else:
#         if y == 'FRAME1' or y == 'FRAME2' or y == 'FRAME20' or y == 'FRAME30' or y == 'BALANCER1':  # 更改零件變數H
#             mprog.param_change(y, 'H', H[i])
#             mprog.save_file_part(path, y)
#         elif y == 'FRAME3' or y == 'FRAME4' or y == 'FRAME9' or y == 'FRAME32' or y == 'FRAME41' or y == 'FRAME43':  # 更改零件變數R
#             mprog.param_change(y, 'R', R[i])
#             mprog.save_file_part(path, y)
#         elif y == 'FRAME10' or y == 'FRAME11' or y == 'FRAME12' or y == 'FRAME13' or y == 'BOLSTER1':  # 更改零件變數E
#             mprog.param_change(y, 'E', E[i])
#             if y == 'BOLSTER1':
#                 mprog.param_change('BOLSTER1', "hole_type", hole_type[j])
#             mprog.save_file_part(path, y)
#         elif y == 'FRAME29' or y == 'FRAME8' or y == 'FRAME5':  # 更改零件變數A
#             mprog.param_change(y, 'A', A[i])
#             mprog.save_file_part(path, y)
#         elif y == 'BOLSTER2' or y == 'SLIDE':  # 更改零件變數P
#             mprog.param_change(y, 'P', P[i])
#             mprog.save_file_part(path, y)
#         elif y == 'BOLSTER3':  # 更改零件變數Q
#             mprog.param_change(y, 'Q', Q[i])
#             mprog.save_file_part(path, y)
#         elif y == 'CLUCTH_ASSEMBLY36':  # 更改零件變數O
#             mprog.param_change(y, 'O', O[i])
#             mprog.save_file_part(path, y)
#         elif y == 'BEARING_HOUSING2':  # 更改零件變數S
#             mprog.param_change(y, 'S', S[i])
#             mprog.save_file_part(path, y)
#         else:
#             mprog.save_file_part(path, y)

# file_name_list = ['BALANCER1', 'BALANCER10', 'BALANCER2', 'BALANCER3', 'BALANCER4',
#                       'BALANCER5', 'BALANCER6', 'BALANCER7', 'BALANCER8', 'BALANCER9', 'BALL CUP',
#                       'BEARING_HOUSING1', 'BEARING_HOUSING2', 'BEARING_HOUSING3', 'BEARING_HOUSING4',
#                       'BEARING_HOUSING5',
#                       'BOLSTER1',
#                       'BOLSTER2', 'BOLSTER3', 'BOLSTER4', 'BOLSTER5', 'BUSH', 'CLUCTH_ASSEMBLY1', 'CLUCTH_ASSEMBLY10',
#                       'Fixture', 'FLYWHEEL BRAKE', 'FRAME1', 'FRAME10', 'FRAME11',
#                       'FRAME12', 'FRAME13', 'FRAME14', 'FRAME15', 'FRAME16', 'FRAME17', 'FRAME18', 'FRAME19', 'FRAME2',
#                       'FRAME20',
#                       'FRAME21', 'FRAME22', 'FRAME23', 'FRAME24', 'FRAME25', 'FRAME26', 'FRAME27', 'FRAME28', 'FRAME29',
#                       'FRAME3',
#                       'FRAME30', 'FRAME31', 'FRAME32', 'FRAME33', 'FRAME34', 'FRAME35', 'FRAME36', 'FRAME37', 'FRAME38',
#                       'FRAME39',
#                       'FRAME4', 'FRAME40', 'FRAME41', 'FRAME43', 'FRAME5', 'FRAME6', 'FRAME7', 'FRAME8', 'FRAME9',
#                       'GIB1',
#                       'GIB2',
#                       'HOLD MOUNTING MANIFOLD', 'inside the punch 9', 'JOINT1', 'JOINT10', 'JOINT2', 'JOINT3', 'JOINT4',
#                       'JOINT5',
#                       'JOINT6', 'JOINT7', 'JOINT8', 'JOINT9', 'MAIN_GEAR1', 'MAIN_GEAR2', 'MAIN_GEAR3', 'MAIN_GEAR4',
#                       'OVERLOAD PROTECTER', 'Pin', 'POINT', 'POINT2', 'POINT3', 'PRESSURE_PLATE', 'SLIDE', 'SPROCKET',
#                       'washer left', 'washer right', 'WORM GEAR COVER', 'WORM GEAR', 'WORM WHEEL COVER', 'WORM WHEEL',
#                       'BALANCER_LEFT_All', 'BALANCER_RIGHT_ALL', 'CRANK_SHAFT_All', 'CLUCTH_ASSEMBLY_All']

# 開啟組合檔
# file_product_name_list = ['SLIDE_UNIT']
# for x in file_product_name_list:
#     mprog.import_product("C:\\Users\\USER\\Desktop\\stamping_press", x)
#     # print('open' + str(x))
#     mprog.save_file_product(path, x)

# 開啟組合檔
# file_product_name_list = ['SLIDE_UNIT']
# for x in file_product_name_list:
#     mprog.import_product(path, x)
#     # print('open' + str(x))
#     mprog.save_file_product(path, x)

# 匯入待組合組合檔
# file_Assembly_name_list = ['SLIDE_UNIT', 'BALANCER_RIGHT', 'BALANCER_LEFT', 'CRANK_SHAFT', 'CLUCTH_ASSEMBLY']
# for x in file_Assembly_name_list:
#     # print('import' + str(x) + 'product')
#     mprog.import_file_Product(path, x)

# 平板2, 3組立
mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0)
mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 85, 'XY.PLANE', 1)
mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0)





# 機架組合
mprog.base_lock('BOLSTER1.1', 'BOLSTER1.1')  # 基準零件(定海神針)
# (0表示SAME, 1表示OPPOSITE)
##平板-四底座
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', A_15[i] / 2, 'XZ.PLANE', 1)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', A[i] / 2, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -Z[i], 'XY.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - B[i] + F[i] / 2, 'YZ.PLANE', 1)
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -A_15[i] / 2, 'XZ.PLANE', 0)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -A[i] / 2, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -Z[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - B[i] + F[i] / 2, 'YZ.PLANE', 1)
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -A_15[i] / 2, 'XZ.PLANE', 0)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -A[i] / 2, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -Z[i], 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -B[i], 'YZ.PLANE', 1)
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', A_15[i] / 2, 'XZ.PLANE', 0)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', A[i] / 2, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', -Z[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -B[i], 'YZ.PLANE', 0)
# 左右側板
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R_15[i] / 2 + 140, 'XZ.PLANE', 1)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R[i] / 2 + 140, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', H[i] / 2, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', B[i] / 2, 'YZ.PLANE', 0)
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -R_15[i] / 2 - 140, 'XZ.PLANE', 1)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -R[i] / 2 - 140, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -H[i] / 2, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -B[i] / 2, 'YZ.PLANE', 1)
# 底部前、中板
if l == 0:
    mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', A_15[i] / 2, 'XZ.PLANE', 1)
else:
    mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', A[i] / 2, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'YZ.PLANE', 1)
mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -FRAME2_lower_depth[i] + 5, 'YZ.PLANE', 0)
# 中間左右側板
if l == 0:
    mprog.add_offset_assembly('FRAME11.1', 'BOLSTER1.1', -R_15[i] / 2, 'XZ.PLANE', 0)
else:
    mprog.add_offset_assembly('FRAME11.1', 'BOLSTER1.1', -R[i] / 2, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME11.1', -T[i], 'XY.PLANE', 0)
if l == 0:
    mprog.add_offset_assembly('FRAME8.1', 'FRAME11.1', FRAME2_lower_depth[i] * 1.5 - 3.984 + 7.968, 'YZ.PLANE', 1)
else:
    mprog.add_offset_assembly('FRAME8.1', 'FRAME11.1', FRAME2_lower_depth[i] - 3.984 + 7.968, 'YZ.PLANE', 1)
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 0)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R[i] / 2 - 90, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -T[i], 'XY.PLANE', 0)
if l == 0:
    mprog.add_offset_assembly('FRAME5.1', 'FRAME10.1', -FRAME2_lower_depth[i] * 1.5 + 3.984 - 7.968, 'YZ.PLANE', 0)
else:
    mprog.add_offset_assembly('FRAME5.1', 'FRAME10.1', -FRAME2_lower_depth[i] + 3.984 - 7.968, 'YZ.PLANE', 0)
# 底部後面ㄇ形角鐵
mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', B[i], 'YZ.PLANE', 0)
# 平板底部零件
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -R_15[i] / 2, 'XZ.PLANE', 1)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -R[i] / 2, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -T[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', 0, 'YZ.PLANE', 1)
# 平板底部零件
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', R_15[i] / 2, 'XZ.PLANE', 1)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', R[i] / 2, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', -T[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', 0, 'YZ.PLANE', 1)
# 前中上軸承板&角鐵
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME20.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', -H[i] + 40, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', 0, 'YZ.PLANE', 0)
mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -5, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -550, 'YZ.PLANE', 0)
mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', -FRAME20_H[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'YZ.PLANE', 1)
# 氣壓缸鎖固板左右
mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', 242, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME3.1', 'FRAME21.1', -H[i] + 40, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME20.1', 'FRAME21.1', 280, 'YZ.PLANE', 1)
mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', -242, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME3.1', 'FRAME22.1', -H[i] + 40, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME20.1', 'FRAME22.1', 280, 'YZ.PLANE', 1)
# 鎖固平板六兄弟
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', -T[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME3.1', 'FRAME14.1', -75, 'YZ.PLANE', 1)
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', R[i] / 2 + 75 + 140, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -T[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -332.5 + 1 + 17.5, 'YZ.PLANE', 1)
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', R[i] / 2 + 75 + 140, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', -T[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', 332.5 - 37.5 + 7.5, 'YZ.PLANE', 1)
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -R[i] / 2 - 75 - 140, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -T[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', 332.5 - 37.5 + 7.5, 'YZ.PLANE', 0)
if l == 0:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1)
else:
    mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -R[i] / 2 - 75 - 140, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -T[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -332.5 + 117.5, 'YZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', -T[i], 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME3.1', 'FRAME17.1', FRAME1_cutout_bottom[i] - 37, 'YZ.PLANE', 0)
# 左右側板前GIB
mprog.add_offset_assembly('GIB1.1', 'FRAME5.1', FRAME1_lower_high[i] + 40, 'XY.PLANE', 0)
mprog.add_offset_assembly('GIB1.1', 'FRAME1.1', 72.5, 'XZ.PLANE', 0)
mprog.add_offset_assembly('GIB1.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0)
mprog.add_offset_assembly('GIB2.1', 'FRAME8.1', FRAME1_lower_high[i] + 40, 'XY.PLANE', 1)
mprog.add_offset_assembly('GIB2.1', 'FRAME2.1', -72.5, 'XZ.PLANE', 0)
mprog.add_offset_assembly('GIB2.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0)
# 左GIB後鎖固用方塊
mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -690, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME2.1', 'FRAME23.1', 50 + 35, 'XZ.PLANE', 1)
mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', 0, 'YZ.PLANE', 0)
mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME2.1', 'FRAME24.1', 50 + 35, 'XZ.PLANE', 0)
mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -690 / 2, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME2.1', 'FRAME27.1', 50, 'XZ.PLANE', 1)
mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', 0, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -690, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -50, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -130.35, 'YZ.PLANE', 0)
mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -50, 'YZ.PLANE', 0)
mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -62.5, 'XZ.PLANE', 0)
# 右GIB後鎖固用方塊
mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -690, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME1.1', 'FRAME25.1', -50 - 35, 'XZ.PLANE', 0)
mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', 0, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME1.1', 'FRAME26.1', -50 - 35, 'XZ.PLANE', 0)
mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'YZ.PLANE', 0)
mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -690 / 2, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME1.1', 'FRAME28.1', -50, 'XZ.PLANE', 0)
mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', 0, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -690, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', -50, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', 130.35, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB1.1', 'FRAME31.4', -150, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'YZ.PLANE', 0)
mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'XZ.PLANE', 0)
# 中間與平板鎖固方塊下板子
mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', -48, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'YZ.PLANE', 1)
# 後面安裝馬達下板子
mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', H[i] / 2, 'XY.PLANE', 0)  # 要找他所有變數
mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'YZ.PLANE', 0)
# FRAME2中間下板子
mprog.add_offset_assembly('FRAME32.1', 'BOLSTER1.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME32.1', 'FRAME30.1', -1023, 'XY.PLANE', 0)  # 要找他所有變數
mprog.add_offset_assembly('FRAME32.1', 'FRAME30.1', 0, 'YZ.PLANE', 0)
# 後方大軸承支架
if l == 0:
    mprog.add_offset_assembly('FRAME34.1', 'BOLSTER1.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 1)
else:
    mprog.add_offset_assembly('FRAME34.1', 'BOLSTER1.1', -R[i] / 2 - 90, 'XZ.PLANE', 1)
if l == 0:
    mprog.add_offset_assembly('FRAME34.1', 'FRAME3.1', -2029, 'XY.PLANE', 1)  # 要找他所有變數
else:
    mprog.add_offset_assembly('FRAME34.1', 'FRAME3.1', -2390.5, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME34.1', 'FRAME7.1', 1141, 'YZ.PLANE', 0)  # 要找他所有變數
if l == 0:
    mprog.add_offset_assembly('FRAME34.2', 'BOLSTER1.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 0)
else:
    mprog.add_offset_assembly('FRAME34.2', 'BOLSTER1.1', -R[i] / 2 - 90, 'XZ.PLANE', 0)
if l == 0:
    mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', 2029, 'XY.PLANE', 0)  # 要找他所有變數
else:
    mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', 2390.5, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME34.2', 'FRAME7.1', 1141, 'YZ.PLANE', 0)  # 要找他所有變數
# 後方大軸承
mprog.add_offset_assembly('FRAME35.1', 'FRAME3.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -868, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -FRAME20_FRAME2_YZ[i] - 32, 'YZ.PLANE', 0)
# 後方馬達下支撐板
mprog.add_offset_assembly('FRAME39.1', 'FRAME1.1', 50, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME39.1', 'FRAME7.1', 2590, 'XY.PLANE', 1)  # 要找他所有變數
mprog.add_offset_assembly('FRAME39.1', 'FRAME7.1', 320, 'YZ.PLANE', 0)  # 要找他所有變數
mprog.add_offset_assembly('FRAME38.1', 'FRAME2.1', -50, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME38.1', 'FRAME6.1', -2590, 'XY.PLANE', 1)  # 要找他所有變數
mprog.add_offset_assembly('FRAME38.1', 'FRAME6.1', -320, 'YZ.PLANE', 1)  # 要找他所有變數
mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 19, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 125, 'YZ.PLANE', 1)
mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', 19, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', -125, 'YZ.PLANE', 1)
mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', 0, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', -125, 'YZ.PLANE', 1)
mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 0, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 125, 'YZ.PLANE', 1)
# FRAME2右邊角鐵
mprog.add_offset_assembly('FRAME33.1', 'FRAME2.1', 140, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME33.1', 'FRAME15.1', -708, 'XY.PLANE', 0)  # 要找他所有變數
mprog.add_offset_assembly('FRAME33.1', 'FRAME11.1', 313.984, 'YZ.PLANE', 0)  # 要找他所有變數
# FRAME2中間半圓形零件
mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', 130.5, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', -320, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', FRAME20_FRAME2_YZ[i], 'YZ.PLANE', 1)
# FRAME35上兩圓管
mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', 88, 'YZ.PLANE', 1)
mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 272.236, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', -272.236, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 88, 'YZ.PLANE', 1)
# 後方馬達下支撐板上治具
mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 95, 'XZ.PLANE', 0)
mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 22 + 18, 'XY.PLANE', 0)
mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 312.5, 'YZ.PLANE', 0)
# 馬達下方板與後方橫板
mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', 71.5, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', -19, 'YZ.PLANE', 1)
# 平板2, 3組立
mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0)
# 平板1, 3組立
mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0)
if k == 0:
    if h == 0:
        mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', D_S[i], 'XY.PLANE', 0)
    elif h == 1:
        mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', D_H[i], 'XY.PLANE', 0)
    else:
        mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', D_P[i], 'XY.PLANE', 0)
else:
    if h == 0:
        mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', D_S[i] + D_S[i], 'XY.PLANE', 0)
    elif  h == 1:
        mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', D_H[i] + D_H[i], 'XY.PLANE', 0)
    else:
        mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', D_P[i] + D_P[i], 'XY.PLANE', 0)
#SLIDE跟平板3組立
mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 85, 'XY.PLANE', 1)
mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0)
# 離合器跟FRAME3組立
mprog.add_offset_product_assembly('CRANK_SHAFT_All.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'XY.PLANE', 0)
mprog.add_offset_product_assembly('CRANK_SHAFT_All.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_product_assembly('CRANK_SHAFT_All.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'YZ.PLANE', 0)
# 氣壓缸跟FRAME3組立
mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE', 1)
if l == 0:
    mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -688, 'XZ.PLANE', 0)
else:
    mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -460, 'XZ.PLANE', 0)
mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -260, 'YZ.PLANE', 0)
mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE', 1)
if l == 0:
    mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -688, 'XZ.PLANE', 1)
else:
    mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -460, 'XZ.PLANE', 1)
mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 260, 'YZ.PLANE', 1)
# 離合器與FRAME3結合
mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XY.PLANE', 0)
mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 510, 'YZ.PLANE', 1)

# 更新
mprog.update()
mprog.hide_ass_all_Constraint()