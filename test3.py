import main_program as mprog
Q = [250, 300, 350, 400, 460, 520, 580, 650, 720]
Q_15 = [375 , 450 , 525 , 600 , 690 , 780 , 870 , 975 , 1080]

A = [720, 830, 890, 940, 1050, 1160, 1300, 1480, 1560]
A_15 = [1080 , 1245 , 1335 , 1410 , 1575 , 1740 , 1950 , 2220 , 2340]
B = [1058, 1125, 1210, 1315, 1480, 1680, 1985, 2113, 2400]
B_15 =[1587 , 1688 , 1815 , 1973 , 2220 , 2520 , 2978 , 3170 , 3600]
H = [2060, 2185, 2290, 2540, 2755, 2990, 3270, 3725, 4005]
R = [388, 486, 516, 544, 614, 670, 730, 900, 970]
R_15 = [582 , 729 , 774 , 816 , 921 , 1005 , 1095 , 1350 , 1455]
E = [700, 780, 840, 900, 1050, 1150, 1250, 1400, 1500]
E_15 = [1050 , 1170 , 1260 , 1350 , 1575 , 1725 , 1875 , 2100 , 2250]
D_DH = [250, 280, 330, 350, 380, 430, 490, 550, 580]
D_S = [80, 90, 110, 130, 150, 180, 200, 220, 250]
D_H = [50, 60, 70, 80, 100, 110, 130, 150, 180]
D_P = [35, 40, 45, 50, 60, 70, 80, 90, 100]
DH_S = [230, 250, 270, 300, 330, 350, 400, 450, 450]
DH_H = [200, 220, 240, 270, 300, 320, 360, 400, 400]
DH_P = [200, 220, 240, 270, 300, 320, 360, 400, 400]
S = [983, 1068, 1158, 1285, 1445, 1630, 1809, 2067, 2262]
H_Z = [1260, 1385, 1490, 1640, 1855, 2086.933, 2370, 2725, 3005]
O = [1045, 1075, 1125, 1145, 1175, 1225, 1285, 1345, 1375]
hole_type = [0, 1, 2]
P = [330, 380, 430, 480, 560, 650, 720, 860, 960]
P_15 = [495 , 570 , 645 , 720 , 840 , 975 , 1080 , 1290 , 1440]
Q = [250, 300, 350, 400, 460, 520, 580, 650, 720]
Q_15 = [375 , 450 , 525 , 600 , 690 , 780 , 870 , 975 , 1080]
T = [85, 100, 115, 130, 140, 155, 165, 180, 180]
Z = [800, 800, 800, 900, 900, 900, 900, 1000, 1000]
FRAME1_cutout_bottom = [197, 277, 317, 397, 477, 557, 637, 717, 797]
F = [320, 400, 440, 520, 600, 680, 760, 840, 900]
FRAME1_cutout = [655 + 370, 675 + 415, 695 + 450, 715 + 577.5, 735 + 670, 755 + 770, 775 + 895, 795 + 1080, 815 + 1210]
FRAME20_H = [280, 355, 420, 437.5, 550, 680, 825, 990, 1120]
# FRAME2_lower_depth = [165, 245, 285, 365, 445, 525, 605, 685, 745]
FRAME2_lower_depth = [166.016, 246.016, 286.016, 366.016, 446.016, 526.016, 606.016, 686.016, 766.016]
FRAME2_lower_depth_15 = [265.016, 385.016, 445.016, 565.016, 685.016, 805.016, 925.016, 1045.016, 1165.016]
FRAME1_lower_high = [1330, 1335, 1340, 1445, 1455, 1460, 1470, 1575, 1595]
FRAME20_FRAME2_YZ = [805, 805, 979, 979, 979, 979, 979, 979, 979]
BALANCER1_XZ = [204, 253, 268, 282, 317, 345, 375, 460, 575]  # (R+180mm)/2+80mm
FRAME_10_H = [558, 620.5, 673, 798, 905.5, 1023, 1163, 1320.5, 1530.5]

E = [700, 780, 840, 900, 1050, 1150, 1250, 1400, 1500]
E_15 = [1051 , 1170 , 1260 , 1350 , 1575 , 1725 , 1875 , 2100 , 2250]

B = [1058, 1125, 1210, 1315, 1480, 1680, 1985, 2113, 2400]
B_15 =[1587 , 1688 , 1815 , 1973 , 2220 , 2520 , 2978 , 3170 , 3600]
i = 5
# mprog.folder_file_name()
# x = ['back straight', 'BALANCER', 'BALANCER1', 'BALANCER10', 'BALANCER2', 'BALANCER3', 'BALANCER4', 'BALANCER4',
#      'BALANCER5', 'BALANCER6', 'BALANCER7', 'BALANCER8', 'BALANCER9', 'BALL CUP',
#      'BEARING_HOUSING1', 'BEARING_HOUSING2', 'BEARING_HOUSING3', 'BEARING_HOUSING4', 'BEARING_HOUSING5', 'BOLSTER1',
#      'BOLSTER2', 'BOLSTER3', 'BOLSTER4', 'BOLSTER5', 'BUSH', 'CLUCTH_ASSEMBLY1', 'CLUCTH_ASSEMBLY10',
#      'CLUCTH_ASSEMBLY11', 'CLUCTH_ASSEMBLY12', 'CLUCTH_ASSEMBLY13', 'CLUCTH_ASSEMBLY14', 'CLUCTH_ASSEMBLY15',
#      'CLUCTH_ASSEMBLY16', 'CLUCTH_ASSEMBLY17', 'CLUCTH_ASSEMBLY18', 'CLUCTH_ASSEMBLY19', 'CLUCTH_ASSEMBLY2',
#      'CLUCTH_ASSEMBLY20', 'CLUCTH_ASSEMBLY21', 'CLUCTH_ASSEMBLY22', 'CLUCTH_ASSEMBLY23', 'CLUCTH_ASSEMBLY24',
#      'CLUCTH_ASSEMBLY25', 'CLUCTH_ASSEMBLY26', 'CLUCTH_ASSEMBLY27', 'CLUCTH_ASSEMBLY28', 'CLUCTH_ASSEMBLY29',
#      'CLUCTH_ASSEMBLY3', 'CLUCTH_ASSEMBLY30', 'CLUCTH_ASSEMBLY31', 'CLUCTH_ASSEMBLY32', 'CLUCTH_ASSEMBLY33',
#      'CLUCTH_ASSEMBLY34', 'CLUCTH_ASSEMBLY35', 'CLUCTH_ASSEMBLY36', 'CLUCTH_ASSEMBLY4', 'CLUCTH_ASSEMBLY5',
#      'CLUCTH_ASSEMBLY6', 'CLUCTH_ASSEMBLY7', 'CLUCTH_ASSEMBLY8', 'CLUCTH_ASSEMBLY9', 'COUNTER_PARTS',
#      'COVER (BALL SCREW)', 'CRANK_SHAFT', 'Fixture', 'FLYWHEEL BRAKE', 'FRAME1', 'FRAME10', 'FRAME11',
#      'FRAME12', 'FRAME13', 'FRAME14', 'FRAME15', 'FRAME16', 'FRAME17', 'FRAME18', 'FRAME19', 'FRAME2', 'FRAME20',
#      'FRAME21', 'FRAME22', 'FRAME23', 'FRAME24', 'FRAME25', 'FRAME26', 'FRAME27', 'FRAME28', 'FRAME29', 'FRAME3',
#      'FRAME30', 'FRAME31', 'FRAME32', 'FRAME33', 'FRAME34', 'FRAME35', 'FRAME36', 'FRAME37', 'FRAME38', 'FRAME39',
#      'FRAME4', 'FRAME40', 'FRAME41', 'FRAME43', 'FRAME5', 'FRAME6', 'FRAME7', 'FRAME8', 'FRAME9', 'GIB1', 'GIB2',
#      'HOLD MOUNTING MANIFOLD', 'inside the punch 9', 'JOINT1', 'JOINT10', 'JOINT2', 'JOINT3', 'JOINT4', 'JOINT5',
#      'JOINT6', 'JOINT7', 'JOINT8', 'JOINT9', 'MAIN_GEAR1', 'MAIN_GEAR2', 'MAIN_GEAR3', 'MAIN_GEAR4',
#      'OVERLOAD PROTECTER', 'Pin', 'POINT', 'POINT2', 'POINT3', 'PRESSURE_PLATE', 'SLIDE', 'SPROCKET',
#      'washer left', 'washer right', 'WORM GEAR COVER', 'WORM GEAR', 'WORM WHEEL COVER', 'WORM WHEEL']
# y = ['BALANCER_LEFT', 'BALANCER_RIGHT', 'CLUCTH_ASSEMBLY', 'CRANK_SHAFT', 'SLIDE_UNIT']
l = (0)
# mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE', 1)
# if l == 0:
#     mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',  -R_15[i] / 2 - 13, 'XZ.PLANE', 0)
# else:
#     mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',  -R[i] / 2 - 13, 'XZ.PLANE', 0)
# mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -260, 'YZ.PLANE', 0)
# mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE', 1)
# if l == 0:
#     mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',  -R_15[i] / 2 - 13, 'XZ.PLANE', 1)
# else:
#     mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -R[i] / 2 - 13, 'XZ.PLANE', 1)
# mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 260, 'YZ.PLANE', 1)

# if l == 0:
#     mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', R_15[i] / 2, 'XZ.PLANE', 0)
# else:
#     mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', R[i] / 2, 'XZ.PLANE', 0)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME11.1', -T[i], 'XY.PLANE', 1)
# if l == 0:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', FRAME2_lower_depth_15[i], 'YZ.PLANE', 1)
# else:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', FRAME2_lower_depth[i], 'YZ.PLANE', 1)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 0)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R[i] / 2 - 90, 'XZ.PLANE', 0)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -T[i], 'XY.PLANE', 0)
# if l == 0:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', FRAME2_lower_depth_15[i], 'YZ.PLANE', 1)
# else:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', FRAME2_lower_depth[i], 'YZ.PLANE', 1)
#
# mprog.add_offset_assembly('FRAME10.1', 'FRAME34.2', -485.543, 'YZ.PLANE', 0)
# mprog.add_offset_assembly('FRAME34.2', 'FRAME10.1', FRAME_10_H[i] + 40, 'XY.PLANE', 1)
# if l == 0:
#     mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -R_15[i] / 2 - 90, 'XZ.PLANE', 0)
# else:
#     mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -R[i] / 2 - 90, 'XZ.PLANE', 0)
# mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -R_15[i] - 180, 'XZ.PLANE', 1)
# mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'XY.PLANE', 1)
# mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'YZ.PLANE', 0)

mprog.axis_system()
mprog.scaling()