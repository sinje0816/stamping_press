import os
import win32com.client as win32
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
Z = [800, 800, 800, 900, 900, 900, 900, 1000, 1000]


mprog.base_lock('BOLSTER1.1', 'BOLSTER1.1')

## (0表示SAME, 1表示OPPOSITE)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', A[8]/2, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -Z[8], 'XY.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -2140, 'YZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -A[8]/2, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -Z[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -2140, 'YZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -A[8]/2, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -Z[8], 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME6.1', 'FRAME5.1', -B[8], 'YZ.PLANE', 0)


# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -R[8]/2--140, 'XZ.PLANE', 1)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', R[8]/2--140, 'XY.PLANE', 0)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R[8]/2--140, 'XZ.PLANE', 1)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R[8]/2--140, 'XY.PLANE', 0)