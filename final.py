import os
import win32com.client as win32
import main_program as mprog

#開啟CATIA
env = mprog.set_CATIA_workbench_env()
# env.Generative_Sheetmetal_Design()

# #開啟product檔
mprog.assembly_create()

# 匯入組合檔
file_name_list = ['SLIDE_UNIT', 'FRAME_BALANCER', 'CRANK_SHAFT', 'CLUCTH_ASSEMBLY']
for x in file_name_list:
    mprog.import_file("C:\\Users\\USER\\Desktop\\stamping_press",x)

#各型號參數對應尺寸
A = [720, 830, 890, 940, 1050, 1160, 1300, 1480, 1560]
B = [1058, 1125, 1210, 1315, 1480, 1680, 1985, 2113, 2400]
H = [2060, 2185, 2290, 2540, 2755, 2990, 3270, 3725, 4005]
R = [388, 486, 516, 544, 614, 670, 730, 900, 970]
E = [700, 780, 840, 900, 1050, 1150, 1250, 1400, 1500]
D_DH = [250, 280, 330, 350, 380, 430, 490, 550, 580]
S = [983, 1068, 1158, 1285, 1445, 1630, 1809, 2067, 2262]
H_Z = [1260, 1385, 1490, 1640, 1855, 2086.933, 2370, 2725, 3005]
O = [1045, 1075, 1125, 1145, 1175, 1225, 1285, 1345, 1375]

#FRAME變數H板件修改
product_file_name = ['FRAME1', 'FRAME2', 'FRAME20', 'FRAME30']
for x in product_file_name:
    mprog.param_change(x, 'H', H[7])
# #FRAME變數R板件修改
product_file_name = ['FRAME3', 'FRAME4', 'FRAME9', 'FRAME43']
for x in product_file_name:
    mprog.param_change(x, 'R', R[7])
#FRAME變數E板件修改
product_file_name = ['FRAME10', 'FRAME11']
for x in product_file_name:
    mprog.param_change(x, 'E', E[0])

#FRAME組合
# combined_dimension('A', A[0], 'Offset.173')

#組合尺寸-機架_衝頭
product_file_name = ['BOLSTER3']
for x in product_file_name:
    mprog.param_change(x, 'D+DH', D_DH[8])
# #組合尺寸-機架_曲輪軸
product_file_name = ['BEARING_HOUSING2']
for x in product_file_name:
    mprog.param_change(x, 'S', S[8])
# #組合尺寸-機架_離合器組合圖
product_file_name = ['CLUCTH_ASSEMBLY36']
for x in product_file_name:
    mprog.param_change(x, 'O', O[8])
# #組合尺寸-機架_左右平衡氣壓缸
product_file_name = ['BALANCER1']
for x in product_file_name:
    mprog.param_change(x, 'H_Z', H_Z[5])

#組合product-機架, 衝頭
mprog.add_offset_assembly_0('SLIDE_UNIT', 'BOLSTER3', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'XY.PLANE')
mprog.add_offset_assembly_0('SLIDE_UNIT', 'BOLSTER3', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'YZ.PLANE')
mprog.add_offset_assembly_0('SLIDE_UNIT', 'BOLSTER3', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'XZ.PLANE')
# 組合product-機架,曲輪軸
mprog.add_offset_assembly_1('CRANK_SHAFT', 'BEARING_HOUSING2.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'XY.PLANE')
mprog.add_offset_assembly_1('CRANK_SHAFT', 'BEARING_HOUSING2.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'XZ.PLANE')
mprog.add_offset_assembly_1('CRANK_SHAFT', 'BEARING_HOUSING2.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'YZ.PLANE')
#組合product-機架,離合器組立圖
mprog.add_offset_assembly_1('CLUCTH_ASSEMBLY', 'CLUCTH_ASSEMBLY36.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'XY.PLANE')
mprog.add_offset_assembly_0('CLUCTH_ASSEMBLY', 'CLUCTH_ASSEMBLY36.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'YZ.PLANE')
mprog.add_offset_assembly_0('CLUCTH_ASSEMBLY', 'CLUCTH_ASSEMBLY36.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'XZ.PLANE')
#組合product-機架,左平衡氣壓缸
mprog.add_offset_assembly_1('BALANCER_LEFT', 'BALANCER1.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'XY.PLANE')
mprog.add_offset_assembly_1('BALANCER_LEFT', 'BALANCER1.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'YZ.PLANE')
mprog.add_offset_assembly_1('BALANCER_LEFT', 'BALANCER1.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'XZ.PLANE')
#組合product-機架,右平衡氣壓缸
mprog.add_offset_assembly_1('BALANCER_RIGHT', 'BALANCER1.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'XY.PLANE')
mprog.add_offset_assembly_0('BALANCER_RIGHT', 'BALANCER1.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'YZ.PLANE')
mprog.add_offset_assembly_0('BALANCER_RIGHT', 'BALANCER1.1', 'FRAME_BALANCER', 'FRAME', 'BOLSTER1.1', 0, 'XZ.PLANE')


