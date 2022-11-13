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
T = [85, 100, 115, 130, 140, 155, 165, 180, 180]
H_Z = [1260, 1385, 1490, 1640, 1855, 2086.933, 2370, 2725, 3005]
O = [1045, 1075, 1125, 1145, 1175, 1225, 1285, 1345, 1375]
Z = [800, 800, 800, 900, 900, 900, 900, 1000, 1000]
FRAME1_cutout_bottom = [197, 277, 317, 397, 477, 557, 637, 717, 797]
F = [320, 400, 440, 520, 600, 680, 760, 840, 900]
FRAME1_cutout = [655+370, 675+415, 695+450, 715+577.5, 735+670, 755+770, 775+895, 795+1080, 815+1210]

#SN1-250
mprog.base_lock('BOLSTER1.1', 'BOLSTER1.1')
# (0表示SAME, 1表示OPPOSITE)

##平板-四底座
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', A[8]/2, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -Z[8], 'XY.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80-B[8]+F[8]/2, 'YZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -A[8]/2, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -Z[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80-B[8]+F[8]/2, 'YZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -A[8]/2, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -Z[8], 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -B[8], 'YZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', A[8]/2, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', -Z[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -B[8], 'YZ.PLANE', 0)
#左右側板
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', R[8]/2+140, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', H[8]/2, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', B[8]/2, 'YZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -R[8]/2-140, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -H[8]/2, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', B[8]/2, 'YZ.PLANE', 0)
#底部前、中板
mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', -A[8]/2, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'YZ.PLANE', 0)
mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -FRAME1_cutout_bottom[8], 'YZ.PLANE', 0)
#中間左右側板
mprog.add_offset_assembly('FRAME11.1', 'BOLSTER1.1', R[8]/2+90, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME11.1', -T[8], 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME8.1', 'FRAME11.1', FRAME1_cutout_bottom[8], 'YZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -R[8]/2-90, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -T[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME5.1', 'FRAME10.1', FRAME1_cutout_bottom[8], 'YZ.PLANE', 1)
#底部後面ㄇ形角鐵
mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', B[8], 'YZ.PLANE', 0)
#平板底部零件
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -R[8]/2, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -T[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', 0, 'YZ.PLANE', 1)
#平板底部零件
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', R[8]/2, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', -T[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', 0, 'YZ.PLANE', 1)
#前中上軸承板&角鐵
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME20.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', -H[8]+40, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', 0, 'YZ.PLANE', 0)
mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -5, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -550, 'YZ.PLANE', 0)
mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', -H[8]+Z[8]-T[8]+FRAME1_cutout[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'YZ.PLANE', 1)
#氣壓缸鎖固板左右
mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', 242, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME3.1', 'FRAME21.1', -H[8]+40, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME20.1', 'FRAME21.1', 280, 'YZ.PLANE', 1)
mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', -242, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME3.1', 'FRAME22.1', -H[8]+40, 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME20.1', 'FRAME22.1', 280, 'YZ.PLANE', 1)
#鎖固平板六兄弟
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', -T[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME3.1', 'FRAME14.1', -75, 'YZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', R[8]/2+75+140, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -T[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -332.5, 'YZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', R[8]/2+75+140, 'XZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', -T[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', 332.5, 'YZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -R[8]/2-75-140, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -T[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', 332.5, 'YZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -R[8]/2-75-140, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -T[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -332.5, 'YZ.PLANE', 0)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', -T[8], 'XY.PLANE', 0)
mprog.add_offset_assembly('FRAME3.1', 'FRAME17.1', FRAME1_cutout_bottom[8], 'YZ.PLANE', 0)
#左右側板前GIB
mprog.add_offset_assembly('GIB1.1', 'FRAME1.1', -367.5, 'XY.PLANE', 1)
mprog.add_offset_assembly('GIB1.1', 'FRAME1.1', 72.5, 'XZ.PLANE', 0)
mprog.add_offset_assembly('GIB1.1', 'FRAME1.1', -865.35, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB2.1', 'FRAME2.1', -367.5, 'XY.PLANE', 1)
mprog.add_offset_assembly('GIB2.1', 'FRAME2.1', -72.5, 'XZ.PLANE', 0)
mprog.add_offset_assembly('GIB2.1', 'FRAME2.1', -865.35, 'YZ.PLANE', 1)
#左GIB後鎖固用方塊
mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -690, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME2.1', 'FRAME23.1', 50+35, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME2.1', 'FRAME23.1', -865.35, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME2.1', 'FRAME24.1', 50+35, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME2.1', 'FRAME24.1', -865.35, 'YZ.PLANE', 0)
mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -690/2, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME2.1', 'FRAME27.1', 50, 'XZ.PLANE', 1)
mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', 0, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -690, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -50, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -130.35, 'YZ.PLANE', 0)
mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -50, 'YZ.PLANE', 0)
mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -130.35, 'XZ.PLANE', 0)
#右GIB後鎖固用方塊
mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -690, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME1.1', 'FRAME25.1', -50-35, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME1.1', 'FRAME25.1', -865.35, 'YZ.PLANE', 0)
mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME1.1', 'FRAME26.1', -50-35, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME1.1', 'FRAME26.1', -865.35, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -690/2, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME1.1', 'FRAME28.1', -50, 'XZ.PLANE', 0)
mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', 0, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -690, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', -50, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', 130.35, 'YZ.PLANE', 1)
mprog.add_offset_assembly('GIB1.1', 'FRAME31.4', -150, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'YZ.PLANE', 0)
mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'XZ.PLANE', 0)
#中間與平板鎖固方塊下板子
mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', -48, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'YZ.PLANE', 1)
#後面安裝馬達下板子
mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', H[8]/2, 'XY.PLANE', 0)#要找他所有變數
mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'YZ.PLANE', 0)
#FRAME2中間下板子
mprog.add_offset_assembly('FRAME32.1', 'BOLSTER1.1', 0, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME32.1', 'FRAME20.1', -1023, 'XY.PLANE', 0)#要找他所有變數
mprog.add_offset_assembly('FRAME32.1', 'FRAME20.1', 0, 'YZ.PLANE', 0)
#後方大軸承支架
mprog.add_offset_assembly('FRAME34.1', 'BOLSTER1.1', -R[8]/2-90, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME34.1', 'FRAME3.1', -2390.5, 'XY.PLANE', 1)#要找他所有變數
mprog.add_offset_assembly('FRAME34.1', 'FRAME7.1', 1141, 'YZ.PLANE', 0)
mprog.add_offset_assembly('FRAME34.2', 'BOLSTER1.1', -R[8]/2-90, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', 2390.5, 'XY.PLANE', 0)#要找他所有變數
mprog.add_offset_assembly('FRAME34.2', 'FRAME7.1', 1141, 'YZ.PLANE', 0)
#後方大軸承
mprog.add_offset_assembly('FRAME35.1', 'FRAME3.1', 0, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME35.1', 'FRAME34.1', 272.236, 'XY.PLANE', 1)#要找他所有變數
mprog.add_offset_assembly('FRAME35.1', 'FRAME3.1', -1500, 'YZ.PLANE', 0)#要找他所有變數
#後方馬達下支撐板
mprog.add_offset_assembly('FRAME39.1', 'FRAME1.1', 50, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME39.1', 'FRAME7.1', 2590, 'XY.PLANE', 1)#要找他所有變數
mprog.add_offset_assembly('FRAME39.1', 'FRAME7.1', 320, 'YZ.PLANE', 0)#要找他所有變數
mprog.add_offset_assembly('FRAME38.1', 'FRAME2.1', -50, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME38.1', 'FRAME6.1', -2590, 'XY.PLANE', 1)#要找他所有變數
mprog.add_offset_assembly('FRAME38.1', 'FRAME6.1', -320, 'YZ.PLANE', 1)#要找他所有變數
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
#FRAME2右邊角鐵
mprog.add_offset_assembly('FRAME33.1', 'FRAME2.1', 140, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME33.1', 'FRAME15.1', -708, 'XY.PLANE', 0)#要找他所有變數
mprog.add_offset_assembly('FRAME33.1', 'FRAME11.1', 313.984, 'YZ.PLANE', 0)#要找他所有變數
#FRAME2中間半圓形零件
mprog.add_offset_assembly('FRAME40.1', 'FRAME2.1', -344.5, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', -320, 'XY.PLANE', 0)#要找他所有變數
mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', 979, 'YZ.PLANE', 1)#要找他所有變數
#FRAME2中間半圓形零件
mprog.add_offset_assembly('FRAME40.1', 'FRAME2.1', -344.5, 'XZ.PLANE', 0)
mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', -320, 'XY.PLANE', 0)#要找他所有變數
mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', 979, 'YZ.PLANE', 1)#要找他所有變數
#FRAME35上兩圓管
mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', 88, 'YZ.PLANE', 1)
mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 272.236, 'XZ.PLANE', 1)
mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', -272.236, 'XY.PLANE', 1)
mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 88, 'YZ.PLANE', 1)

#更新
mprog.update()