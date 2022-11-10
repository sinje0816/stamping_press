import os
import win32com.client as win32
import main_program as mprog
i = ()
j= ()
type = ()
hole = ()

#電子型錄規格
A = [720, 830, 890, 940, 1050, 1160, 1300, 1480, 1560]
B = [1058, 1125, 1210, 1315, 1480, 1680, 1985, 2113, 2400]
H = [2060, 2185, 2290, 2540, 2755, 2990, 3270, 3725, 4005]
R = [388, 486, 516, 544, 614, 670, 730, 900, 970]
E = [700, 780, 840, 900, 1050, 1150, 1250, 1400, 1500]
D_DH = [250, 280, 330, 350, 380, 430, 490, 550, 580]
S = [983, 1068, 1158, 1285, 1445, 1630, 1809, 2067, 2262]
H_Z = [1260, 1385, 1490, 1640, 1855, 2086.933, 2370, 2725, 3005]
O = [1045, 1075, 1125, 1145, 1175, 1225, 1285, 1345, 1375]
hole_type = [0 , 1 , 2]
P = [330 , 380 , 430 , 480 , 560 , 650 , 720 , 860 , 960]
Q = [250 , 300 , 350 , 400 , 460 , 520 , 580 , 650 , 720]

#新增資料夾
path, dir =mprog.new_Folder()
print(path)

#開啟零件檔
# mprog.open_part()
# catapp = win32.Dispatch('CATIA.Application')

#確認型號
print("輸入型號")
type = input()
if type == "SN1-25" :
    i=0
elif type == "SN1-35":
    i=1
elif type == "SN1-45":
    i=2
elif type == "SN1-60":
    i=3
elif type == "SN1-80":
    i=4
elif type == "SN1-110":
    i=5
elif type == "SN1-160":
    i=6
elif type == "SN1-200":
    i=7
elif type == "SN1-250":
    i=8
print(i)

#輸入平板型號
print("請輸入平板型式(0 = 圓形平板 , 1 = 方形平板 , 2 = 模墊型平板)")
hole = input()
if hole == "0":
    j = 0
elif hole == "1":
    j = 1
elif hole == "2":
    j = 2

# 開啟CATIA
env = mprog.set_CATIA_workbench_env()

#匯入零件檔
file_name_FRAME = ['FRAME1' , 'FRAME2' , 'FRAME3' , 'FRAME4' , 'FRAME9' , 'FRAME10' , 'FRAME11' , 'FRAME12' , 'FRAME13'
    , 'FRAME20' , 'FRAME29' , 'FRAME30' , 'FRAME32' , 'FRAME41' , 'FRAME43' , 'BOLSTER1' , 'BOLSTER2' , 'BOLSTER3'
    , 'SLIDE']
for x in file_name_FRAME:
    mprog.import_part("C:\\Users\\USER\\Desktop\\stamping_press",x)

#更改零件變數H
product_file_name = ['FRAME1', 'FRAME2', 'FRAME20', 'FRAME30']
for x in product_file_name:
    mprog.param_change(x, 'H', H[i])

#更改零件變數R
product_file_name = ['FRAME3' , 'FRAME4' , 'FRAME9' , 'FRAME32' , 'FRAME41' , 'FRAME43']
for x in product_file_name:
    mprog.param_change(x, 'R', R[i])

#更改零件變數E
product_file_name = ['FRAME10' , 'FRAME11' , 'FRAME12' , 'FRAME13' , 'BOLSTER1']
for x in product_file_name:
    mprog.param_change(x, 'E', E[i])

#更改零件變數A
product_file_name = ['FRAME29']
for x in product_file_name:
    mprog.param_change(x, 'A', A[i])

#更改平板變數
mprog.param_change('BOLSTER1' , "hole_type" , hole_type[j])

#更改零件變數P
product_file_name = ['BOLSTER2' , 'SLIDE']
for x in product_file_name:
    mprog.param_change(x, 'P', P[i])

#更改零件變數Q
product_file_name = ['BOLSTER3']
for x in product_file_name:
    mprog.param_change(x, 'Q', Q[i])

#儲存零件並關閉
file_name_FRAME = ['FRAME1' , 'FRAME2' , 'FRAME3' , 'FRAME4' , 'FRAME9' , 'FRAME10' , 'FRAME11' , 'FRAME12' , 'FRAME13'
    , 'FRAME20' , 'FRAME29' , 'FRAME30' , 'FRAME32' , 'FRAME41' , 'FRAME43' , 'BOLSTER1' , 'BOLSTER2' , 'BOLSTER3'
    , 'SLIDE']
for x in file_name_FRAME:
    mprog.save_file(path,x)
