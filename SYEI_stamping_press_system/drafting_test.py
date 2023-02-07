

A = [720, 830, 890, 940, 1050, 1160, 1300, 1480, 1560, 1760]
A_15 = [1080, 1245, 1335, 1410, 1575, 1740, 1950, 2220, 2340, 2640]
B = [1058, 1125, 1210, 1315, 1480, 1680, 1985, 2113, 2400, 2700]
B_15 = [1587, 1688, 1815, 1973, 2220, 2520, 2978, 3170, 3600, 4050]
H = [2060, 2185, 2290, 2540, 2755, 2990, 3270, 3725, 4005, 4285]
circle_code_1_X = 60
circle_code_1_Y = 780

def drafting_parameter_calculation(width , height , depth , scale_p):
#-------------------------------初始參數設定(從圖面左上角建立)----------------------------------
    scale = 1/scale_p
    circle_code_1_center =[circle_code_1_X , circle_code_1_Y]
    w_scale = width*scale
    h_scale = height*scale
    d_scale = depth*scale
#--------------------------------左壁板虛擬方框範圍計算----------------------------
    part1_range_center = [circle_code_1_center[0]+d_scale+20+10 , circle_code_1_center[1]-10-h_scale/2]#左壁板中心範圍[x , y]
    part1_range = [part1_range_center[0]-d_scale-20-10-5 , part1_range_center[1]-h_scale/2-15 ,
                   part1_range_center[0]+20+d_scale+10 , part1_range_center[1]+h_scale/2+15]#左壁板虛擬方框[Xmin , Ymin , Xmax , Ymax]
    while True:
        if part1_range [0] < 20:#防止方框X方向超出圖框
            circle_code_1_center[0] += 5
            part1_range_center[0] = circle_code_1_center[0]+d_scale+20
            part1_range[0] = part1_range_center[0]-d_scale-20-10-5
            part1_range[1] = part1_range_center[0]+20+d_scale+10#利用位移完的結果重新計算方框X範圍
        else:
            break
    print(part1_range_center)
    print(part1_range)
#--------------------------------右壁板虛擬方框計算-------------------------------
    circle_code_2_center = [circle_code_1_center[0] , part1_range[1]-5]
    part2_range_center = [circle_code_2_center[0]+d_scale+20 , circle_code_2_center[1]-10-h_scale/2]#左壁板中心範圍[x , y]








