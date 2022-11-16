import main_program as mprog
path, dir = mprog.new_Folder()
print(path)
A = [720, 830, 890, 940, 1050, 1160, 1300, 1480, 1560]
B = [1058, 1125, 1210, 1315, 1480, 1680, 1985, 2113, 2400]
H = [2060, 2185, 2290, 2540, 2755, 2990, 3270, 3725, 4005]
R = [388, 486, 516, 544, 614, 670, 730, 900, 970]
E = [700, 780, 840, 900, 1050, 1150, 1250, 1400, 1500]
D_DH = [250, 280, 330, 350, 380, 430, 490, 550, 580]
S = [983, 1068, 1158, 1285, 1445, 1630, 1809, 2067, 2262]
HZ = [1260, 1385, 1490, 1640, 1855, 2086.933, 2370, 2725, 3005]
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
FRAME2_lower_depth = [166.016, 246.016, 286.016, 366.016, 446.016, 526.016, 606.016, 686.016, 746.016]

env = mprog.set_CATIA_workbench_env()
y = 'BALANCER1'

print(y)
if y == 'FRAME1' or y == 'FRAME2' or y == 'FRAME20' or y == 'FRAME30':  # 更改零件變數H
    mprog.param_change(y, 'H', H[i])
    mprog.save_file(path, y)
elif y == 'FRAME3' or y == 'FRAME4' or y == 'FRAME9' or y == 'FRAME32' or y == 'FRAME41' or y == 'FRAME43':  # 更改零件變數R
    mprog.param_change(y, 'R', R[i])
    mprog.save_file(path, y)
elif y == 'FRAME10' or y == 'FRAME11' or y == 'FRAME12' or y == 'FRAME13' or y == 'BOLSTER1':  # 更改零件變數E
    mprog.param_change(y, 'E', E[i])
    if y == 'BOLSTER1':
        mprog.param_change('BOLSTER1', "hole_type", hole_type[j])
    mprog.save_file(path, y)
elif y == 'FRAME29':  # 更改零件變數A
    mprog.param_change(y, 'A', A[i])
    mprog.save_file(path, y)
elif y == 'BOLSTER2' or y == 'SLIDE':  # 更改零件變數P
    mprog.param_change(y, 'P', P[i])
    mprog.save_file(path, y)
elif y == 'BOLSTER3':  # 更改零件變數Q
    mprog.param_change(y, 'Q', Q[i])
    mprog.save_file(path, y)
elif y == 'CLUCTH_ASSEMBLY36':  # 更改零件變數O
    mprog.param_change(y, 'O', O[i])
    mprog.save_file(path, y)
elif y == 'BALANCER1':  # 更改零件變數'H_Z'
    mprog.param_change(y, 'HZ', HZ[8])
    # mprog.save_file(path, y)
elif y == 'BEARING_HOUSING2':  # 更改零件變數S
    mprog.param_change(y, 'S', S[i])
    # mprog.save_file(path, y)
else:
    mprog.save_file(path, y)