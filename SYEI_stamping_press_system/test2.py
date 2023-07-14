# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', par.A_15[i] / 2, 'XZ.PLANE', 1, 1)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', par.A[i] / 2, 'XZ.PLANE', 1, 1)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -par.Z[i], 'XY.PLANE', 1, 2)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - par.B_15[i] + par.F[i] / 2, 'YZ.PLANE', 1, 3)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME6.1', -80 - par.B[i] + par.F[i] / 2, 'YZ.PLANE', 1, 3)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -par.A_15[i] / 2, 'XZ.PLANE', 0, 4)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -par.A[i] / 2, 'XZ.PLANE', 0, 4)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -par.Z[i], 'XY.PLANE', 0, 5)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - par.B_15[i] + par.F[i] / 2, 'YZ.PLANE', 1, 6)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME7.1', -80 - par.B[i] + par.F[i] / 2, 'YZ.PLANE', 1, 6)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -par.A_15[i] / 2, 'XZ.PLANE', 0, 7)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -par.A[i] / 2, 'XZ.PLANE', 0, 7)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME5.1', -par.Z[i], 'XY.PLANE', 1, 8)
# if l == 0:
#     mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -par.B_15[i], 'YZ.PLANE', 1, 9)
# else:
#     mprog.add_offset_assembly('FRAME7.1', 'FRAME5.1', -par.B[i], 'YZ.PLANE', 1, 9)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', par.A_15[i] / 2, 'XZ.PLANE', 0, 10)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', par.A[i] / 2, 'XZ.PLANE', 0, 10)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME8.1', -par.Z[i], 'XY.PLANE', 0, 11)
# if l == 0:
#     mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -par.B_15[i], 'YZ.PLANE', 0, 12)
# else:
#     mprog.add_offset_assembly('FRAME6.1', 'FRAME8.1', -par.B[i], 'YZ.PLANE', 0, 12)
# # 左右側板
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', par.R_15[i] / 2 + 140, 'XZ.PLANE', 0, 13)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME2.1', par.R[i] / 2 + 140, 'XZ.PLANE', 0, 13)
# mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', (par.H[i] - 32 - 40) / 2 + 40, 'XY.PLANE', 0, 14)
# if l == 0:
#     mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', par.B_15[i] / 2, 'YZ.PLANE', 1, 15)
# else:
#     mprog.add_offset_assembly('FRAME8.1', 'FRAME2.1', par.B[i] / 2, 'YZ.PLANE', 1, 15)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -par.R_15[i] / 2 - 140, 'XZ.PLANE', 0, 16)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME1.1', -par.R[i] / 2 - 140, 'XZ.PLANE', 0, 16)
# mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -(par.H[i] - 40 - 32) / 2 - 40, 'XY.PLANE', 1, 17)
# if l == 0:
#     mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -par.B_15[i] / 2, 'YZ.PLANE', 0, 18)
# else:
#     mprog.add_offset_assembly('FRAME5.1', 'FRAME1.1', -par.B[i] / 2, 'YZ.PLANE', 0, 18)
# # 底部前、中板
# if l == 0:
#     mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', par.A_15[i] / 2, 'XZ.PLANE', 1, 19)
# else:
#     mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', par.A[i] / 2, 'XZ.PLANE', 1, 19)
# mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'XY.PLANE', 0, 20)
# mprog.add_offset_assembly('FRAME5.1', 'FRAME3.1', 0, 'YZ.PLANE', 1, 21)
# mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XZ.PLANE', 0, 22)
# mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', 0, 'XY.PLANE', 0, 23)
# if l == 0:
#     mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -par.FRAME2_lower_depth_15[i] + 5, 'YZ.PLANE', 0, 24)
# else:
#     mprog.add_offset_assembly('FRAME9.1', 'FRAME3.1', -par.FRAME2_lower_depth[i] + 5, 'YZ.PLANE', 0, 24)
# # 中間左右側板
# if l == 0:
#     mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', par.R_15[i] / 2, 'XZ.PLANE', 0, 25)
# else:
#     mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', par.R[i] / 2, 'XZ.PLANE', 0, 25)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME11.1', -par.T[i], 'XY.PLANE', 1, 26)
# if l == 0:
#     if i == 4:
#         mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', par.FRAME2_lower_depth_15[i] + 80, 'YZ.PLANE', 1, 27)
#     else:
#         mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', par.FRAME2_lower_depth_15[i], 'YZ.PLANE', 1, 27)
# else:
#     if i == 4:
#         mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', par.FRAME2_lower_depth[i] + 80, 'YZ.PLANE', 1, 27)
#     else:
#         mprog.add_offset_assembly('FRAME3.1', 'FRAME11.1', par.FRAME2_lower_depth[i], 'YZ.PLANE', 1, 27)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -par.R_15[i] / 2 - 90, 'XZ.PLANE', 0, 28)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -par.R[i] / 2 - 90, 'XZ.PLANE', 0, 28)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME10.1', -par.T[i] - alpha, 'XY.PLANE', 0, 29)
# if l == 0:
#     if i == 4:
#         mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', par.FRAME2_lower_depth_15[i] + 80, 'YZ.PLANE', 1, 30)
#     else:
#         mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', par.FRAME2_lower_depth_15[i], 'YZ.PLANE', 1, 30)
# else:
#     if i == 4:
#         mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', par.FRAME2_lower_depth[i] + 80, 'YZ.PLANE', 1, 30)
#     else:
#         mprog.add_offset_assembly('FRAME3.1', 'FRAME10.1', par.FRAME2_lower_depth[i], 'YZ.PLANE', 1, 30)
# # 底部後面ㄇ形角鐵
# mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XZ.PLANE', 1, 31)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', 0, 'XY.PLANE', 0, 32)
# if l == 0:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', par.B_15[i], 'YZ.PLANE', 0, 33)
# else:
#     mprog.add_offset_assembly('FRAME3.1', 'FRAME4.1', par.B[i], 'YZ.PLANE', 0, 33)
# # 平板底部零件
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -par.R_15[i] / 2, 'XZ.PLANE', 1, 34)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -par.R[i] / 2, 'XZ.PLANE', 1, 34)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', -par.T[i], 'XY.PLANE', 0, 35)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME12.1', 0, 'YZ.PLANE', 1, 36)
# # 平板底部零件
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', par.R_15[i] / 2, 'XZ.PLANE', 1, 37)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', par.R[i] / 2, 'XZ.PLANE', 1, 37)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', -par.T[i], 'XY.PLANE', 0, 38)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME13.1', 0, 'YZ.PLANE', 1, 39)
# # 前中上軸承板&角鐵
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME20.1', 0, 'XZ.PLANE', 1, 40)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', -par.H[i] + 40 - alpha - 8, 'XY.PLANE', 0, 41)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME20.1', 0, 'YZ.PLANE', 0, 42)
# mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', 0, 'XZ.PLANE', 0, 43)
# mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -9, 'XY.PLANE', 0, 44)
# mprog.add_offset_assembly('FRAME30.1', 'FRAME20.1', -550, 'YZ.PLANE', 0, 45)
# mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'XZ.PLANE', 1, 46)
# mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', -par.FRAME20_H[i] + alpha, 'XY.PLANE', 0, 47)
# mprog.add_offset_assembly('FRAME29.1', 'FRAME20.1', 0, 'YZ.PLANE', 1, 48)
# # 氣壓缸鎖固板左右
# if i == 4:
#     mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', -183, 'XZ.PLANE', 0, 49)
# else:
#     mprog.add_offset_assembly('FRAME2.1', 'FRAME21.1', -240, 'XZ.PLANE', 0, 49)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME21.1', -par.H[i] + 40 - alpha - 8, 'XY.PLANE', 1, 50)
# mprog.add_offset_assembly('FRAME20.1', 'FRAME21.1', 280, 'YZ.PLANE', 1, 51)
# if i == 4:
#     mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', 183, 'XZ.PLANE', 0, 52)
# else:
#     mprog.add_offset_assembly('FRAME1.1', 'FRAME22.1', 240, 'XZ.PLANE', 0, 52)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME22.1', -par.H[i] + 40 - alpha - 8, 'XY.PLANE', 0, 53)
# mprog.add_offset_assembly('FRAME20.1', 'FRAME22.1', 280, 'YZ.PLANE', 1, 54)
# # 鎖固平板六兄弟
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', 0, 'XZ.PLANE', 0, 55)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME14.1', -par.T[i], 'XY.PLANE', 0, 56)
# mprog.add_offset_assembly('FRAME3.1', 'FRAME14.1', -75, 'YZ.PLANE', 1, 57)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', par.R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 58)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', par.R[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 58)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -par.T[i], 'XY.PLANE', 0, 59)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME15.1', -par.F[i] / 2 + 80 + 37.5, 'YZ.PLANE', 1, 60)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', par.R_15[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 61)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', par.R[i] / 2 + 75 + 140, 'XZ.PLANE', 0, 61)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', -par.T[i], 'XY.PLANE', 0, 62)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME16.1', par.F[i] / 2 - 80 - 37.5, 'YZ.PLANE', 1, 63)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -par.R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 64)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -par.R[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 64)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', -par.T[i], 'XY.PLANE', 0, 65)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME18.1', par.F[i] / 2 - 80 - 37.5, 'YZ.PLANE', 0, 66)
# if l == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -par.R_15[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 67)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -par.R[i] / 2 - 75 - 140, 'XZ.PLANE', 1, 67)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -par.T[i], 'XY.PLANE', 0, 68)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME19.1', -par.F[i] / 2 + 80 + 37.5, 'YZ.PLANE', 0, 69)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', 0, 'XZ.PLANE', 1, 70)
# mprog.add_offset_assembly('BOLSTER1.1', 'FRAME17.1', -par.T[i], 'XY.PLANE', 0, 71)
# mprog.add_offset_assembly('FRAME9.1', 'FRAME17.1', 0, 'YZ.PLANE', 0, 72)
# # 左右側板前GIB
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME3.1', par.FRAME1_lower_high[i] + 40 - 34.5 + alpha, 'XY.PLANE', 0, 73)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME3.1', par.FRAME1_lower_high[i] + 40 + alpha, 'XY.PLANE', 0, 73)
# mprog.add_offset_assembly('GIB1.1', 'FRAME1.1', 72.5, 'XZ.PLANE', 1, 74)
# mprog.add_offset_assembly('GIB1.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0, 75)
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME3.1', par.FRAME1_lower_high[i] + 40 - 34.5 + alpha, 'XY.PLANE', 0, 76)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME3.1', par.FRAME1_lower_high[i] + 40 + alpha, 'XY.PLANE', 0, 76)
# mprog.add_offset_assembly('GIB2.1', 'FRAME2.1', -72.5, 'XZ.PLANE', 1, 77)
# mprog.add_offset_assembly('GIB2.1', 'FRAME5.1', 334.65, 'YZ.PLANE', 0, 78)
# # 左GIB後鎖固用方塊
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -par.pocket_1_upper_hole[i] + 80 + 3, 'XY.PLANE', 1, 79)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', -par.pocket_1_upper_hole[i] + 80 + 3, 'XY.PLANE', 1, 79)
# mprog.add_offset_assembly('FRAME2.1', 'FRAME23.1', -50 - 35, 'XZ.PLANE', 0, 80)
# mprog.add_offset_assembly('GIB2.1', 'FRAME23.1', 0, 'YZ.PLANE', 0, 81)
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', -34.5, 'XY.PLANE', 1, 82)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'XY.PLANE', 1, 82)
# mprog.add_offset_assembly('FRAME2.1', 'FRAME24.1', -50 - 35, 'XZ.PLANE', 1, 83)
# mprog.add_offset_assembly('GIB2.1', 'FRAME24.1', 0, 'YZ.PLANE', 1, 84)
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -par.pocket_1_upper_hole[i] + 340 + 80, 'XY.PLANE', 1, 85)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', -par.pocket_1_upper_hole[i] + 340 + 80, 'XY.PLANE', 1, 85)
# mprog.add_offset_assembly('FRAME2.1', 'FRAME27.1', -50, 'XZ.PLANE', 0, 86)
# mprog.add_offset_assembly('GIB2.1', 'FRAME27.1', 0, 'YZ.PLANE', 1, 87)
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -par.pocket_1_upper_hole[i], 'XY.PLANE', 1, 88)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME31.1', -par.pocket_1_upper_hole[i], 'XY.PLANE', 1, 88)
# mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -35, 'XZ.PLANE', 0, 89)
# mprog.add_offset_assembly('FRAME23.1', 'FRAME31.1', -130.35 - 22.5, 'YZ.PLANE', 0, 90)
# if i == 4:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150 - 34.5, 'XY.PLANE', 1, 91)
# else:
#     mprog.add_offset_assembly('GIB2.1', 'FRAME31.2', -150, 'XY.PLANE', 1, 91)
# mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -130.35 - 22.5, 'YZ.PLANE', 0, 92)
# mprog.add_offset_assembly('FRAME23.1', 'FRAME31.2', -35, 'XZ.PLANE', 0, 93)
# # 右GIB後鎖固用方塊
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -par.pocket_1_upper_hole[i] + 80 + 3, 'XY.PLANE', 1, 94)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', -par.pocket_1_upper_hole[i] + 80 + 3, 'XY.PLANE', 1, 94)
# mprog.add_offset_assembly('FRAME1.1', 'FRAME25.1', 50 + 35, 'XZ.PLANE', 1, 95)
# mprog.add_offset_assembly('GIB1.1', 'FRAME25.1', 0, 'YZ.PLANE', 1, 96)
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', -34.5 - 3, 'XY.PLANE', 1, 97)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 3, 'XY.PLANE', 1, 97)
# mprog.add_offset_assembly('FRAME1.1', 'FRAME26.1', 50 + 35, 'XZ.PLANE', 1, 98)
# mprog.add_offset_assembly('GIB1.1', 'FRAME26.1', 0, 'YZ.PLANE', 0, 99)
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -par.pocket_1_upper_hole[i] + 340 + 80, 'XY.PLANE', 1, 100)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', -par.pocket_1_upper_hole[i] + 340 + 80, 'XY.PLANE', 1, 100)
# mprog.add_offset_assembly('FRAME1.1', 'FRAME28.1', 50, 'XZ.PLANE', 1, 101)
# mprog.add_offset_assembly('GIB1.1', 'FRAME28.1', 0, 'YZ.PLANE', 1, 102)
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -par.pocket_1_upper_hole[i], 'XY.PLANE', 1, 103)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME31.3', -par.pocket_1_upper_hole[i], 'XY.PLANE', 1, 103)
# mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', -35, 'XZ.PLANE', 1, 104)
# mprog.add_offset_assembly('FRAME25.1', 'FRAME31.3', 130.35 + 22.5, 'YZ.PLANE', 1, 105)
# if i == 4:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME31.4', -150 - 34.5 - 3, 'XY.PLANE', 1, 106)
# else:
#     mprog.add_offset_assembly('GIB1.1', 'FRAME31.4', -150, 'XY.PLANE', 1, 106)
# mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'YZ.PLANE', 0, 107)
# mprog.add_offset_assembly('FRAME31.3', 'FRAME31.4', 0, 'XZ.PLANE', 0, 108)
# # 中間與平板鎖固方塊下板子
# mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', -48, 'XY.PLANE', 1, 109)
# mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'XZ.PLANE', 0, 110)
# mprog.add_offset_assembly('FRAME43.1', 'FRAME17.1', 0, 'YZ.PLANE', 1, 111)
# # 後面安裝馬達下板子
# mprog.add_offset_assembly('FRAME41.1', 'FRAME20.1', -1340, 'XY.PLANE', 0, 112)
# mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'XZ.PLANE', 0, 113)
# mprog.add_offset_assembly('FRAME41.1', 'FRAME4.1', 0, 'YZ.PLANE', 0, 114)
# # FRAME2中間下板子
# mprog.add_offset_assembly('FRAME32.1', 'BOLSTER1.1', 0, 'XZ.PLANE', 0, 115)
# mprog.add_offset_assembly('FRAME32.1', 'FRAME3.1', par.FRAME_32_XY[i] + 20 + alpha, 'XY.PLANE', 0, 116)
# mprog.add_offset_assembly('FRAME32.1', 'FRAME30.1', 0, 'YZ.PLANE', 0, 117)
# # 後方大軸承支架
# mprog.add_offset_assembly('FRAME10.1', 'FRAME34.2', -485.543, 'YZ.PLANE', 1, 118)
# mprog.add_offset_assembly('FRAME34.2', 'FRAME10.1', -(par.FRAME_10_H[i] + 40 + alpha), 'XY.PLANE', 0, 119)
# if l == 0:
#     mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -par.R_15[i] / 2 - 90, 'XZ.PLANE', 0, 120)
# else:
#     mprog.add_offset_assembly('FRAME34.2', 'FRAME3.1', -par.R[i] / 2 - 90, 'XZ.PLANE', 0, 120)
# if l == 0:
#     mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -par.R_15[i] - 180, 'XZ.PLANE', 1, 121)
# else:
#     mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', -par.R[i] - 180, 'XZ.PLANE', 1, 121)
# mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'XY.PLANE', 1, 122)
# mprog.add_offset_assembly('FRAME34.1', 'FRAME34.2', 0, 'YZ.PLANE', 0, 123)
# # 後方大軸承
# mprog.add_offset_assembly('FRAME35.1', 'FRAME3.1', 0, 'XZ.PLANE', 1, 124)
# if i == 4:
#     mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -(par.H[i] - par.Z[i] - par.S[i] - 34 + 448 - 80),
#                               'XY.PLANE', 0, 125)
# else:
#     mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -(par.H[i] - par.Z[i] - par.S[i] - 34 + 448),
#                               'XY.PLANE', 0, 125)
# mprog.add_offset_assembly('FRAME35.1', 'FRAME20.1', -par.FRAME20_FRAME2_YZ[i] - 36, 'YZ.PLANE', 0, 126)
# # 後方馬達下支撐板
# mprog.add_offset_assembly('FRAME39.1', 'FRAME1.1', 50, 'XZ.PLANE', 1, 127)
# mprog.add_offset_assembly('FRAME39.1', 'FRAME20.1', -360, 'XY.PLANE', 0, 128)
# mprog.add_offset_assembly('FRAME39.1', 'FRAME7.1', 320, 'YZ.PLANE', 0, 129)
# mprog.add_offset_assembly('FRAME38.1', 'FRAME2.1', -50, 'XZ.PLANE', 1, 130)
# mprog.add_offset_assembly('FRAME38.1', 'FRAME20.1', 360 + 19, 'XY.PLANE', 1, 131)
# mprog.add_offset_assembly('FRAME38.1', 'FRAME6.1', -320, 'YZ.PLANE', 1, 132)
# mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 0, 'XZ.PLANE', 0, 133)
# mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 19, 'XY.PLANE', 1, 134)
# mprog.add_offset_assembly('FRAME37.1', 'FRAME39.1', 125, 'YZ.PLANE', 1, 135)
# mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', 0, 'XZ.PLANE', 0, 136)
# mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', 19, 'XY.PLANE', 1, 137)
# mprog.add_offset_assembly('FRAME37.2', 'FRAME39.1', -125, 'YZ.PLANE', 1, 138)
# mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', 0, 'XZ.PLANE', 1, 139)
# mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', 0, 'XY.PLANE', 0, 140)
# mprog.add_offset_assembly('FRAME37.3', 'FRAME38.1', -125, 'YZ.PLANE', 1, 141)
# mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 0, 'XZ.PLANE', 1, 142)
# mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 0, 'XY.PLANE', 0, 143)
# mprog.add_offset_assembly('FRAME37.4', 'FRAME38.1', 125, 'YZ.PLANE', 1, 144)
# # FRAME2右邊角鐵
# mprog.add_offset_assembly('FRAME33.1', 'FRAME2.1', 140, 'XZ.PLANE', 0, 145)
# mprog.add_offset_assembly('FRAME33.1', 'FRAME15.1', -708, 'XY.PLANE', 0, 146)
# mprog.add_offset_assembly('FRAME33.1', 'FRAME11.1', 313.984, 'YZ.PLANE', 0, 147)
# # FRAME2中間半圓形零件
# mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', 130.5, 'XZ.PLANE', 0, 148)
# mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', -320, 'XY.PLANE', 0, 149)
# mprog.add_offset_assembly('FRAME40.1', 'FRAME20.1', par.FRAME20_FRAME2_YZ[i], 'YZ.PLANE', 1, 150)
# # FRAME35上兩圓管
# mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', -272.236, 'XZ.PLANE', 1, 151)
# mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', 272.236, 'XY.PLANE', 0, 152)
# mprog.add_offset_assembly('FRAME36.1', 'FRAME35.1', 41, 'YZ.PLANE', 0, 153)
# mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 272.236, 'XZ.PLANE', 1, 154)
# mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 272.236, 'XY.PLANE', 0, 155)
# mprog.add_offset_assembly('FRAME36.2', 'FRAME35.1', 41, 'YZ.PLANE', 0, 156)
# # 後方馬達下支撐板上治具
# mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 95, 'XZ.PLANE', 0, 157)
# mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 22 + 18, 'XY.PLANE', 0, 158)
# if i == 4 and l == 1:
#     mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 312.5 - 185, 'YZ.PLANE', 0, 159)
# else:
#     mprog.add_offset_assembly('Fixture.1', 'FRAME41.1', 312.5, 'YZ.PLANE', 0, 159)
# # 馬達下方板與後方橫板
# mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', 0, 'XZ.PLANE', 1, 160)
# mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', 71.5, 'XY.PLANE', 1, 161)
# mprog.add_offset_assembly('FRAME42.1', 'FRAME41.1', -19, 'YZ.PLANE', 1, 162)
# # 平板2, 3組立
# mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0, 163)
# mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'XY.PLANE', 0, 164)
# mprog.add_offset_assembly('BOLSTER2.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0, 165)
# # 平板1, 3組立
# mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 0, 166)
# mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0, 167)
# # if k == 0:
# if h == 0:
#     mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', par.DH_S[i] + alpha, 'XY.PLANE', 0, 168)
# elif h == 1:
#     mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', par.DH_H[i] + alpha, 'XY.PLANE', 0, 168)
# else:
#     mprog.add_offset_assembly('BOLSTER1.1', 'BOLSTER3.1', par.DH_P[i] + alpha, 'XY.PLANE', 0, 168)
# # SLIDE跟平板3組立
# if i == 4:
#     mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 267.38, 'XY.PLANE',
#                                       1, 169)
# else:
#     mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 85, 'XY.PLANE', 1,
#                                       170)
# mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'XZ.PLANE', 1, 171)
# if i == 4:
#     mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 138.171,
#                                       'YZ.PLANE', 0, 172)
# else:
#     mprog.add_offset_product_assembly('SLIDE_UNIT_All.1', 'Geometrical Set.1', 'BOLSTER3.1', 0, 'YZ.PLANE', 0,
#                                       172)
# # 時鐘跟FRAME20組立
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', 0, 'XZ.PLANE', 0,
#                                   173)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1', 33, 'YZ.PLANE', 0,
#                                   174)
# if i == 4:
#     mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1',
#                                       -(par.H[i] - 34 - par.S[i] - par.Z[i] - 80), 'XY.PLANE',
#                                       0, 175)
# else:
#     mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'FRAME20.1',
#                                       -(par.H[i] - 34 - par.S[i] - par.Z[i]), 'XY.PLANE',
#                                       0, 175)
# # 氣壓缸跟FRAME20組立
# if i == 4:
#     mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -87.5 - 8,
#                                       'XY.PLANE', 1, 176)
# else:
#     mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE',
#                                       1, 176)
# if i == 4:
#     if l == 0:
#         mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -500.5,
#                                           'XZ.PLANE', 0, 177)
#     else:
#         mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -347,
#                                           'XZ.PLANE', 0, 177)
# else:
#     if l == 0:
#         mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                           -par.R_15[i] / 2 - 13, 'XZ.PLANE', 0, 178)
#     else:
#         mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                           -par.R[i] / 2 - 13, 'XZ.PLANE', 0, 178)
# if i == 4:
#     mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -261.248,
#                                       'YZ.PLANE', 0, 179)
# else:
#     mprog.add_offset_product_assembly('BALANCER_RIGHT_All.1', 'Geometrical Set.1', 'FRAME20.1', -260,
#                                       'YZ.PLANE', 0, 179)
# if i == 4:
#     mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -21.8 - 8,
#                                       'XY.PLANE', 1, 180)
# else:
#     mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -32, 'XY.PLANE',
#                                       1, 180)
# if i == 4:
#     if l == 0:
#         mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -500.5,
#                                           'XZ.PLANE', 1, 181)
#     else:
#         mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', -347,
#                                           'XZ.PLANE', 1, 181)
# else:
#     if l == 0:
#         mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                           -par.R_15[i] / 2 - 13, 'XZ.PLANE', 1, 182)
#     else:
#         mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                           -par.R[i] / 2 - 13, 'XZ.PLANE', 1, 182)
# if i == 4:
#     mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 261.248,
#                                       'YZ.PLANE', 1, 183)
# else:
#     mprog.add_offset_product_assembly('BALANCER_LEFT_All.1', 'Geometrical Set.1', 'FRAME20.1', 260, 'YZ.PLANE',
#                                       1, 183)
# # 離合器與FRAME20結合
# if i == 4:
#     mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                       -(par.H[i] - par.Z[i] - par.S[i] - 34 + 448 - 80), 'XY.PLANE', 0, 184)
# else:
#     mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                       -(par.H[i] - par.Z[i] - par.S[i] - 34 + 448), 'XY.PLANE', 0, 184)
#
# mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE', 1,
#                                   185)
# mprog.add_offset_product_assembly('CLUCTH_ASSEMBLY_All.1', 'Geometrical Set.1', 'FRAME20.1', 512, 'YZ.PLANE', 1,
#                                   186)
# # 曲軸與時鐘結合
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 0, 'XZ.PLANE',
#                                   0, 187)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 3, 'YZ.PLANE',
#                                   1, 188)
# mprog.add_offset_product_assembly('CRANK_SHAFT_CLOCK.1', 'Geometrical Set.2', 'CRANK_SHAFT.1.1', 0, 'XY.PLANE',
#                                   1, 189)
# # 大齒輪結合
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 0, 'XZ.PLANE', 0, 190)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 77.172, 'XY.PLANE', 0, 191)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'MAIN_GEAR4.1', 0, 'YZ.PLANE', 1, 192)
# # 大齒輪結合曲軸
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'CRANK_SHAFT.1.1', 0, 'XZ.PLANE', 1, 193)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'CRANK_SHAFT.1.1', 0, 'XY.PLANE', 0, 194)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'FRAME20.1', 592.5 + 150, 'YZ.PLANE', 1, 195)
# # 機架20結合曲軸旁圓管
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', 0, 'XZ.PLANE', 0, 196)
# if i == 4:
#     mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', par.H[i] - par.Z[i] - par.S[i] - 34 - 80, 'XY.PLANE', 1,
#                               197)
# else:
#     mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', par.H[i] - par.Z[i] - par.S[i] - 34, 'XY.PLANE', 1, 197)
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'FRAME20.1', 520.5, 'YZ.PLANE', 1, 198)
# # 大齒輪結合長棒
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', 0, 'XZ.PLANE', 0, 199)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', 0, 'XY.PLANE', 1, 200)
# mprog.add_offset_assembly('MAIN_GEAR1.1', 'JOINT1.1', -84, 'YZ.PLANE', 1, 201)
# #
# mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', 0, 'XZ.PLANE',
#                                   0, 202)
# mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1', 1010, 'YZ.PLANE',
#                                   1, 203)
# if i == 4:
#     mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                       -(par.H[i] - par.Z[i] - par.S[i] - 34 - 80), 'XY.PLANE',
#                                       0, 204)
# else:
#     mprog.add_offset_product_assembly('JOINT_All.1', 'Geometrical Set.1', 'FRAME20.1',
#                                       -(par.H[i] - par.Z[i] - par.S[i] - 34), 'XY.PLANE',
#                                       0, 204)
# # 大齒輪內套環
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 0, 'XZ.PLANE', 1, 205)
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 0, 'XY.PLANE', 0, 206)
# mprog.add_offset_assembly('MAIN_GEAR2.1', 'MAIN_GEAR3.1', 11, 'YZ.PLANE', 1, 207)
# # BOLSTER後板
# mprog.add_offset_assembly('FRAME45.1', 'FRAME17.1', 0, 'XZ.PLANE', 0, 208)
# mprog.add_offset_assembly('FRAME45.1', 'FRAME17.1', -48, 'XY.PLANE', 1, 209)
# mprog.add_offset_assembly('FRAME45.1', 'FRAME17.1', 0, 'YZ.PLANE', 1, 210)
# #
# mprog.add_offset_assembly('FRAME46.1', 'FRAME17.1', -48, 'XY.PLANE', 1, 211)
# mprog.add_offset_assembly('FRAME46.1', 'FRAME17.1', 0, 'XZ.PLANE', 0, 212)
# mprog.add_offset_assembly('FRAME46.1', 'FRAME17.1', 0, 'YZ.PLANE', 1, 213)
# #
# mprog.add_offset_assembly('FRAME20.1', 'FRAME44.1', 9, 'XY.PLANE', 0, 214)
# mprog.add_offset_assembly('FRAME20.1', 'FRAME44.1', 0, 'XZ.PLANE', 0, 215)
# mprog.add_offset_assembly('FRAME20.1', 'FRAME44.1', 979, 'YZ.PLANE', 1, 216)
# #
# mprog.add_offset_assembly('FRAME26.1', 'FRAME47.1', 29 + 3, 'XY.PLANE', 1, 217)
# mprog.add_offset_assembly('FRAME26.1', 'FRAME47.1', 35, 'XZ.PLANE', 0, 218)
# mprog.add_offset_assembly('FRAME26.1', 'FRAME47.1', -99.35, 'YZ.PLANE', 1, 219)
# mprog.add_offset_assembly('FRAME24.1', 'FRAME47.2', 10, 'XY.PLANE', 0, 220)
# mprog.add_offset_assembly('FRAME24.1', 'FRAME47.2', -35, 'XZ.PLANE', 1, 221)
# mprog.add_offset_assembly('FRAME24.1', 'FRAME47.2', 099.35, 'YZ.PLANE', 0, 222)
# # 方管
# mprog.add_offset_assembly('FRAME44.1', 'FRAME48.1', 124, 'XY.PLANE', 1, 223)
# mprog.add_offset_assembly('FRAME44.1', 'FRAME48.1', 159 + 16, 'XZ.PLANE', 0, 224)
# mprog.add_offset_assembly('FRAME44.1', 'FRAME48.1', -32, 'YZ.PLANE', 0, 225)
# # 圓柱
# mprog.add_offset_assembly('FRAME44.1', 'FRAME49.1', 220, 'XY.PLANE', 0, 226)
# mprog.add_offset_assembly('FRAME44.1', 'FRAME49.1', 145, 'XZ.PLANE', 0, 227)
# mprog.add_offset_assembly('FRAME44.1', 'FRAME49.1', -32, 'YZ.PLANE', 0, 228)
# # 2023/02/23更新 FRAME1左邊角鐵
# mprog.add_offset_assembly('FRAME1.1', 'FRAME33.2', 140, 'XZ.PLANE', 1, 229)
# mprog.add_offset_assembly('FRAME19.1', 'FRAME33.2', 708, 'XY.PLANE', 0, 230)
# mprog.add_offset_assembly('FRAME10.1', 'FRAME33.2', -313.984, 'YZ.PLANE', 1, 231)
#
# # 更新
# mprog.update()
# mprog.Close_All()
# # 儲存Product檔

import win32com.client as win32
import time
import parameter as par
import file_path as fp
import main_program as mprog


# # 你確定要挑戰玄學?
# def StartCommand(name):  # 孤立元素
#     catapp = win32.Dispatch("CATIA.Application")
#     selection = catapp.ActiveDocument.Selection
#     # selection.Clear()
#     # selection.Search("Name='" + str(name) + "', all")
#     catapp.StartCommand("isometric View")
#
def create():
    partDocument1 = CATIA.ActiveDocument
    part1 = partDocument1.Part
    shapeFactory1 = part1.ShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    sketches1 = body1.Sketches
    sketch1 = sketches1.Item("Sketch.1")
    pad1 = shapeFactory1.AddNewPad(sketch1, 20)
    part1.Update()






klllllllllllllllllllllllllllllllllll

create()
def change_parameter(name, i):
    if name == 'FRAME1':
        excel = epc.ExcelOp('FRAME1')
        try:
            FRAME1_parameter_name, FRAME1_parameter_value = excel.part_parameter('FRAME1', i)
            print('FRAME1 Parameter change success')
        except:
            print('FRAME1 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角填料_1')
                mprog.partbodyfeatureactivate('R倒角填料_2')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('小噸數E倒角_B')
                mprog.partbodyfeatureactivate('小噸數Y倒角_B1')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_B')
                mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
                mprog.activatefeature('110通孔', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('6-M5X16L(配管)', 5)
                mprog.activatefeature('5-M5X10L(配管用)', 2)
                mprog.activatefeature('兩點組合', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角填料_1')
                mprog.partbodyfeatureactivate('R倒角填料_2')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('小噸數E倒角_B')
                mprog.partbodyfeatureactivate('小噸數Y倒角_B1')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_B')
                mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
                mprog.activatefeature('110通孔', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('6-M5X16L(配管)', 5)
                mprog.activatefeature('5-M5X10L(配管用)', 3)
                mprog.activatefeature('兩點組合', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角填料_1')
                mprog.partbodyfeatureactivate('R倒角填料_2')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('小噸數E倒角_B')
                mprog.partbodyfeatureactivate('小噸數Y倒角_B1')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_B')
                mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
                mprog.activatefeature('110通孔', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('6-M5X16L(配管)', 5)
                mprog.activatefeature('5-M5X10L(配管用)', 2)
                mprog.activatefeature('兩點組合', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角填料_1')
                mprog.partbodyfeatureactivate('R倒角填料_2')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_60')
                mprog.partbodyfeatureactivate('E_倒角')
                mprog.activatefeature('110通孔', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('6-M5X16L(配管)', 5)
                mprog.activatefeature('5-M5X10L(配管用)', 2)
                mprog.activatefeature('兩點組合', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
                mprog.partbodyfeatureactivate('B倒角_E段')
                mprog.partbodyfeatureactivate('喉口倒角補料_80')
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('6-M5X16L(配管)', 5)
                mprog.activatefeature('5-M5X10L(配管用)', 3)
                mprog.activatefeature('兩點組合', 0)
                mprog.activatefeature('3-M5X10L(配管用)(r1)', 1)

            elif i == 5:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_110')
                mprog.partbodyfeatureactivate('E_倒角')
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('6-M5X16L(配管)', 0)
                mprog.activatefeature('5-M5X10L(配管用)', 3)
                mprog.activatefeature('兩點組合', 0)
                mprog.activatefeature('3-M5X10L(配管用)(r1)', 2)

            elif i == 6:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_160')
                mprog.partbodyfeatureactivate('E_倒角')
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('3-M5X10L(配管用)(r1)', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('6-M5X16L(配管)', 0)
                mprog.activatefeature('5-M5X10L(配管用)', 5)
                mprog.activatefeature('兩點組合', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_200')
                mprog.partbodyfeatureactivate('E_倒角')
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('3-M5X10L(配管用)(r1)', 2)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('6-M5X16L(配管)', 0)
                mprog.activatefeature('5-M5X10L(配管用)', 5)
                mprog.activatefeature('兩點組合', 0)
            elif i == 8:
                mprog.partbodyfeatureactivate('T250')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1_250')
                mprog.partbodyfeatureactivate('V3_250')
                mprog.partbodyfeatureactivate('V5_250')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('喉口倒角_250')
                mprog.partbodyfeatureactivate('喉口倒圓角_250')
                mprog.partbodyfeatureactivate('喉口倒角補料_250')
                mprog.partbodyfeatureactivate('E_倒角')
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('3-M5X10L(配管用)(r1)', 0)
                mprog.activatefeature('油面劑用', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('電動黃油泵用', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('6-M5X16L(配管)', 0)
                mprog.activatefeature('5-M5X10L(配管用)', 0)
                mprog.activatefeature('兩點組合', 0)

        except:
            print('FRAME1 Part activate error')
        finally:
            mprog.update()
            print('FRAME1 update success')
    elif name == 'FRAME2':
        start = time.time()
        excel = epc.ExcelOp('FRAME2')
        try:
            FRAME2_parameter_name, FRAME2_parameter_value = excel.part_parameter('FRAME2', i)

            print('FRAME2 Parameter change success')
        except:
            print('FRAME2 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒圓角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_45')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角_qqqq')
                mprog.partbodyfeatureactivate('R倒角_G1')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段_B1倒角')
                mprog.partbodyfeatureactivate('E段_B倒角_Y')
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 4)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 1)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('解角器裝置線用', 0)
                mprog.activatefeature('解角器用', 0)
                mprog.activatefeature('變頻器用', 0)
                mprog.activatefeature('解角器護蓋用', 0)
                mprog.activatefeature('M5X10L(配管用)', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒圓角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_60')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角_qqqq')
                mprog.partbodyfeatureactivate('R倒角_G1')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段_B1倒角')
                mprog.partbodyfeatureactivate('E段_B倒角_Y')
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 4)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 1)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('解角器裝置線用', 0)
                mprog.activatefeature('解角器用', 0)
                mprog.activatefeature('變頻器用', 0)
                mprog.activatefeature('解角器護蓋用', 0)
                mprog.activatefeature('M5X10L(配管用)', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('1-M5X10L底孔20L配管用', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒圓角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_45')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角_qqqq')
                mprog.partbodyfeatureactivate('R倒角_G1')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段_B1倒角')
                mprog.partbodyfeatureactivate('E段_B倒角_Y')
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 4)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('解角器裝置線用', 0)
                mprog.activatefeature('解角器用', 0)
                mprog.activatefeature('變頻器用', 0)
                mprog.activatefeature('解角器護蓋用', 0)
                mprog.activatefeature('M5X10L(配管用)', 0)
                mprog.activatefeature('1-M5X10L底孔20L配管用', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_60')
                mprog.partbodyfeatureactivate('R倒角_無qqqq')
                mprog.partbodyfeatureactivate('R倒角_G1')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_非B')
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 4)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('解角器裝置線用', 0)
                mprog.activatefeature('解角器用', 0)
                mprog.activatefeature('變頻器用', 0)
                mprog.activatefeature('解角器護蓋用', 0)
                mprog.activatefeature('M5X10L(配管用)', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒圓角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_80')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_B')
                mprog.partbodyfeatureactivate('jjj4')
                mprog.partbodyfeatureactivate('mmm4')
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 4)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('解角器裝置線用', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('4-M8X16L', 0)
                mprog.activatefeature('2-M5(配管用)', 0)
                mprog.activatefeature('台達變頻器用10HP', 0)
                mprog.activatefeature('東元變頻器用10HP', 0)
                mprog.activatefeature('電氣箱連桿', 2)
            elif i == 5:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_110')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_非B')
                mprog.partbodyfeatureactivate('dd4')
                mprog.partbodyfeatureactivate('x4')
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 5)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('4-M8X16L', 0)
                mprog.activatefeature('2-M5(配管用)', 0)
                mprog.activatefeature('電氣箱連桿', 0)
                mprog.activatefeature('東元變頻器10,15HP440V用', 0)
                mprog.activatefeature('台達變頻器10,15HP', 0)
                mprog.activatefeature('東元變頻器,15HP220V', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_110')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_非B')
                mprog.partbodyfeatureactivate('ggg4')
                mprog.partbodyfeatureactivate('xx3')
                mprog.partbodyfeatureactivate('iii4')
                mprog.partbodyfeatureactivate('fff4')
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 5)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('4-M8X16L', 0)
                mprog.activatefeature('2-M5(配管用)', 0)
                mprog.activatefeature('電氣箱連桿', 0)
                mprog.activatefeature('東元變頻器440V,15,20HP', 0)
                mprog.activatefeature('台達變頻器440V,15,20HP；220V15HP', 0)
                mprog.activatefeature('台達變頻器220V,20HP', 0)
                mprog.activatefeature('東元變頻器220V,15,20HP', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_200')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_非B')
                mprog.partbodyfeatureactivate('iii4')
                mprog.partbodyfeatureactivate('ssss4')
                mprog.partbodyfeatureactivate('sss4')
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 5)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('4-M8X16L', 0)
                mprog.activatefeature('2-M5(配管用)', 0)
                mprog.activatefeature('電氣箱連桿', 0)
                mprog.activatefeature('台達變頻器220V,20HP', 0)
                mprog.activatefeature('東元變頻器220V/440V,20HP', 0)
                mprog.activatefeature('台達變頻器440V,20HP', 0)
            elif i == 8:
                mprog.partbodyfeatureactivate('T250')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1_250')
                mprog.partbodyfeatureactivate('V3_250')
                mprog.partbodyfeatureactivate('V5_250')
                mprog.partbodyfeatureactivate('喉口倒角_250')
                mprog.partbodyfeatureactivate('喉口倒圓角_250')
                mprog.partbodyfeatureactivate('喉口倒角補料_250')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_非B')
                mprog.activatefeature('90通孔', 0)
                mprog.activatefeature('WIRE_CASING', 5)
                mprog.activatefeature('2-M5(側邊配管用)', 0)
                mprog.activatefeature('吊孔', 0)
                mprog.activatefeature('3-M5X10L(配管用)', 0)
                mprog.activatefeature('60通孔', 0)
                mprog.activatefeature('5-M5X15L(配管用)', 0)
                mprog.activatefeature('氣壓錶箱用', 0)
                mprog.activatefeature('電氣箱用', 0)
                mprog.activatefeature('2-M4X10L', 0)
                mprog.activatefeature('2-M10X20L', 0)
                mprog.activatefeature('1-M5X10L(配管用)', 0)
                mprog.activatefeature('4-M8X16L', 0)
                mprog.activatefeature('2-M5(配管用)', 0)
                mprog.activatefeature('電氣箱連桿', 0)
                mprog.activatefeature('東元變頻器25HP', 0)
                mprog.activatefeature('台達變頻器25HP', 0)
        except:
            print('FRAME2 Part activate error')
        finally:
            mprog.update()
            print('FRAME2 update success')
    elif name == 'FRAME5':
        excel = epc.ExcelOp('FRAME5')
        try:
            FRAME5_parameter_name, FRAME5_parameter_value = excel.part_parameter('FRAME5', i)
            print('FRAME5 Parameter change success')
        except:
            print('FRAME5 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('SN1-25_X')
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('Y')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
                mprog.activatefeature('PT1/2', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('Y')
                mprog.partbodyfeatureactivate('AA')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
                mprog.activatefeature('PT1/2', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('Y')
                mprog.partbodyfeatureactivate('AA')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
                mprog.activatefeature('PT1/2', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('S')
                mprog.partbodyfeatureactivate('Rx2')
                mprog.partbodyfeatureactivate('SN1-60_R')
                mprog.partbodyfeatureactivate('Chamfer.2')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 5:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('6-M6通C5')
                mprog.activatefeature('6-M6通', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('3-M8通_C5')
                mprog.activatefeature('3-M8通', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('5-M8通_C5')
                mprog.activatefeature('5-M8通', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 8:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('5-M8通_C5')
                mprog.activatefeature('5-M8通', 0)
                mprog.activatefeature('M6通孔LED', 0)
        except:
            print('FRAME5 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME5 Update success')
            except:
                print('FRAME5 Update error')
    elif name == 'FRAME8':
        excel = epc.ExcelOp('FRAME8')
        # try:
        FRAME8_parameter_name, FRAME8_parameter_value = excel.part_parameter('FRAME8', i)
        #     print('FRAME8 Parameter change success')
        # except:
        #     print('FRAME8 Parameter change error')
        if i == 0:
            mprog.partbodyfeatureactivate('chamfer.1')
            mprog.partbodyfeatureactivate('chamfer.2')
            mprog.activatefeature('')
    elif name == 'FRAME11':
        start = time.time()
        excel = epc.ExcelOp('FRAME11')
        try:
            FRAME11_parameter_name, FRAME11_parameter_value = excel.part_parameter('FRAME11', i)
            print('FRAME11 Parameter change success')
        except:
            print('FRAME11 Parameter change error')
        if i == 0:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 1:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 2:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 3:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 4:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 5:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 6:
            mprog.activatefeature('挖孔', 0)
        elif i == 7:
            pass
        elif i == 8:
            pass
        try:
            mprog.Update()
            print('FRAME11 Update success')
        except:
            print('FRAME11 Update error')
change_parameter('FRAME11', 0)