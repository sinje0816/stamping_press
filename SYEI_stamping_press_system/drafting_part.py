from PyQt5 import QtCore, QtGui, QtWidgets
import sys, datetime, os
from GUI import Ui_Dialog
import main_program as mprog
import win32com.client as win32
import drafting_part_calculate as DP
import drafting as draft
import parameter as par

# 圖框範圍
circle_Xgap = [150, 150, 150, 150, 200, 250, 250, 300, 400, 400]
circle_gap = [100, 100, 100, 100, 300, 300, 350, 400, 370, 430]
circle_15_gap = [200, 200, 200, 200, 300, 300, 350, 400, 370, 430]
ALL_range = []
gap = 5
drafting_min_Y = 44
drafting_max_Y = 810
drafting_min_X = 20
drafting_max_X = 1169
# 第一虛擬方框位置
box_1_Xmax = 372
box_1_Ymax = 810
box_1_Xmin = 25
box_1_Ymin = 25
box_width_gap = 80 + 2 * gap  # 虛擬方框一的寬度間隙
box_heigth_gap = 150 + 2 * gap  # 虛擬方框一的高度間隙


def Parts_drawing_generation(i, l, path, alpha):
    scale, box_1_center, box_1_range = DP.scale_Adjustment(i, l)
    ALL_range, scale = DP.drafting_parameter_calculation(i, l, scale,
                                                         box_1_range)  # ALL_range = [方寬編號[Xmax , Xmin , Ymax , Ymin]]
    projection_file_name_list = ['FRAME1', 'FRAME2', 'FRAME30', 'FRAME44', 'FRAME41', 'FRAME34', 'FRAME9', 'FRAME45',
                                 'FRAME43', 'FRAME7',
                                 'MAIN_GEAR2', 'FRAME32', 'FRAME47', 'FRAME23', 'FRAME31', 'FRAME24', 'FRAME38',
                                 'FRAME11', 'FRAME39', 'FRAME17',
                                 'FRAME3', 'FRAME13', 'FRAME14', 'FRAME22',
                                 'FRAME37', 'FRAME8', 'FRAME29', 'FRAME20', 'FRAME33', 'FRAME48', 'FRAME49']
    part_circle_position = {
        '1': ['1', -par.B[i] / 2 - circle_Xgap[i], par.H[i] / 2 + circle_gap[i]],
        '2': ['2', -par.B[i] / 2 - circle_Xgap[i], par.H[i] / 2 + circle_gap[i]],
        '3': ['3', -(par.R[i] + 180) / 2 - circle_Xgap[i], par.FRAME44_height[i] / 2 + circle_gap[i]],
        '4': ['4', -(par.R[i] + 180) / 2 - circle_Xgap[i], par.FRAME44_height[i] / 2 + circle_gap[i]],
        '5': ['5', -(par.R[i] + 180) / 2 - circle_Xgap[i], par.FRAME_41_depth[i] / 2 + circle_gap[i]],
        '6': ['6', -182 / 2 - circle_Xgap[i], 80 / 2 + circle_gap[i]],
        '7': ['7', -(par.R[i] + 180) / 2 - circle_Xgap[i], 50 / 2 + circle_gap[i]],
        '8': ['8', -par.R[i] / 2 - circle_Xgap[i], 476 / 2 + circle_gap[i]],
        '9': ['8-1', -par.R[i] / 2 - circle_Xgap[i], 150 / 2 + circle_gap[i] + 50],
        '10': ['10', -par.FRAME_7_width[i] / 2 - circle_Xgap[i], 300 / 2 + circle_gap[i]],
        '11': ['11', -125 / 2 - circle_Xgap[i], 270 / 2 + circle_gap[i]],
        '12': ['12', -par.R[i] / 2 - circle_Xgap[i], 429 / 2 + circle_gap[i]],
        '13': ['13', -50 / 2 - circle_Xgap[i], 74 / 2 + circle_gap[i]],
        '14': ['14', -99.35 / 2 - circle_Xgap[i], 35 / 2 + circle_gap[i]],
        '15': ['15A', -40 / 2 - circle_Xgap[i], 150 / 2 + circle_gap[i] + 50],
        '16': ['15', -40 / 2 - circle_Xgap[i], 150 / 2 + circle_gap[i]],
        '17': ['16', -99.35 / 2 - circle_Xgap[i], 35 / 2 + circle_gap[i]],
        '18': ['18', -290 / 2 - circle_Xgap[i], 145 / 2 + circle_gap[i]],
        '19': ['19', -par.FRAME_11_width[i] / 2 - circle_Xgap[i], 90 / 2 + circle_gap[i]],
        '20': ['20', -145 / 2 - circle_Xgap[i], 300 / 2 + circle_gap[i]],
        '21': ['21', -75 / 2 - circle_Xgap[i], 75 / 2 + circle_gap[i]],
        '22': ['22', -(par.R[i] + 180) / 2 - circle_Xgap[i], 50 / 2 + circle_gap[i]],
        '23': ['24', -85 / 2 - circle_Xgap[i], par.FRAME_13_depth[i] / 2 + circle_gap[i]],
        '24': ['25', -75 / 2 - circle_Xgap[i], 75 / 2 + circle_gap[i]],
        '25': ['27', -460 / 2 - circle_Xgap[i], 280 / 2 + circle_gap[i] - 150],
        '26': ['29', -140 / 2 - circle_Xgap[i], 80 / 2 + circle_gap[i]],
        '27': ['30', -40 / 2 - circle_Xgap[i], 300 / 2 + circle_gap[i]],
        '28': ['33', -65 / 2 - circle_Xgap[i], 30 / 2 + circle_gap[i]],
        '29': ['35', -(par.R[i] + 180) / 2 - circle_Xgap[i], 50 / 2 + circle_gap[i]],
        '30': ['41', -50 / 2 - circle_Xgap[i], 180 / 2 + circle_gap[i]],
        '31': ['43', -32 / 2 - circle_Xgap[i], 32 / 2 + circle_gap[i]],
        '32': ['44', -30 / 2 - circle_Xgap[i], 30 / 2 + circle_gap[i]]
    }
    part_circle_15_position = {
        '1': ['1', -par.B_15[i] / 2 - circle_Xgap[i], par.H[i] / 2 + circle_15_gap[i]],
        '2': ['2', -par.B_15[i] / 2 - circle_Xgap[i], par.H[i] / 2 + circle_15_gap[i]],
        '3': ['3', -(par.R_15[i] + 180) / 2 - circle_Xgap[i], par.FRAME44_height[i] / 2 + circle_15_gap[i]],
        '4': ['4', -(par.R_15[i] + 180) / 2 - circle_Xgap[i], par.FRAME44_height[i] / 2 + circle_15_gap[i]],
        '5': ['5', -(par.R_15[i] + 180) / 2 - circle_Xgap[i], par.FRAME_41_depth[i] / 2 + circle_15_gap[i]],
        '6': ['6', -182 / 2 - circle_Xgap[i], 80 / 2 + circle_15_gap[i]],
        '7': ['7', -(par.R_15[i] + 180) / 2 - circle_Xgap[i], 50 / 2 + circle_15_gap[i]],
        '8': ['8', -par.R_15[i] / 2 - circle_Xgap[i], 476 / 2 + circle_15_gap[i]],
        '9': ['8-1', -par.R_15[i] / 2 - circle_Xgap[i], 150 / 2 + circle_15_gap[i] + 50],
        '10': ['10', -par.FRAME_7_15_width[i] / 2 - circle_Xgap[i], 300 / 2 + circle_15_gap[i]],
        '11': ['11', -125 / 2 - circle_Xgap[i], 270 / 2 + circle_15_gap[i]],
        '12': ['12', -par.R_15[i] / 2 - circle_Xgap[i], 429 / 2 + circle_15_gap[i]],
        '13': ['13', -50 / 2 - circle_Xgap[i], 74 / 2 + circle_15_gap[i]],
        '14': ['14', -99.35 / 2 - circle_Xgap[i], 35 / 2 + circle_15_gap[i]],
        '15': ['15A', -40 / 2 - circle_Xgap[i], 150 / 2 + circle_15_gap[i] + 50],
        '16': ['15', -40 / 2 - circle_Xgap[i], 150 / 2 + circle_15_gap[i]],
        '17': ['16', -99.35 / 2 - circle_Xgap[i], 35 / 2 + circle_15_gap[i]],
        '18': ['18', -290 / 2 - circle_Xgap[i], 145 / 2 + circle_15_gap[i]],
        '19': ['19', -par.FRAME_11_15_width[i] / 2 - circle_Xgap[i], 90 / 2 + circle_15_gap[i]],
        '20': ['20', -145 / 2 - circle_Xgap[i], 300 / 2 + circle_15_gap[i]],
        '21': ['21', -75 / 2 - circle_Xgap[i], 75 / 2 + circle_15_gap[i]],
        '22': ['22', -(par.R_15[i] + 180) / 2 - circle_Xgap[i], 50 / 2 + circle_15_gap[i]],
        '23': ['24', -85 / 2 - circle_Xgap[i], par.FRAME_13_depth[i] / 2 + circle_15_gap[i]],
        '24': ['25', -75 / 2 - circle_Xgap[i], 75 / 2 + circle_15_gap[i]],
        '25': ['27', -460 / 2 - circle_Xgap[i], 280 / 2 + circle_15_gap[i] - 150],
        '26': ['29', -140 / 2 - circle_Xgap[i], 80 / 2 + circle_15_gap[i]],
        '27': ['30', -40 / 2 - circle_Xgap[i], 300 / 2 + circle_15_gap[i]],
        '28': ['33', -65 / 2 / 2 - circle_Xgap[i], 30 / 2 + circle_15_gap[i]],
        '29': ['35', -(par.R_15[i] - 180) / 2 - circle_Xgap[i], 50 / 2 + circle_15_gap[i]],
        '30': ['41', -50 / 2 - circle_Xgap[i], 180 / 2 + circle_15_gap[i]],
        '31': ['43', -32 / 2 - circle_Xgap[i], 32 / 2 + circle_15_gap[i]],
        '32': ['44', -30 / 2 - circle_Xgap[i], 30 / 2 + circle_15_gap[i]]
    }
    for x in projection_file_name_list:
        mprog.import_part(path, x)
        number = 1
        part_view_number = str(number)
        if l == 1:
            if x == 'FRAME1':
                p = 4  # 左視圖投影
                p2 = 5  # 右視圖投影
                X1 = (ALL_range[0][0] + ALL_range[0][1]) / 4 + 10 + 25  # 第一個投影X中心+尺寸標註預留量
                X2 = (ALL_range[0][0] + ALL_range[0][1]) * 3 / 4  # 第二個投影X中心(X中心為將X範圍分成兩塊之中心)
                Y1 = (ALL_range[0][2] + ALL_range[0][3]) * 3 / 4  # 第一個投影Y中心
                Y2 = (ALL_range[0][2] + ALL_range[0][3]) * 3 / 4  # 第二個投影Y中心(Y中心為將Y範圍分成兩塊之上半部中心)
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['1'][0], part_circle_position['1'][1],
                                           part_circle_position['1'][2], '1', scale)
                draft.add_dimension_to_view('FRAME1_1', '1', 0, par.B[i] / 2 + alpha, (par.H[i] - 72) / 2 + alpha,
                                            par.B[i] / 2, -(par.H[i] - 72) / 2, 90)
                draft.add_dimension_to_view('FRAME1_1', '2', 0, -par.B[i] / 2, -(par.H[i] - 72) / 2, par.B[i] / 2,
                                            -(par.H[i] - 72) / 2, 0)
                mprog.save_file_part(path, x)
            elif x == 'FRAME2':
                p = 4  # 左視圖投影
                p2 = 5  # 右視圖投影
                X1 = (ALL_range[0][0] + ALL_range[0][1]) / 4 + 10 + 25  # 第一個投影X中心+尺寸標註預留量
                X2 = (ALL_range[0][0] + ALL_range[0][1]) * 3 / 4  # 第二個投影X中心(X中心為將X範圍分成兩塊之中心)
                Y1 = (ALL_range[0][2] + ALL_range[0][3]) / 4  # 第一個投影Y中心
                Y2 = (ALL_range[0][2] + ALL_range[0][3]) / 4  # 第二個投影Y中心(Y中心為將Y範圍分成兩塊之上半部中心)
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['2'][0], part_circle_position['2'][1],
                                           part_circle_position['2'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME30':
                p = 0  # 前視圖投影
                p2 = 5  # 右視圖投影
                p3 = 3  # 下視圖投影
                X1 = ALL_range[1][1] + (par.R[i] + 180) * (scale) / 2 + gap  # BOX_2_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[1][0] - (40) * (scale) / 2 - gap  # BOX_2_Xmax-FRAME30厚度一半-間隙
                X3 = ALL_range[1][1] + (par.R[i] + 180) * (scale) / 2 + gap  # BOX_2_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[1][2] - (par.FRAME44_height[i]) * (
                    scale) / 2 - gap  # BOX_2_Ymax-前視圖高度-間隙(取完整高度是為了使零件圖與右壁板高度相近)
                Y2 = ALL_range[1][2] - (par.FRAME44_height[i]) * (
                    scale) / 2 - gap  # BOX_2_Ymax-前視圖高度-間隙(取完整高度是為了使零件圖與右壁板高度相近)
                Y3 = ALL_range[1][3] + (40) * (scale) / 2 + gap  # BOX_2_Ymin+FRAME30厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p3, X3, Y3, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['3'][0], part_circle_position['3'][1],
                                           part_circle_position['3'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME44':
                p = 0  # 前視圖投影
                p2 = 3  # 下視圖投影
                X1 = ALL_range[2][1] + (par.R[i] + 180) * (scale) / 2 + gap  # box_3_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[2][1] + (par.R[i] + 180) * (scale) / 2 + gap  # box_3_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[2][2] - (par.FRAME44_height[i]) * (scale) / 2 - gap  # box_3_Ymax+FRAME44高度一半-間隙
                Y2 = ALL_range[2][3] + 40 * (scale) / 2 + gap  # box_3_Ymin+FRAME44厚度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['4'][0], part_circle_position['4'][1],
                                           part_circle_position['4'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME41':
                p = 6  # 上視圖(左轉向)投影
                p2 = 5  # 右視圖投影
                X1 = ALL_range[3][1] + (par.R[i] + 180) * (scale) / 2 + gap  # box_4_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[3][1] + (par.R[i] + 180) * (scale) / 2 + gap  # box_4_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[3][2] - (par.FRAME_41_depth[i]) * (scale) / 2 - gap  # box_3_Ymax+FRAME41高度一半-間隙
                Y2 = ALL_range[3][3] + 22 * (scale) / 2 + gap  # box_3_Yin+FRAME41厚度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['5'][0], part_circle_position['5'][1],
                                           part_circle_position['5'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME34':
                p2 = 1  # 後視圖投影
                p = 5  # 右視圖投影
                p3 = 7  # 下視圖翻轉180度投影
                X2 = ALL_range[4][0] - 182 * scale / 2 - gap  # box_5_Xmax-FRAME34寬度一半-間隙
                X1 = ALL_range[4][1] + 40 * scale / 2 + gap  # box_5_Xmin+FRAME34厚度一半+間隙
                X3 = ALL_range[4][0] - 182 * scale / 2 - gap  # box_5_Xmax-FRAME34寬度一半-間隙
                Y2 = ALL_range[4][2] - 80 * scale / 2 - gap  # box_5_Ymax-FRAME34深度一半-間隙
                Y1 = ALL_range[4][2] - 80 * scale / 2 - gap  # box_5_Ymax-FRAME34深度一半-間隙
                Y3 = ALL_range[4][3] + 40 * scale / 2 + gap  # box_5_Ymax+FRAME34厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p3, X3, Y3, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['6'][0], part_circle_position['6'][1],
                                           part_circle_position['6'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME9':
                p = 2  # 上視圖投影
                p2 = 0  # 前視圖投影
                X1 = ALL_range[5][1] + (par.R[i] + 180) * scale / 2 + gap  # box_6_Xmin+FRAME9寬度一半+間隙
                X2 = ALL_range[5][1] + (par.R[i] + 180) * scale / 2 + gap  # box_6_Xmin+FRAME9寬度一半+間隙
                Y1 = ALL_range[5][2] - 50 * scale / 2 - gap  # box_6_Ymax-FRAME9厚度一半-間隙
                Y2 = ALL_range[5][3] + (par.Z[i] - par.T[i] - 40) * scale / 2 + gap  # box_6_Ymin+FRAME9高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['7'][0], part_circle_position['7'][1],
                                           part_circle_position['7'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME45':
                p = 2  # 上視圖投影
                p2 = 9  # 右側視圖(左旋轉)投影
                X1 = ALL_range[6][1] + par.R[i] * scale / 2 + gap  # box_7_Xmin+FRAME45寬度一半+間隙
                X2 = ALL_range[6][0] - 19 * scale / 2 - gap  # box_7_Xmax+FRAME45厚度一半-間隙
                Y1 = ALL_range[6][2] - 476 * scale / 2 - gap  # box_7_Ymax-FRAME45深度一半-間隙
                Y2 = ALL_range[6][2] - 476 * scale / 2 - gap  # box_7_Ymax-FRAME45深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['8'][0], part_circle_position['8'][1],
                                           part_circle_position['8'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME43':
                p = 0  # 前視圖投影
                X1 = ALL_range[7][1] + par.R[i] * scale / 2 + gap  # box_8_Xmin+FRAME43寬度一半+間隙
                Y1 = ALL_range[7][2] - 150 * scale / 2 - gap  # box_8_Ymax-FRAME45深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['9'][0], part_circle_position['9'][1],
                                           part_circle_position['9'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME7':
                p = 8  # 下視圖(右旋)投影
                X1 = ALL_range[8][1] + par.FRAME_7_width[i] * scale / 2 + gap  # box_9_Xmin+FRAME7寬度一半+間隙
                Y1 = ALL_range[8][2] - 300 * scale / 2 - gap  # box_9_Ymax-FRAME7深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['10'][0], part_circle_position['10'][1],
                                           part_circle_position['10'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'MAIN_GEAR2':
                p = 5  # 右視圖投影
                X1 = ALL_range[9][1] + 125 * scale / 2 + gap  # box_10_Xmin+MAIN_GEAR2深度一半+間隙
                Y1 = ALL_range[9][2] - 270 * scale / 2 - gap  # box_10_Ymax-MAIN_GEAR2高度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['11'][0], part_circle_position['11'][1],
                                           part_circle_position['11'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME32':
                p = 6  # 上視圖(左旋)投影
                p2 = 10  # 背視圖(左旋)投影
                X1 = ALL_range[10][1] + par.R[i] * scale / 2 + gap  # box_11_Xmin+FRAME32寬度一半+間隙
                X2 = ALL_range[10][0] - 19 * scale / 2 - gap  # box_11_Xmax-FRAME32厚度一半-間隙
                Y1 = ALL_range[10][2] - 429 * scale / 2 - gap  # box_11_Ymax-FRAME深度一半-間隙
                Y2 = ALL_range[10][2] - 429 * scale / 2 - gap  # box_11_Ymax-FRAME深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['12'][0], part_circle_position['12'][1],
                                           part_circle_position['12'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME47':
                p = 6  # 上視圖(左旋)投影
                p2 = 1  # 背視圖投影
                X1 = ALL_range[11][1] + 50 * scale / 2 + gap  # box_12_Xmin+FRAME47寬度一半+間隙
                X2 = ALL_range[11][0] - 19 * scale / 2 - gap  # box_12_Xmax-FRAME47厚度一半-間隙
                Y1 = ALL_range[11][2] - 74 * scale / 2 - gap  # box_12_Ymax-FRAME47深度一半-間隙
                Y2 = ALL_range[11][2] - 74 * scale / 2 - gap  # box_12_Ymax-FRAME47深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['13'][0], part_circle_position['13'][1],
                                           part_circle_position['13'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME23':
                p = 6  # 上視圖(左旋)投影
                p2 = 5  # 右視圖投影
                X1 = ALL_range[12][1] + 99.35 * scale / 2 + gap  # box_13_Xmin+FRAME23深度一半+間隙
                X2 = ALL_range[12][1] + 99.35 * scale / 2 + gap  # box_13_Xmin+FRAME23深度一半+間隙
                Y1 = ALL_range[12][2] - 35 * scale / 2 - gap  # box_13_Ymax-FRAME23寬度一半-間隙
                Y2 = ALL_range[12][3] + 150 * scale / 2 + gap  # box_13_Ymin+FRAME23高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['14'][0], part_circle_position['14'][1],
                                           part_circle_position['14'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME31':
                p = 0  # 前視圖投影
                p2 = 5  # 右視圖投影
                p3 = 3  # 下視圖投影
                p4 = 0  # 前視圖投影
                p5 = 3  # 下視圖投影
                X1 = ALL_range[13][1] + 40 * scale / 2 + gap  # box_14_Xmin+FRAME31寬度一半+間隙
                X2 = (ALL_range[13][1] + ALL_range[13][0]) / 2  # box_14_X中心
                X3 = ALL_range[13][1] + 40 * scale / 2 + gap  # box_14_Xmin+FRAME31寬度一半+間隙
                X4 = ALL_range[13][0] - 40 * scale / 2 - gap  # box_14_Xmax-FRAME31寬度一半-間隙
                X5 = ALL_range[13][0] - 40 * scale / 2 - gap  # box_14_Xmax-FRAME31寬度一半-間隙
                Y1 = ALL_range[13][2] - 150 * scale / 2 - gap  # box_14_Ymax-FRAME31高度一半-間隙
                Y2 = ALL_range[13][2] - 150 * scale / 2 - gap  # box_14_Ymax-FRAME31高度一半-間隙
                Y3 = ALL_range[13][3] + 45 * scale / 2 + gap  # box_14_Ymin+FRAME31厚度一半+間隙
                Y4 = ALL_range[13][2] - 150 * scale / 2 - gap  # box_14_Ymax-FRAME31高度一半-間隙
                Y5 = ALL_range[13][3] + 45 * scale / 2 + gap  # box_14_Ymin+FRAME31厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p3, X3, Y3, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p4, X4, Y4, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p5, X5, Y5, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['15'][0], part_circle_position['15'][1],
                                           part_circle_position['15'][2], '1', scale)
                DP.Parts_drafting_balloons(x, part_circle_position['16'][0], part_circle_position['16'][1],
                                           part_circle_position['16'][2], '4', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME24':
                p = 6  # 上視圖(左旋)投影
                p2 = 1  # 背視圖投影
                p3 = 5  # 右視圖投影
                X1 = ALL_range[14][1] + 99.35 * scale / 2 + gap  # box_15Xmin+FRAME24深度一半+間隙
                X2 = ALL_range[14][0] - 35 * scale / 2 + gap  # box_15Xmax-FRAME24寬度一半-間隙
                X3 = ALL_range[14][1] + 99.35 * scale / 2 + gap  # box_15Xmin+FRAME24深度一半+間隙
                Y1 = ALL_range[14][2] - 35 * scale / 2 - gap  # box_15Ymax-FRAME24寬度一半-間隙
                Y2 = ALL_range[14][3] + 150 * scale / 2 + gap  # box_15Ymin+FRAME24高度一半+間隙
                Y3 = ALL_range[14][3] + 150 * scale / 2 + gap  # box_15Ymin+FRAME24高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p3, X3, Y3, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['17'][0], part_circle_position['17'][1],
                                           part_circle_position['17'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME38':
                p = 11  # 上視圖(右旋)投影
                X1 = ALL_range[15][1] + 290 * scale / 2 + gap  # box_16_Xmin+FRAME38寬度一半+間隙
                Y1 = ALL_range[15][2] - 145 * scale / 2 - gap  # box_16_Ymax-FRAME38深度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['18'][0], part_circle_position['18'][1],
                                           part_circle_position['18'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME11':
                p2 = 5  # 右視圖投影
                p = 6  # 上視圖(左旋)投影
                X2 = ALL_range[16][1] + par.FRAME_11_width[i] * scale / 2 + gap  # box_17_Xmin+FRAME11寬度一半+間隙
                X1 = ALL_range[16][1] + par.FRAME_11_width[i] * scale / 2 + gap  # box_17_Xmin+FRAME11寬度一半+間隙
                Y2 = ALL_range[16][3] + par.FRAME_11_height[i] * scale / 2 + gap  # box_17_Ymin+FRAME11高度一半+間隙
                Y1 = ALL_range[16][2] - 90 * scale / 2 - gap  # box_17_Ymax-FRAME11厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['19'][0], part_circle_position['19'][1],
                                           part_circle_position['19'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME39':
                p = 12  # 上視圖(翻轉180度)投影
                p2 = 1  # 後視圖
                X1 = ALL_range[17][1] + 145 * scale / 2 + gap  # box_18_Xmin+FRAME39深度一半+間隙
                X2 = ALL_range[17][1] + 145 * scale / 2 + gap  # box_18_Xmin+FRAME39深度一半+間隙
                Y1 = ALL_range[17][2] - 300 * scale / 2 - gap  # box_18_Ymax-FRAME39寬度一半-間隙
                Y2 = ALL_range[17][3] + 19 * scale / 2 + gap  # box_18_Ymin+FRAME39厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['20'][0], part_circle_position['20'][1],
                                           part_circle_position['20'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME17':
                p = 12  # 上視圖(翻轉180度)投影
                p2 = 4  # 左視圖投影
                X1 = ALL_range[18][1] + 75 * scale / 2 + gap  # box_19_Xmin+FRAME17寬度一半+間隙
                X2 = ALL_range[18][1] + 75 * scale / 2 + gap  # box_19_Xmin+FRAME17寬度一半+間隙
                Y1 = ALL_range[18][2] - 75 * scale / 2 - gap  # box_19_Ymax-FRAME17深度一半-間隙
                Y2 = ALL_range[18][3] + 48 * scale / 2 + gap  # box_19_Ymin+FRAME17高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['21'][0], part_circle_position['21'][1],
                                           part_circle_position['21'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME3':
                p2 = 0  # 前視圖投影
                p = 2  # 上視圖投影
                X2 = ALL_range[19][1] + (par.R[i] + 180) * scale / 2 + gap  # box_20_Xmin+FRAME3寬度一半+間隙
                X1 = ALL_range[19][1] + (par.R[i] + 180) * scale / 2 + gap  # box_20_Xmin+FRAME3寬度一半+間隙
                Y2 = ALL_range[19][3] + (par.Z[i] - par.T[i] - 40) * scale / 2 + gap  # box_20_Ymin+FRAME3高度一半+間隙
                Y1 = ALL_range[19][2] - 50 * scale / 2 - gap  # box_20_Ymax-FRAME3厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['22'][0], part_circle_position['22'][1],
                                           part_circle_position['22'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME13':
                p = 2  # 上視圖投影
                p2 = 9  # 右視圖(左旋)投影
                X1 = ALL_range[20][1] + 85 * scale / 2 + gap  # box_21_Xmin+FRAME13寬度一半+間隙
                X2 = ALL_range[20][0] - 55 * scale / 2 - gap  # box_21_Xmax-FRAME13高度一半-間隙
                Y1 = ALL_range[20][2] - par.FRAME_13_depth[i] * scale / 2 - gap  # box_21_Ymax-FRAME13深度一半-間隙
                Y2 = ALL_range[20][2] - par.FRAME_13_depth[i] * scale / 2 - gap  # box_21_Ymax-FRAME13深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['23'][0], part_circle_position['23'][1],
                                           part_circle_position['23'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME14':
                p = 6  # 上視圖(左旋)投影
                p2 = 4  # 左視圖投影
                X1 = ALL_range[21][1] + 75 * scale / 2 + gap  # box_22_Xmin+FRAME14深度一半+間隙
                X2 = ALL_range[21][1] + 75 * scale / 2 + gap  # box_22_Xmin+FRAME14深度一半+間隙
                Y1 = ALL_range[21][2] - 75 * scale / 2 - gap  # box_22_Ymax-FRAME14寬度一半-間隙
                Y2 = ALL_range[21][3] + 70 * scale / 2 + gap  # box_22_Ymin+FRAME14高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['24'][0], part_circle_position['24'][1],
                                           part_circle_position['24'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME22':
                p = 4  # 左視圖投影
                X1 = (ALL_range[22][1] + ALL_range[22][0]) / 2  # box_23_Xmin+FRAME22寬度一半+間隙
                Y1 = (ALL_range[22][2] + ALL_range[22][3]) / 2  # box_23_Ymax-FRAME22高度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['25'][0], part_circle_position['25'][1],
                                           part_circle_position['25'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME37':
                p = 1  # 後視圖投影
                X1 = ALL_range[23][1] + 140 * scale / 2 + gap  # box_24_Xmin+FRAME37寬度一半+間隙
                Y1 = ALL_range[23][2] - 80 * scale / 2 - gap  # box_24_Ymax-FRAME37高度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['26'][0], part_circle_position['26'][1],
                                           part_circle_position['26'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME8':
                p2 = 13  # 下視圖(左旋)投影
                p = 10  # 後視圖(左旋)投影
                X2 = ALL_range[24][0] - par.FRAME_8_width[i] * scale / 2 - gap  # box_25_Xmax-FRAME8寬度一半-間隙
                X1 = ALL_range[24][1] + 40 * scale / 2 + gap  # box_25_Xmin+FRAME8厚度一半+間隙
                Y2 = ALL_range[24][2] - 300 * scale / 2 - gap  # box_25_Ymax-FRAME8深度一半-間隙
                Y1 = ALL_range[24][2] - 300 * scale / 2 - gap  # box_25_Ymax-FRAME8深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['27'][0], part_circle_position['27'][1],
                                           part_circle_position['27'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME29':
                p2 = 7  # 下視圖(180度)投影
                p = 9  # 右視圖(左旋)投影
                X2 = ALL_range[25][0] - (par.R[i] + 180) * scale / 2 - gap  # box_26_Xmax-FRAME29寬度一半-間隙
                X1 = ALL_range[25][1] + 65 * scale / 2 + gap  # box_26_Xmin+FRAME29深度一半+間隙
                Y2 = ALL_range[25][2] - 30 * scale / 2 - gap  # box_26_Ymax-FRAME29厚度一半-間隙
                Y1 = ALL_range[25][2] - 30 * scale / 2 - gap  # box_26_Ymax-FRAME29厚度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['28'][0], part_circle_position['28'][1],
                                           part_circle_position['28'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME20':
                p2 = 0  # 前視圖投影
                p = 2  # 上視圖投影
                X2 = ALL_range[26][1] + (par.R[i] + 180) * scale / 2 + gap  # box_27_Xmin+FRAME20寬度一半+間隙
                X1 = ALL_range[26][1] + (par.R[i] + 180) * scale / 2 + gap  # box_27_Xmin+FRAME20寬度一半+間隙
                Y1 = ALL_range[26][2] - 50 * scale / 2 - gap  # box_27_Ymax-FRAME20厚度一半-間隙
                Y2 = ALL_range[26][3] + par.FRAME20_H[i] * scale / 2 + gap  # box_27_Ymin+FRAME20高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['29'][0], part_circle_position['29'][1],
                                           part_circle_position['29'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME33':
                p = 12  # 上視圖(轉180度)投影
                p2 = 1  # 後視圖投影
                X1 = ALL_range[27][1] + 50 * scale / 2 + gap  # box_28_Xmin+FRAME33深度一半+間隙
                X2 = ALL_range[27][1] + 50 * scale / 2 + gap  # box_28_Xmin+FRAME33深度一半+間隙
                Y1 = ALL_range[27][2] - 180 * scale / 2 - gap  # box_28_Ymax-FRAME33寬度一半-間隙
                Y2 = ALL_range[27][3] + 50 * scale / 2 + gap  # box_28_Ymin+FRAME33高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['30'][0], part_circle_position['30'][1],
                                           part_circle_position['30'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME48':
                p = 4  # 左視圖投影
                p2 = 0  # 前視圖投影
                X1 = ALL_range[28][1] + 32 * scale / 2 + gap  # box_29_Xmin+FRAME48寬度一半+間隙
                X2 = ALL_range[28][0] - 82 * scale / 2 - gap  # box_29_Xmax-FRAME48深度一半-間隙
                Y1 = ALL_range[28][2] - 32 * scale / 2 - gap  # box_29_Ymax-FRAME48高度一半-間隙
                Y2 = ALL_range[28][2] - 32 * scale / 2 - gap  # box_29_Ymax-FRAME48高度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['31'][0], part_circle_position['31'][1],
                                           part_circle_position['31'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME49':
                p = 0  # 前視圖投影
                p2 = 5  # 右視圖投影
                X1 = ALL_range[29][1] + 30 * scale / 2 + gap  # box_30_Xmin+FRAME49寬度一半+間隙
                X2 = ALL_range[29][0] - 54 * scale / 2 - gap  # box_30_Xmax-FRAME49深度一半-間隙
                Y1 = ALL_range[29][2] - 30 * scale / 2 - gap  # box_30_Ymax-FRAME49高度一半-間隙
                Y2 = ALL_range[29][2] - 30 * scale / 2 - gap  # box_30_Ymax-FRAME49高度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_position['32'][0], part_circle_position['32'][1],
                                           part_circle_position['32'][2], '1', scale)
                mprog.save_file_part(path, x)
            else:
                break
        elif l == 0:
            if x == 'FRAME1':
                p = 4  # 左視圖投影
                p2 = 5  # 右視圖投影
                X1 = (ALL_range[0][0] + ALL_range[0][1]) / 4 + 10  # 第一個投影X中心+尺寸標註預留量
                X2 = (ALL_range[0][0] + ALL_range[0][1]) * 3 / 4  # 第二個投影X中心(X中心為將X範圍分成兩塊之中心)
                Y1 = (ALL_range[0][2] + ALL_range[0][3]) * 3 / 4  # 第一個投影Y中心
                Y2 = (ALL_range[0][2] + ALL_range[0][3]) * 3 / 4  # 第二個投影Y中心(Y中心為將Y範圍分成兩塊之上半部中心)
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['1'][0], part_circle_15_position['1'][1],
                                           part_circle_15_position['1'][2], '1', scale)
                draft.add_dimension_to_view('FRAME1_1', '1', 0, par.B_15[i] / 2, (par.H[i] - 80) / 2, par.B_15[i] / 2,
                                            -(par.H[i] - 80) / 2, 90)
                draft.add_dimension_to_view('FRAME1_1', '2', 0, -par.B_15[i] / 2, -(par.H[i] - 80) / 2, par.B_15[i] / 2,
                                            -(par.H[i] - 80) / 2, 0)
                mprog.save_file_part(path, x)
            elif x == 'FRAME2':
                p = 4  # 左視圖投影
                p2 = 5  # 右視圖投影
                X1 = (ALL_range[0][0] + ALL_range[0][1]) / 4 + 10  # 第一個投影X中心+尺寸標註預留量
                X2 = (ALL_range[0][0] + ALL_range[0][1]) * 3 / 4  # 第二個投影X中心(X中心為將X範圍分成兩塊之中心)
                Y1 = (ALL_range[0][2] + ALL_range[0][3]) / 4  # 第一個投影Y中心
                Y2 = (ALL_range[0][2] + ALL_range[0][3]) / 4  # 第二個投影Y中心(Y中心為將Y範圍分成兩塊之上半部中心)
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['2'][0], part_circle_15_position['2'][1],
                                           part_circle_15_position['2'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME30':
                p = 0  # 前視圖投影
                p2 = 5  # 右視圖投影
                p3 = 3  # 下視圖投影
                X1 = ALL_range[1][1] + (par.R_15[i] + 180) * (scale) / 2 + gap  # BOX_2_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[1][0] - (40) * (scale) / 2 - gap  # BOX_2_Xmax-FRAME30厚度一半-間隙
                X3 = ALL_range[1][1] + (par.R_15[i] + 180) * (scale) / 2 + gap  # BOX_2_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[1][2] - (par.FRAME44_height[i]) * (
                    scale) / 2 - gap  # BOX_2_Ymax-前視圖高度-間隙(取完整高度是為了使零件圖與右壁板高度相近)
                Y2 = ALL_range[1][2] - (par.FRAME44_height[i]) * (
                    scale) / 2 - gap  # BOX_2_Ymax-前視圖高度-間隙(取完整高度是為了使零件圖與右壁板高度相近)
                Y3 = ALL_range[1][3] + (40) * (scale) / 2 + gap  # BOX_2_Ymin+FRAME30厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p3, X3, Y3, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['3'][0], part_circle_15_position['3'][1],
                                           part_circle_15_position['3'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME44':
                p = 0  # 前視圖投影
                p2 = 3  # 下視圖投影
                X1 = ALL_range[2][1] + (par.R_15[i] + 180) * (scale) / 2 + gap  # box_3_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[2][1] + (par.R_15[i] + 180) * (scale) / 2 + gap  # box_3_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[2][2] - (par.FRAME44_height[i]) * (scale) / 2 - gap  # box_3_Ymax+FRAME44高度一半-間隙
                Y2 = ALL_range[2][3] + 40 * (scale) / 2 + gap  # box_3_Ymin+FRAME44厚度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['4'][0], part_circle_15_position['4'][1],
                                           part_circle_15_position['4'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME41':
                p = 6  # 上視圖(左轉向)投影
                p2 = 5  # 右視圖投影
                X1 = ALL_range[3][1] + (par.R_15[i] + 180) * (scale) / 2 + gap  # box_4_Xmin+前視圖寬度一半+間隙
                X2 = ALL_range[3][1] + (par.R_15[i] + 180) * (scale) / 2 + gap  # box_4_Xmin+前視圖寬度一半+間隙
                Y1 = ALL_range[3][2] - (par.FRAME_41_depth[i]) * (scale) / 2 - gap  # box_3_Ymax+FRAME41高度一半-間隙
                Y2 = ALL_range[3][3] + 22 * (scale) / 2 + gap  # box_3_Yin+FRAME41厚度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['5'][0], part_circle_15_position['5'][1],
                                           part_circle_15_position['5'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME34':
                p2 = 1  # 後視圖投影
                p = 5  # 右視圖投影
                p3 = 7  # 下視圖翻轉180度投影
                X2 = ALL_range[4][0] - 182 * scale / 2 - gap  # box_5_Xmax-FRAME34寬度一半-間隙
                X1 = ALL_range[4][1] + 40 * scale / 2 + gap  # box_5_Xmin+FRAME34厚度一半+間隙
                X3 = ALL_range[4][0] - 182 * scale / 2 - gap  # box_5_Xmax-FRAME34寬度一半-間隙
                Y2 = ALL_range[4][2] - 80 * scale / 2 - gap  # box_5_Ymax-FRAME34深度一半-間隙
                Y1 = ALL_range[4][2] - 80 * scale / 2 - gap  # box_5_Ymax-FRAME34深度一半-間隙
                Y3 = ALL_range[4][3] + 40 * scale / 2 + gap  # box_5_Ymax+FRAME34厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p3, X3, Y3, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['6'][0], part_circle_15_position['6'][1],
                                           part_circle_15_position['6'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME9':
                p = 2  # 上視圖投影
                p2 = 0  # 前視圖投影
                X1 = ALL_range[5][1] + (par.R_15[i] + 180) * scale / 2 + gap  # box_6_Xmin+FRAME9寬度一半+間隙
                X2 = ALL_range[5][1] + (par.R_15[i] + 180) * scale / 2 + gap  # box_6_Xmin+FRAME9寬度一半+間隙
                Y1 = ALL_range[5][2] - 50 * scale / 2 - gap  # box_6_Ymax-FRAME9厚度一半-間隙
                Y2 = ALL_range[5][3] + (par.Z[i] - par.T[i] - 40) * scale / 2 + gap  # box_6_Ymin+FRAME9高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['7'][0], part_circle_15_position['7'][1],
                                           part_circle_15_position['7'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME45':
                p = 2  # 上視圖投影
                p2 = 9  # 右側視圖(左旋轉)投影
                X1 = ALL_range[6][1] + par.R_15[i] * scale / 2 + gap  # box_7_Xmin+FRAME45寬度一半+間隙
                X2 = ALL_range[6][0] - 19 * scale / 2 - gap  # box_7_Xmax+FRAME45厚度一半-間隙
                Y1 = ALL_range[6][2] - 476 * scale / 2 - gap  # box_7_Ymax-FRAME45深度一半-間隙
                Y2 = ALL_range[6][2] - 476 * scale / 2 - gap  # box_7_Ymax-FRAME45深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['8'][0], part_circle_15_position['8'][1],
                                           part_circle_15_position['8'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME43':
                p = 0  # 前視圖投影
                X1 = ALL_range[7][1] + par.R_15[i] * scale / 2 + gap  # box_8_Xmin+FRAME43寬度一半+間隙
                Y1 = ALL_range[7][2] - 150 * scale / 2 - gap  # box_8_Ymax-FRAME45深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['9'][0], part_circle_15_position['9'][1],
                                           part_circle_15_position['9'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME7':
                p = 8  # 下視圖(右旋)投影
                X1 = ALL_range[8][1] + par.FRAME_7_15_width[i] * scale / 2 + gap  # box_9_Xmin+FRAME7寬度一半+間隙
                Y1 = ALL_range[8][2] - 300 * scale / 2 - gap  # box_9_Ymax-FRAME7深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['10'][0], part_circle_15_position['10'][1],
                                           part_circle_15_position['10'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'MAIN_GEAR2':
                p = 5  # 右視圖投影
                X1 = ALL_range[9][1] + 125 * scale / 2 + gap  # box_10_Xmin+MAIN_GEAR2深度一半+間隙
                Y1 = ALL_range[9][2] - 270 * scale / 2 - gap  # box_10_Ymax-MAIN_GEAR2高度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['11'][0], part_circle_15_position['11'][1],
                                           part_circle_15_position['11'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME32':
                p = 6  # 上視圖(左旋)投影
                p2 = 10  # 背視圖(左旋)投影
                X1 = ALL_range[10][1] + par.R_15[i] * scale / 2 + gap  # box_11_Xmin+FRAME32寬度一半+間隙
                X2 = ALL_range[10][0] - 19 * scale / 2 - gap  # box_11_Xmax-FRAME32厚度一半-間隙
                Y1 = ALL_range[10][2] - 429 * scale / 2 - gap  # box_11_Ymax-FRAME深度一半-間隙
                Y2 = ALL_range[10][2] - 429 * scale / 2 - gap  # box_11_Ymax-FRAME深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['12'][0], part_circle_15_position['12'][1],
                                           part_circle_15_position['12'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME47':
                p = 6  # 上視圖(左旋)投影
                p2 = 1  # 背視圖投影
                X1 = ALL_range[11][1] + 50 * scale / 2 + gap  # box_12_Xmin+FRAME47寬度一半+間隙
                X2 = ALL_range[11][0] - 19 * scale / 2 - gap  # box_12_Xmax-FRAME47厚度一半-間隙
                Y1 = ALL_range[11][2] - 74 * scale / 2 - gap  # box_12_Ymax-FRAME47深度一半-間隙
                Y2 = ALL_range[11][2] - 74 * scale / 2 - gap  # box_12_Ymax-FRAME47深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['13'][0], part_circle_15_position['13'][1],
                                           part_circle_15_position['13'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME23':
                p = 6  # 上視圖(左旋)投影
                p2 = 5  # 右視圖投影
                X1 = ALL_range[12][1] + 99.35 * scale / 2 + gap  # box_13_Xmin+FRAME23深度一半+間隙
                X2 = ALL_range[12][1] + 99.35 * scale / 2 + gap  # box_13_Xmin+FRAME23深度一半+間隙
                Y1 = ALL_range[12][2] - 35 * scale / 2 - gap  # box_13_Ymax-FRAME23寬度一半-間隙
                Y2 = ALL_range[12][3] + 150 * scale / 2 + gap  # box_13_Ymin+FRAME23高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['14'][0], part_circle_15_position['14'][1],
                                           part_circle_15_position['14'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME31':
                p = 0  # 前視圖投影
                p2 = 5  # 右視圖投影
                p3 = 3  # 下視圖投影
                p4 = 0  # 前視圖投影
                p5 = 3  # 下視圖投影
                X1 = ALL_range[13][1] + 40 * scale / 2 + gap  # box_14_Xmin+FRAME31寬度一半+間隙
                X2 = (ALL_range[13][1] + ALL_range[13][0]) / 2  # box_14_X中心
                X3 = ALL_range[13][1] + 40 * scale / 2 + gap  # box_14_Xmin+FRAME31寬度一半+間隙
                X4 = ALL_range[13][0] - 40 * scale / 2 - gap  # box_14_Xmax-FRAME31寬度一半-間隙
                X5 = ALL_range[13][0] - 40 * scale / 2 - gap  # box_14_Xmax-FRAME31寬度一半-間隙
                Y1 = ALL_range[13][2] - 150 * scale / 2 - gap  # box_14_Ymax-FRAME31高度一半-間隙
                Y2 = ALL_range[13][2] - 150 * scale / 2 - gap  # box_14_Ymax-FRAME31高度一半-間隙
                Y3 = ALL_range[13][3] + 45 * scale / 2 + gap  # box_14_Ymin+FRAME31厚度一半+間隙
                Y4 = ALL_range[13][2] - 150 * scale / 2 - gap  # box_14_Ymax-FRAME31高度一半-間隙
                Y5 = ALL_range[13][3] + 45 * scale / 2 + gap  # box_14_Ymin+FRAME31厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p3, X3, Y3, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p4, X4, Y4, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p5, X5, Y5, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['15'][0], part_circle_15_position['15'][1],
                                           part_circle_15_position['15'][2], '1', scale)
                DP.Parts_drafting_balloons(x, part_circle_15_position['16'][0], part_circle_15_position['16'][1],
                                           part_circle_15_position['16'][2], '4', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME24':
                p = 6  # 上視圖(左旋)投影
                p2 = 1  # 背視圖投影
                p3 = 5  # 右視圖投影
                X1 = ALL_range[14][1] + 99.35 * scale / 2 + gap  # box_15Xmin+FRAME24深度一半+間隙
                X2 = ALL_range[14][0] - 35 * scale / 2 + gap  # box_15Xmax-FRAME24寬度一半-間隙
                X3 = ALL_range[14][1] + 99.35 * scale / 2 + gap  # box_15Xmin+FRAME24深度一半+間隙
                Y1 = ALL_range[14][2] - 35 * scale / 2 - gap  # box_15Ymax-FRAME24寬度一半-間隙
                Y2 = ALL_range[14][3] + 150 * scale / 2 + gap  # box_15Ymin+FRAME24高度一半+間隙
                Y3 = ALL_range[14][3] + 150 * scale / 2 + gap  # box_15Ymin+FRAME24高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p3, X3, Y3, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['17'][0], part_circle_15_position['17'][1],
                                           part_circle_15_position['17'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME38':
                p = 11  # 上視圖(右旋)投影
                X1 = ALL_range[15][1] + 290 * scale / 2 + gap  # box_16_Xmin+FRAME38寬度一半+間隙
                Y1 = ALL_range[15][2] - 145 * scale / 2 - gap  # box_16_Ymax-FRAME38深度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['18'][0], part_circle_15_position['18'][1],
                                           part_circle_15_position['18'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME11':
                p2 = 5  # 右視圖投影
                p = 6  # 上視圖(左旋)投影
                X2 = ALL_range[16][1] + par.FRAME_11_15_width[i] * scale / 2 + gap  # box_17_Xmin+FRAME11寬度一半+間隙
                X1 = ALL_range[16][1] + par.FRAME_11_15_width[i] * scale / 2 + gap  # box_17_Xmin+FRAME11寬度一半+間隙
                Y2 = ALL_range[16][3] + par.FRAME_11_height[i] * scale / 2 + gap  # box_17_Ymin+FRAME11高度一半+間隙
                Y1 = ALL_range[16][2] - 90 * scale / 2 - gap  # box_17_Ymax-FRAME11厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['19'][0], part_circle_15_position['19'][1],
                                           part_circle_15_position['19'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME39':
                p = 12  # 上視圖(翻轉180度)投影
                p2 = 1  # 後視圖
                X1 = ALL_range[17][1] + 145 * scale / 2 + gap  # box_18_Xmin+FRAME39深度一半+間隙
                X2 = ALL_range[17][1] + 145 * scale / 2 + gap  # box_18_Xmin+FRAME39深度一半+間隙
                Y1 = ALL_range[17][2] - 300 * scale / 2 - gap  # box_18_Ymax-FRAME39寬度一半-間隙
                Y2 = ALL_range[17][3] + 19 * scale / 2 + gap  # box_18_Ymin+FRAME39厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['20'][0], part_circle_15_position['20'][1],
                                           part_circle_15_position['20'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME17':
                p = 12  # 上視圖(翻轉180度)投影
                p2 = 4  # 左視圖投影
                X1 = ALL_range[18][1] + 75 * scale / 2 + gap  # box_19_Xmin+FRAME17寬度一半+間隙
                X2 = ALL_range[18][1] + 75 * scale / 2 + gap  # box_19_Xmin+FRAME17寬度一半+間隙
                Y1 = ALL_range[18][2] - 75 * scale / 2 - gap  # box_19_Ymax-FRAME17深度一半-間隙
                Y2 = ALL_range[18][3] + 48 * scale / 2 + gap  # box_19_Ymin+FRAME17高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['21'][0], part_circle_15_position['21'][1],
                                           part_circle_15_position['21'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME3':
                p2 = 0  # 前視圖投影
                p = 2  # 上視圖投影
                X2 = ALL_range[19][1] + (par.R_15[i] + 180) * scale / 2 + gap  # box_20_Xmin+FRAME3寬度一半+間隙
                X1 = ALL_range[19][1] + (par.R_15[i] + 180) * scale / 2 + gap  # box_20_Xmin+FRAME3寬度一半+間隙
                Y2 = ALL_range[19][3] + (par.Z[i] - par.T[i] - 40) * scale / 2 + gap  # box_20_Ymin+FRAME3高度一半+間隙
                Y1 = ALL_range[19][2] - 50 * scale / 2 - gap  # box_20_Ymax-FRAME3厚度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['22'][0], part_circle_15_position['22'][1],
                                           part_circle_15_position['22'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME13':
                p = 2  # 上視圖投影
                p2 = 9  # 右視圖(左旋)投影
                X1 = ALL_range[20][1] + 85 * scale / 2 + gap  # box_21_Xmin+FRAME13寬度一半+間隙
                X2 = ALL_range[20][0] - 55 * scale / 2 - gap  # box_21_Xmax-FRAME13高度一半-間隙
                Y1 = ALL_range[20][2] - par.FRAME_13_depth[i] * scale / 2 - gap  # box_21_Ymax-FRAME13深度一半-間隙
                Y2 = ALL_range[20][2] - par.FRAME_13_depth[i] * scale / 2 - gap  # box_21_Ymax-FRAME13深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['23'][0], part_circle_15_position['23'][1],
                                           part_circle_15_position['23'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME14':
                p = 6  # 上視圖(左旋)投影
                p2 = 4  # 左視圖投影
                X1 = ALL_range[21][1] + 75 * scale / 2 + gap  # box_22_Xmin+FRAME14深度一半+間隙
                X2 = ALL_range[21][1] + 75 * scale / 2 + gap  # box_22_Xmin+FRAME14深度一半+間隙
                Y1 = ALL_range[21][2] - 75 * scale / 2 - gap  # box_22_Ymax-FRAME14寬度一半-間隙
                Y2 = ALL_range[21][3] + 70 * scale / 2 + gap  # box_22_Ymin+FRAME14高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['24'][0], part_circle_15_position['24'][1],
                                           part_circle_15_position['24'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME22':
                p = 4  # 左視圖投影
                X1 = ALL_range[22][1] + 460 * scale / 2 + gap  # box_23_Xmin+FRAME22寬度一半+間隙
                Y1 = ALL_range[22][2] - 280 * scale / 2 + gap  # box_23_Ymax-FRAME22高度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['25'][0], part_circle_15_position['25'][1],
                                           part_circle_15_position['25'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME37':
                p = 1  # 後視圖投影
                X1 = ALL_range[23][1] + 140 * scale / 2 + gap  # box_24_Xmin+FRAME37寬度一半+間隙
                Y1 = ALL_range[23][2] - 80 * scale / 2 - gap  # box_24_Ymax-FRAME37高度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['26'][0], part_circle_15_position['26'][1],
                                           part_circle_15_position['26'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME8':
                p2 = 13  # 下視圖(左旋)投影
                p = 10  # 後視圖(左旋)投影
                X2 = ALL_range[24][0] - par.FRAME_8_15_width[i] * scale / 2 - gap  # box_25_Xmax-FRAME8寬度一半-間隙
                X1 = ALL_range[24][1] + 40 * scale / 2 + gap  # box_25_Xmin+FRAME8厚度一半+間隙
                Y2 = ALL_range[24][2] - 300 * scale / 2 - gap  # box_25_Ymax-FRAME8深度一半-間隙
                Y1 = ALL_range[24][2] - 300 * scale / 2 - gap  # box_25_Ymax-FRAME8深度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['27'][0], part_circle_15_position['27'][1],
                                           part_circle_15_position['27'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME29':
                p2 = 7  # 下視圖(180度)投影
                p = 9  # 右視圖(左旋)投影
                X2 = ALL_range[25][0] - (par.R_15[i] + 180) * scale / 2 - gap  # box_26_Xmax-FRAME29寬度一半-間隙
                X1 = ALL_range[25][1] + 65 * scale / 2 + gap  # box_26_Xmin+FRAME29深度一半+間隙
                Y2 = ALL_range[25][2] - 30 * scale / 2 - gap  # box_26_Ymax-FRAME29厚度一半-間隙
                Y1 = ALL_range[25][2] - 30 * scale / 2 - gap  # box_26_Ymax-FRAME29厚度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['28'][0], part_circle_15_position['28'][1],
                                           part_circle_15_position['28'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME20':
                p2 = 0  # 前視圖投影
                p = 2  # 上視圖投影
                X2 = ALL_range[26][1] + (par.R_15[i] + 180) * scale / 2 + gap  # box_27_Xmin+FRAME20寬度一半+間隙
                X1 = ALL_range[26][1] + (par.R_15[i] + 180) * scale / 2 + gap  # box_27_Xmin+FRAME20寬度一半+間隙
                Y1 = ALL_range[26][2] - 50 * scale / 2 - gap  # box_27_Ymax-FRAME20厚度一半-間隙
                Y2 = ALL_range[26][3] + par.FRAME20_H[i] * scale / 2 + gap  # box_27_Ymin+FRAME20高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['29'][0], part_circle_15_position['29'][1],
                                           part_circle_15_position['29'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME33':
                p = 12  # 上視圖(轉180度)投影
                p2 = 1  # 後視圖投影
                X1 = ALL_range[27][1] + 50 * scale / 2 + gap  # box_28_Xmin+FRAME33深度一半+間隙
                X2 = ALL_range[27][1] + 50 * scale / 2 + gap  # box_28_Xmin+FRAME33深度一半+間隙
                Y1 = ALL_range[27][2] - 180 * scale / 2 - gap  # box_28_Ymax-FRAME33寬度一半-間隙
                Y2 = ALL_range[27][3] + 50 * scale / 2 + gap  # box_28_Ymin+FRAME33高度一半+間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['30'][0], part_circle_15_position['30'][1],
                                           part_circle_15_position['30'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME48':
                p = 4  # 左視圖投影
                p2 = 0  # 前視圖投影
                X1 = ALL_range[28][1] + 32 * scale / 2 + gap  # box_29_Xmin+FRAME48寬度一半+間隙
                X2 = ALL_range[28][0] - 82 * scale / 2 - gap  # box_29_Xmax-FRAME48深度一半-間隙
                Y1 = ALL_range[28][2] - 32 * scale / 2 - gap  # box_29_Ymax-FRAME48高度一半-間隙
                Y2 = ALL_range[28][2] - 32 * scale / 2 - gap  # box_29_Ymax-FRAME48高度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['31'][0], part_circle_15_position['31'][1],
                                           part_circle_15_position['31'][2], '1', scale)
                mprog.save_file_part(path, x)
            elif x == 'FRAME49':
                p = 0  # 前視圖投影
                p2 = 5  # 右視圖投影
                X1 = ALL_range[29][1] + 30 * scale / 2 + gap  # box_30_Xmin+FRAME49寬度一半+間隙
                X2 = ALL_range[29][0] - 54 * scale / 2 - gap  # box_30_Xmax-FRAME49深度一半-間隙
                Y1 = ALL_range[29][2] - 30 * scale / 2 - gap  # box_30_Ymax-FRAME49高度一半-間隙
                Y2 = ALL_range[29][2] - 30 * scale / 2 - gap  # box_30_Ymax-FRAME49高度一半-間隙
                DP.Material_diagram_projection(p, X1, Y1, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Material_diagram_projection(p2, X2, Y2, x, scale, part_view_number, x + "_" + part_view_number)
                number += 1
                part_view_number = str(number)
                DP.Parts_drafting_balloons(x, part_circle_15_position['32'][0], part_circle_15_position['32'][1],
                                           part_circle_15_position['32'][2], '1', scale)
                mprog.save_file_part(path, x)
            else:
                break
