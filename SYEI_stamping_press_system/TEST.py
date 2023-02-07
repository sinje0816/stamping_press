import win32com.client as win32
import main_program as mprog
import datetime, time, math
import drafting as draft
import parameter as par

i = 5

# mprog.switch_window()

# 隱藏機架外零件
# for part_name in par.hide_part_name:
#     mprog.hide_show_part(part_name, 1)

# #投影
# mprog.switch_window()
drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale = draft.drafting_welding_view_parameter_calculation(
    par.A[i], par.B[i], par.H[i])
draft.change_Drawing_scale(1 / scale)
# # 副程式名稱須注意是否與立體圖相符
# draft.Front_View_Drawing('SN1-110', drafting_down_Coordinate_Position['Right View'][0], drafting_down_Coordinate_Position['Right View'][1], scale)  # 左側試圖
# draft.Right_View_Drawing('SN1-110', drafting_down_Coordinate_Position['Front View'][0], drafting_down_Coordinate_Position['Front View'][1], scale)  # 前視圖
# draft.Right_Top_View_Drawing('SN1-110', drafting_up_Coordinate_Position['Top View'][0], drafting_up_Coordinate_Position['Top View'][1], scale)  # 上視圖


# 剖面圖座標
Section_Coordinate = {'A-A': [[float(0), float(-2900), float(0), float(600)], 1],
                      'B-B': [[float(par.B[i] + 100), float(520), float(par.B[i] + 100), float(-1200)], 0],
                      'C-C': [[float(-195), float(-1460), float(2286), float(par.B[i] + 100)], 0],
                      'D-D': [[float(-par.A[i] / 2 - 100), float(-par.H[i] / 2), float(par.A[i] / 2 + 100), float(-par.H[i] / 2)], 0],
                      'E-E': [[float(-100), float(-(par.S[i] + par.Z[i]) + 100), float(par.B[i] / 3), float(-(par.S[i] + par.Z[i]) + 100)], 1],
                      'F-F': [[float(-272), float(- 176 / 2), float(-272), float(-176 * 1.5)], 0],
                      'G-G': [[float(par.B[i] + 100), float(-(par.S[i] + par.Z[i]) + 150), float(par.B[i] + 100), float(-(par.S[i] + par.Z[i]) - 150)], 1]}

draft.section('Front view', drafting_down_Coordinate_Position['section A-A'][0],
              drafting_down_Coordinate_Position['section A-A'][1], scale, Section_Coordinate['A-A'][0], Section_Coordinate['A-A'][1])
draft.section('Right view', drafting_down_Coordinate_Position['section B-B'][0],
              drafting_down_Coordinate_Position['section B-B'][1], scale, Section_Coordinate['B-B'][0], Section_Coordinate['B-B'][1])
draft.section('Section view A-A', drafting_up_Coordinate_Position['section C-C'][0],
              drafting_up_Coordinate_Position['section C-C'][1], scale, Section_Coordinate['C-C'][0], Section_Coordinate['C-C'][1])
draft.section('Section view C-C', drafting_up_Coordinate_Position['section D-D'][0],
              drafting_up_Coordinate_Position['section D-D'][1], scale, Section_Coordinate['D-D'][0], Section_Coordinate['D-D'][1])
draft.section('Right view', drafting_up_Coordinate_Position['section E-E'][0],
              drafting_up_Coordinate_Position['section E-E'][1], scale, Section_Coordinate['E-E'][0], Section_Coordinate['E-E'][1])

# # 焊接符號生成
# draft.symbol_of_weld('Front view', 3, -par.A[i] / 2 + 50 + 105, par.H[i] - par.S[i] - par.Z[i] - 32, 1)

# # 圈碼圖生成
# draft.balloons('Front view', '1', -par.A[i] / 2 - 100 * par.cos30, (par.H[i] - par.S[i] - par.Z[i]) * par.sin30 / 2, -par.A[i] / 2, (par.H[i] - par.S[i] - par.Z[i]) / 2, -par.A[i] / 2 - 100 * par.cos30)
