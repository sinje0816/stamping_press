import win32com.client as win32
import main_program as mprog
import datetime, time, math
import drafting as draft
import parameter as par

i = 5


mprog.switch_window()
# 隱藏機架外零件
for part_name in par.hide_part_name:
    mprog.hide_show_part(part_name, 1)
#投影
mprog.switch_window()
drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale = draft.drafting_welding_view_parameter_calculation(par.A[i], par.B[i], par.H[i])
print(drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale)
draft.change_Drawing_scale(1 / scale)
# 副程式名稱須注意是否與立體圖相符
draft.Front_View_Drawing('SN1-110', drafting_down_Coordinate_Position['Right View'][0], drafting_down_Coordinate_Position['Right View'][1], scale)  # 左側試圖
draft.Right_View_Drawing('SN1-110', drafting_down_Coordinate_Position['Front View'][0], drafting_down_Coordinate_Position['Front View'][1], scale)  # 前視圖
draft.Right_Top_View_Drawing('SN1-110', drafting_up_Coordinate_Position['Top View'][0], drafting_up_Coordinate_Position['Top View'][1], scale)  # 上視圖


# 剖面圖座標
Section_Coordinate = {'A-A': [[0, -2900, 0, 600], 1],
                      'B-B': [[par.B[i] + 100, 520, par.B[i] + 100, -1200], 0],
                      'C-C': [[-195, -1460, 2286, par.B[i] + 100], 0],
                      'D-D': [[-par.A[i] / 2 - 100, -par.H[i] / 2, par.A[i] / 2 + 100, -par.H[i] / 2], 0],
                      'E-E': [[-100, -(par.S[i] + par.Z[i]) + 100, par.B[i] / 3, -(par.S[i] + par.Z[i]) + 100], 1],
                      'F-F': [[-272, - 176 / 2, -272, -176 * 1.5], 0],
                      'G-G': [[par.B[i] + 100, -(par.S[i] + par.Z[i]) + 150, par.B[i] + 100, -(par.S[i] + par.Z[i]) - 150], 1]}
draft.Section('Front view', drafting_down_Coordinate_Position['section A-A'][0],  drafting_down_Coordinate_Position['section A-A'][1], scale, Section_Coordinate['A-A'][0], 0)


