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
# # mprog.switch_window()
drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale = draft.drafting_welding_view_parameter_calculation(
    par.A[i], par.B[i], par.H[i], par.S[i], par.Z[i], par.T[i])
draft.change_Drawing_scale(1 / scale)
# 副程式名稱須注意是否與立體圖相符
draft.Front_View_Drawing('SN1-110', drafting_down_Coordinate_Position['Right View'][0],
                         drafting_down_Coordinate_Position['Right View'][1], scale)  # 左側試圖
draft.Right_View_Drawing('SN1-110', drafting_down_Coordinate_Position['Front View'][0],
                         drafting_down_Coordinate_Position['Front View'][1], scale)  # 前視圖
draft.Right_Top_View_Drawing('SN1-110', drafting_up_Coordinate_Position['Top View'][0],
                             drafting_up_Coordinate_Position['Top View'][1], scale)  # 上視圖

# 剖面圖座標
Section_Coordinate = {'A-A': [[float(0), float(-2900), float(0), float(600)], 1],
                      'B-B': [[float(par.B[i] + 100), float(520), float(par.B[i] + 100), float(-1200)], 0],
                      'C-C': [[float(-195), float(-par.S[i] + 100), float(par.B[i] + 100), float(-par.S[i] + 100)], 0],
                      'D-D': [[float(-par.A[i] / 2 - 100), float(-par.H[i] / 2), float(par.A[i] / 2 + 100),
                               float(-par.H[i] / 2)], 1],
                      'E-E': [[float(-100), float(-(par.S[i] + par.Z[i]) + 100), float(par.B[i] / 3),
                               float(-(par.S[i] + par.Z[i]) + 100)], 0],
                      'F-F': [[float(272), float(- 176 / 2), float(272), float(-176 * 3 / 2)], 0],
                      'G-G': [[float(par.B[i] + 100), float(-(par.S[i] + par.Z[i]) + 200), float(par.B[i] + 100),
                               float(-(par.S[i] + par.Z[i]) - 200)], 0]}
Section_Coordinate_title = list(Section_Coordinate.keys())
view_order = ['Front view', 'Right view', 'Section view A-A', 'Front view', 'Right view', 'Section view B-B',
              'Section view A-A']
for title in Section_Coordinate_title:
    for view_name in view_order:
        if 'Section view ' + title in drafting_up_Coordinate_Position:  # 若此元素為此字典之鍵則執行
            draft.Section(view_name, drafting_up_Coordinate_Position['Section view ' + title][0],
                          drafting_up_Coordinate_Position['Section view ' + title][1], scale,
                          Section_Coordinate[title][0],
                          Section_Coordinate[title][1], 'Section view ' + title)
        else:
            draft.Section(view_name, drafting_down_Coordinate_Position['Section view ' + title][0],
                          drafting_down_Coordinate_Position['Section view ' + title][1], scale,
                          Section_Coordinate[title][0],
                          Section_Coordinate[title][1], 'Section view ' + title)
        view_order.pop(0)  # 刪除串列中第一個元素
        break

# 裁切選定圖面範圍
detail_view_coordinate = {
    'Section view D-D': [float(par.A[i] / 2), float(-270), float(-par.A[i] / 2), float(-270), float(-par.A[i] / 2),
                         float(-630), float(par.A[i] / 2), float(-630)],
    'Section view F-F': [float(-979 + 42), float(-25), float(-979 + 42), float(-290), float(-54 - 979), float(-290),
                         float(-979 - 54), float(-25)],
    'Section view G-G': [float(-par.A[i] / 2 - par.FRAME_6_7_width[i]), float(-par.S[i] - par.Z[i] + 150), float(-par.A[i] / 2 - par.FRAME_6_7_width[i]),
                         float(-par.S[i] - par.Z[i] - 150), float(-par.A[i] / 2 + par.FRAME_6_7_width[i] + 170),
                         float(-par.S[i] - par.Z[i] - 150), float(-par.A[i] / 2 + par.FRAME_6_7_width[i] + 170),
                         float(-par.S[i] - par.Z[i] + 150)]}
# detail_view_coordinate_keys = list(detail_view_coordinate.keys())
# for coordinate_keys in detail_view_coordinate_keys:
#     if coordinate_keys in drafting_up_Coordinate_Position:
#         draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
#                                            drafting_up_Coordinate_Position[coordinate_keys][0],
#                                            drafting_up_Coordinate_Position[coordinate_keys][1])
#     else:
#         draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
#                                            drafting_down_Coordinate_Position[coordinate_keys][0],
#                                            drafting_down_Coordinate_Position[coordinate_keys][1])

# # 焊接符號生成
# draft.symbol_of_weld('Front view', 3, -par.A[i] / 2 + 50 + 105, par.H[i] - par.S[i] - par.Z[i] - 32, 1)

# # 圈碼圖生成
# draft.balloons('Front view', '1', -par.A[i] / 2 - 100 * par.cos30, (par.H[i] - par.S[i] - par.Z[i]) * par.sin30 / 2, -par.A[i] / 2, (par.H[i] - par.S[i] - par.Z[i]) / 2, -par.A[i] / 2 - 100 * par.cos30)
