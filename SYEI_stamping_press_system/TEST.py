import main_program as mprog
import datetime, time, math
import drafting as draft
import parameter as par

i = 5

# mprog.switch_window()
#
# 隱藏機架外零件
for part_name in par.hide_part_name:
    mprog.hide_show_part(part_name, 1)

# 投影
mprog.switch_window()
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

# # 裁切選定圖面範圍
detail_view_coordinate = {
    'Section view D-D': [float(par.A[i] / 2), float(-270), float(-par.A[i] / 2), float(-270), float(-par.A[i] / 2),
                         float(-630), float(par.A[i] / 2), float(-630)],
    'Section view F-F': [float(-979 + 42), float(-25), float(-979 + 42), float(-290), float(-54 - 979), float(-290),
                         float(-979 - 54), float(-25)],
    'Section view G-G': [float(-par.A[i] / 2 - par.FRAME_6_7_width[i]), float(-par.S[i] - par.Z[i] + 150),
                         float(-par.A[i] / 2 - par.FRAME_6_7_width[i]),
                         float(-par.S[i] - par.Z[i] - 150), float(-par.A[i] / 2 + par.FRAME_6_7_width[i] + 170),
                         float(-par.S[i] - par.Z[i] - 150), float(-par.A[i] / 2 + par.FRAME_6_7_width[i] + 170),
                         float(-par.S[i] - par.Z[i] + 150)]}
detail_view_coordinate_keys = list(detail_view_coordinate.keys())
for coordinate_keys in detail_view_coordinate_keys:
    if coordinate_keys == 'Section view F-F' or coordinate_keys == 'Section view D-D':
        mprog.switch_window()
        if coordinate_keys == 'Section view F-F':
            for name in par.part_name_Section_view_F_F:
                mprog.hide_show_part(name, 1)
        elif coordinate_keys == 'Section view D-D':
            for name in par.part_name_Section_view_D_D:
                mprog.hide_show_part(name, 1)
        mprog.switch_window()
        if coordinate_keys in drafting_up_Coordinate_Position:
            draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                               drafting_up_Coordinate_Position[coordinate_keys][0],
                                               drafting_up_Coordinate_Position[coordinate_keys][1])
        else:
            draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                               drafting_down_Coordinate_Position[coordinate_keys][0],
                                               drafting_down_Coordinate_Position[coordinate_keys][1])
        mprog.switch_window()
        if coordinate_keys == 'Section view F-F':
            for name in par.part_name_Section_view_F_F:
                mprog.hide_show_part(name, 0)
        elif coordinate_keys == 'Section view D-D':
            for name in par.part_name_Section_view_D_D:
                mprog.hide_show_part(name, 0)
        mprog.switch_window()
    else:
        if coordinate_keys in drafting_up_Coordinate_Position:
            draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                               drafting_up_Coordinate_Position[coordinate_keys][0],
                                               drafting_up_Coordinate_Position[coordinate_keys][1])
        else:
            draft.Define_Polygonal_Detail_View(coordinate_keys, detail_view_coordinate[coordinate_keys],
                                               drafting_down_Coordinate_Position[coordinate_keys][0],
                                               drafting_down_Coordinate_Position[coordinate_keys][1])

# 產生折斷線
# 注意:折斷線若為水平線，方框四角座標須以⌈水平⌋建立；折斷線若為垂直線，方框四角座標須以⌈垂直⌋建立
break_line_Coordinate = {'Section view E-E': [
    [float(0), float(par.A[i] / 6), float(par.B[i] / 3), float(par.A[i] / 6), float(0), float(-par.A[i] / 6),
     float(par.B[i] / 3), float(-par.A[i] / 6)], 0, 1]}
print(break_line_Coordinate['Section view E-E'][0][0:])
draft.break_line('Section view E-E', break_line_Coordinate['Section view E-E'][0],
                 break_line_Coordinate['Section view E-E'][1], break_line_Coordinate['Section view E-E'][2])

# 關閉所有剖線線所產生之虛線
draft.close_all_Generated_Shape()

# 產生防漏試驗之虛線
leakproof_broken_line = {
    'Front view': [[float(-par.A[i] / 2 + par.FRAME_6_7_width[i] + 50), float(0),
                    float(-par.A[i] / 2 + 50 + par.FRAME_6_7_width[i]), float(-par.S[i] / 2 - 200)], 0,
                   'Section view H-H']}
draft.Section('Front view', drafting_down_Coordinate_Position['Section view A-A'][0],
              drafting_down_Coordinate_Position['Section view A-A'][1], scale,
              leakproof_broken_line['Front view'][0], leakproof_broken_line['Front view'][1],
              leakproof_broken_line['Front view'][2])
Detail_view_leak_broken_line = [float(550), float(-552), float(550), float(-583), float(978), float(-583), float(978),
                                float(-552)]
draft.Define_Polygonal_Cipping_View('Section view H-H', Detail_view_leak_broken_line)
draft.selection_Search_delete('Front view', "Name='Callout (Section View).3', all")

# 焊接符號生成
draft.symbol_of_weld('Front view', 3, -par.A[i] / 2 + 50 + 105, par.H[i] - par.S[i] - par.Z[i] - 32, 1, par.drafting_Welding_text['A-A Top'])

# 圈碼圖生成
draft.balloons('Front view', '1', -par.A[i] / 2 - 100 * par.cos30, (par.H[i] - par.S[i] - par.Z[i]) * par.sin30 / 2, -par.A[i] / 2, (par.H[i] - par.S[i] - par.Z[i]) / 2, -par.A[i] / 2 - 100 * par.cos30)
