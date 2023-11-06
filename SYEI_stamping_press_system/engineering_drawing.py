import parameter as par
import main_program as mprog
import drafting as draft
import math


def welding_drawing(l, type, i, alpha):
    # --------焊接圖---------
    # 隱藏機架外零件
    for part_name in par.hide_part_name:
        mprog.hide_show_part(part_name, 1)

    # 投影
    mprog.switch_window()
    mprog.OPEN_Welding_diagram()
    if l == 1:
        drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale = draft.drafting_welding_view_parameter_calculation(
            par.A[i], par.B[i], par.H[i], par.S[i], par.Z[i], par.T[i])
    else:
        drafting_down_Coordinate_Position, drafting_up_Coordinate_Position, scale = draft.drafting_welding_view_parameter_calculation(
            par.A_15[i], par.B_15[i], par.H[i], par.S[i], par.Z[i], par.T[i])

    draft.change_Drawing_scale(1 / scale)
    # 副程式名稱須注意是否與立體圖相符
    draft.Front_View_Drawing(type, drafting_down_Coordinate_Position['Right View'][0],
                             drafting_down_Coordinate_Position['Right View'][1], scale)  # 左側試圖
    draft.Right_View_Drawing(type, drafting_down_Coordinate_Position['Front View'][0],
                             drafting_down_Coordinate_Position['Front View'][1], scale)  # 前視圖
    draft.Right_Top_View_Drawing(type, drafting_up_Coordinate_Position['Top View'][0],
                                 drafting_up_Coordinate_Position['Top View'][1], scale)  # 上視圖
    if l == 1:
        # 剖面圖座標
        Section_Coordinate = {'A-A': [
            [float(0), float(par.H[i] - par.S[i] - par.Z[i] + 150), float(0), float(-par.S[i] - par.Z[i] - 150)],
            0],
            'B-B': [[float(par.B[i] + 100), float(par.H[i] - par.S[i] - par.Z[i] + 150), float(par.B[i] + 100),
                     float(-par.S[i] / 2 - 150)], 0],
            'C-C': [
                [float(-195), float(-par.S[i] + 100), float(par.B[i] + 100), float(-par.S[i] + 100)],
                0],
            'D-D': [[float(-par.A[i] / 2 - 100), float(-par.H[i] / 2), float(par.A[i] / 2 + 100),
                     float(-par.H[i] / 2)], 1],
            'E-E': [[float(-100), float(-(par.S[i] + par.Z[i]) + 100), float(par.B[i] / 3),
                     float(-(par.S[i] + par.Z[i]) + 100)], 0],
            'F-F': [[float(272), float(- 176 / 2), float(272), float(-176 * 3 / 2)], 0],
            'G-G': [
                [float(par.B[i] + 100), float(-(par.S[i] + par.Z[i]) + 200), float(par.B[i] + 100),
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
            'Section view D-D': [float(par.A[i] / 2), float(-270), float(-par.A[i] / 2), float(-270),
                                 float(-par.A[i] / 2),
                                 float(-630), float(par.A[i] / 2), float(-630)],
            'Section view F-F': [float(-979 + 42), float(-25), float(-979 + 42), float(-290), float(-54 - 979),
                                 float(-290),
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
            'Front view': [[float(-par.R[i] / 2 - 20), float(0),
                            float(-par.R[i] / 2 - 20), float(-par.S[i] / 2 - 200)], 0,
                           'Section view H-H']}
        draft.Section('Front view', drafting_down_Coordinate_Position['Section view A-A'][0],
                      drafting_down_Coordinate_Position['Section view A-A'][1], scale,
                      leakproof_broken_line['Front view'][0], leakproof_broken_line['Front view'][1],
                      leakproof_broken_line['Front view'][2])
        Detail_view_leak_broken_line = [float(550), float(par.FRAME_10_11_center_to_Y_1[i]), float(550),
                                        float(par.FRAME_10_11_center_to_Y_2[i]), float(978),
                                        float(par.FRAME_10_11_center_to_Y_2[i]),
                                        float(978), float(par.FRAME_10_11_center_to_Y_1[i])]
        draft.Define_Polygonal_Cipping_View('Section view H-H', Detail_view_leak_broken_line)
        draft.selection_Search_delete('Front view', "Name='Callout (Section View).3', all")
    else:
        # 剖面圖座標
        Section_Coordinate = {'A-A': [
            [float(0), float(par.H[i] - par.S[i] - par.Z[i] + 150), float(0), float(-par.S[i] - par.Z[i] - 150)],
            0],
            'B-B': [[float(par.B_15[i] + 100), float(520), float(par.B_15[i] + 100), float(-1200)], 0],
            'C-C': [
                [float(-195), float(-par.S[i] + 100), float(par.B_15[i] + 100),
                 float(-par.S[i] + 100)],
                0],
            'D-D': [[float(-par.A_15[i] / 2 - 100), float(-par.H[i] / 2), float(par.A_15[i] / 2 + 100),
                     float(-par.H[i] / 2)], 1],
            'E-E': [[float(-100), float(-(par.S[i] + par.Z[i]) + 100), float(par.B_15[i] / 3),
                     float(-(par.S[i] + par.Z[i]) + 100)], 0],
            'F-F': [[float(272), float(- 176 / 2), float(272), float(-176 * 3 / 2)], 0],
            'G-G': [
                [float(par.B_15[i] + 100), float(-(par.S[i] + par.Z[i]) + 200),
                 float(par.B_15[i] + 100),
                 float(-(par.S[i] + par.Z[i]) - 200)], 0]}
        Section_Coordinate_title = list(Section_Coordinate.keys())
        view_order = ['Front view', 'Right view', 'Section view A-A', 'Front view', 'Right view',
                      'Section view B-B',
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
            'Section view D-D': [float(par.A_15[i] / 2), float(-270), float(-par.A_15[i] / 2), float(-270),
                                 float(-par.A_15[i] / 2),
                                 float(-630), float(par.A_15[i] / 2), float(-630)],
            'Section view F-F': [float(-979 + 42), float(-25), float(-979 + 42), float(-290), float(-54 - 979),
                                 float(-290),
                                 float(-979 - 54), float(-25)],
            'Section view G-G': [float(-par.A_15[i] / 2 - par.FRAME_6_7_width[i]), float(-par.S[i] - par.Z[i] + 150),
                                 float(-par.A_15[i] / 2 - par.FRAME_6_7_width[i]),
                                 float(-par.S[i] - par.Z[i] - 150),
                                 float(-par.A_15[i] / 2 + par.FRAME_6_7_width[i] + 170),
                                 float(-par.S[i] - par.Z[i] - 150),
                                 float(-par.A_15[i] / 2 + par.FRAME_6_7_width[i] + 170),
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
            [float(0), float(par.A_15[i] / 6), float(par.B_15[i] / 3), float(par.A_15[i] / 6), float(0),
             float(-par.A_15[i] / 6),
             float(par.B_15[i] / 3), float(-par.A_15[i] / 6)], 0, 1]}
        print(break_line_Coordinate['Section view E-E'][0][0:])
        draft.break_line('Section view E-E', break_line_Coordinate['Section view E-E'][0],
                         break_line_Coordinate['Section view E-E'][1], break_line_Coordinate['Section view E-E'][2])

        # 關閉所有剖線線所產生之虛線
        draft.close_all_Generated_Shape()

        # 產生防漏試驗之虛線
        leakproof_broken_line = {
            'Front view': [[float(-par.R_15[i] / 2 - 20), float(0),
                            float(-par.R_15[i] / 2 - 20), float(-par.S[i] / 2 - 200)], 0,
                           'Section view H-H']}
        draft.Section('Front view', drafting_down_Coordinate_Position['Section view A-A'][0],
                      drafting_down_Coordinate_Position['Section view A-A'][1], scale,
                      leakproof_broken_line['Front view'][0], leakproof_broken_line['Front view'][1],
                      leakproof_broken_line['Front view'][2])
        Detail_view_leak_broken_line = [float(550), float(par.FRAME_10_11_center_to_Y_1[i]), float(550),
                                        float(par.FRAME_10_11_center_to_Y_2[i]), float(978),
                                        float(par.FRAME_10_11_center_to_Y_2[i]),
                                        float(978), float(par.FRAME_10_11_center_to_Y_1[i])]
        draft.Define_Polygonal_Cipping_View('Section view H-H', Detail_view_leak_broken_line)
        draft.selection_Search_delete('Front view', "Name='Callout (Section View).3', all")

    # 焊接符號生成
    if l == 0:
        draft.symbol_of_weld('Front view', 3, -par.R_15[i] / 2 - 140, par.H[i] - par.S[i] - par.Z[i] - 32 + alpha, 1,
                             par.drafting_Welding_text['A-A Top'])
    else:
        draft.symbol_of_weld('Front view', 3, -par.R[i] / 2 - 140, par.H[i] - par.S[i] - par.Z[i] - 32 + alpha,
                             1, par.drafting_Welding_text['A-A Top'])

    # 圈碼圖生成
    if l == 1:
        draft.balloons('Front view', '1', -100 - (par.H[i] - par.S[i] - par.Z[i] - 32 - alpha) * par.cos30 / par.sin30,
                       (par.H[i] - par.S[i] - par.Z[i] - 32 - alpha) / 2, -par.R[i] / 2 - 140, 0)
    else:
        draft.balloons('Front view', '1',
                       -100 - (par.H[i] - par.S[i] - par.Z[i] - 32 - alpha) * par.cos30 / par.sin30,
                       (par.H[i] - par.S[i] - par.Z[i] - 32 - alpha) / 2, -par.R_15[i] / 2 - 140, 0)


def explosion_diagram(l, type, i):
    # 重新定義拘束尺寸
    if l == 1:
        BOLSTER1_Offset_value = 1750
        GIB_Offset_value = -3000
        Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
        CLOCK_Offset_value = 2850
        CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1350
        BOLSTER1_XY_Z = -par.T[i]
        BOLSTER1_XY_T = par.Z[i] - par.T[i] + BOLSTER1_XY_Z
        BOLSTER1_XY = BOLSTER1_XY_T + par.DH_S[i]
        BALANCER = -1700
    else:
        BOLSTER1_Offset_value = 2500
        GIB_Offset_value = -4500
        Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
        CLOCK_Offset_value = 3500
        CLOCK_SHAFT_Offset_value = CLOCK_Offset_value - 1750
        BOLSTER1_XY_Z = -par.T[i]
        BOLSTER1_XY_T = par.Z[i] - par.T[i] + BOLSTER1_XY_Z
        BOLSTER1_XY = BOLSTER1_XY_T + par.DH_S[i]
        BALANCER = -1700
    if l == 1:
        mprog.constaint_value_change(3, -par.B[i] - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(6, -par.B[i] - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(36, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
        mprog.constaint_value_change(39, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
        mprog.constaint_value_change(63, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(66, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
        mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
        mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
        mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
        mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
        mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
        mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
        mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
        mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
        mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
        mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
        mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
        mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
        mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
        mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
        mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(168, BOLSTER1_XY, 0)
    else:
        mprog.constaint_value_change(3, -par.B_15[i] - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(6, -par.B_15[i] - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(36, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
        mprog.constaint_value_change(39, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
        mprog.constaint_value_change(63, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(66, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
        mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
        mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
        mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
        mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
        mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
        mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
        mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
        mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
        mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
        mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
        mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
        mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
        mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
        mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
        mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(168, BOLSTER1_XY, 0)

    mprog.update()  # 更新
    mprog.OPEN_Drawing()
    if l == 1:
        drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale = draft.drafting_parameter_calculation(
            1, par.A[i], par.B[i], par.H[i], par.T[i])  # 計算爆炸圖比例及位置
    else:
        drafting_Coordinate_Position, drafting_isometric_Coordinate_Position, scale = draft.drafting_parameter_calculation(
            0, par.A_15[i], par.B_15[i], par.H[i], par.T[i])  # 計算爆炸圖比例及位置

    draft.change_Drawing_scale(1 / scale)  # 圖面比例
    draft.exploded_Drawing_1(type, drafting_isometric_Coordinate_Position['exploded_1'][0],
                             drafting_isometric_Coordinate_Position['exploded_1'][1], scale)  # 爆炸圖1
    mprog.switch_window()  # 開啟3D圖視窗
    # # 還原零件初始位置
    BOLSTER1_Offset_value = 80 - par.F[i] / 2
    GIB_Offset_value = 334.5 - 45
    CLOCK_Offset_value = 35
    CLOCK_SHAFT_Offset_value = 45
    Behide_GIB_Offset_value = GIB_Offset_value - 334.5 - 65 + 109.85
    BOLSTER1_XY_Z = - par.Z[i]
    BOLSTER1_XY_T = -par.T[i]
    BOLSTER1_XY = par.DH_S[i]
    BALANCER = -32
    if l == 1:
        mprog.constaint_value_change(3, -par.B[i] - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(6, -par.B[i] - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(36, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
        mprog.constaint_value_change(39, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
        mprog.constaint_value_change(63, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(66, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
        mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
        mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
        mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
        mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
        mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
        mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
        mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
        mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
        mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
        mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
        mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
        mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
        mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
        mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
        mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(168, BOLSTER1_XY, 0)
    else:
        mprog.constaint_value_change(3, -par.B_15[i] - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(6, -par.B_15[i] - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(36, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
        mprog.constaint_value_change(39, -BOLSTER1_Offset_value - par.F[i] / 2 + 80, 1)
        mprog.constaint_value_change(63, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(66, -par.F[i] / 2 - 80 - 37.5 - BOLSTER1_Offset_value, 0)
        mprog.constaint_value_change(60, -37.5 - BOLSTER1_Offset_value, 1)
        mprog.constaint_value_change(69, -37.5 - BOLSTER1_Offset_value, 0)
        mprog.constaint_value_change(75, GIB_Offset_value + 45, 0)
        mprog.constaint_value_change(78, GIB_Offset_value + 45, 0)
        mprog.constaint_value_change(84, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(87, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(81, Behide_GIB_Offset_value, 0)
        mprog.constaint_value_change(96, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(102, Behide_GIB_Offset_value, 1)
        mprog.constaint_value_change(99, Behide_GIB_Offset_value, 0)
        mprog.constaint_value_change(174, CLOCK_Offset_value, 0)
        mprog.constaint_value_change(188, CLOCK_SHAFT_Offset_value, 1)
        mprog.constaint_value_change(176, BALANCER, 1)  # 右氣壓缸
        mprog.constaint_value_change(180, BALANCER, 1)  # 左氣壓缸
        mprog.constaint_value_change(2, BOLSTER1_XY_Z, 1)
        mprog.constaint_value_change(5, BOLSTER1_XY_Z, 0)
        mprog.constaint_value_change(8, BOLSTER1_XY_Z, 1)
        mprog.constaint_value_change(11, BOLSTER1_XY_Z, 0)
        mprog.constaint_value_change(26, BOLSTER1_XY_T, 1)
        mprog.constaint_value_change(29, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(35, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(38, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(56, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(59, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(62, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(65, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(68, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(71, BOLSTER1_XY_T, 0)
        mprog.constaint_value_change(168, BOLSTER1_XY, 0)
    mprog.update()
    # --------------------爆炸圖右圖------------------
    if l == 1:
        Fixture_offset_value = 2850
    else:
        Fixture_offset_value = 3250
    if l == 1:
        CLUCTH_ASSEMBLY_offset_value = par.B[i] + 1650
    else:
        CLUCTH_ASSEMBLY_offset_value = par.B_15[i] + 2250
    JOINT_ALL_offset_value = CLUCTH_ASSEMBLY_offset_value
    if l == 1:
        MAIN_GEAR_offset_value = par.B[i] + 850
    else:
        MAIN_GEAR_offset_value = par.B_15[i] + 1500

    mprog.constaint_value_change(159, -312.5 - Fixture_offset_value, 0)  # 支架
    mprog.constaint_value_change(186, CLUCTH_ASSEMBLY_offset_value, 1)  # 離合器
    mprog.constaint_value_change(203, JOINT_ALL_offset_value, 1)  # JOINT_All.1
    mprog.constaint_value_change(195, MAIN_GEAR_offset_value, 1)  # 大齒輪MAIN_GEA1
    mprog.constaint_value_change(201, -375, 1)  # Joint1
    mprog.update()
    mprog.switch_window()
    draft.exploded_Drawing_2(type, drafting_isometric_Coordinate_Position['exploded_2'][0],
                             drafting_isometric_Coordinate_Position['exploded_2'][1], scale)  # 爆炸圖2
    # 復原位置
    mprog.switch_window()
    mprog.constaint_value_change(159, 312.5, 0)  # 支架
    mprog.constaint_value_change(186, 510, 1)  # 離合器
    mprog.constaint_value_change(203, 1010, 1)  # JOINT_All.1
    mprog.constaint_value_change(195, 592.5 + 150, 1)  # 大齒輪MAIN_GEA1
    mprog.constaint_value_change(201, -84, 1)  # Joint1
    mprog.update()
    mprog.switch_window()
    # # --------------------爆炸圖下(前、左、右)----------------
    draft.Front_View_Drawing(type, drafting_Coordinate_Position['Front View'][0],
                             drafting_Coordinate_Position['Front View'][1], scale)
    draft.Left_View_Drawing(type, drafting_Coordinate_Position['Left View'][0],
                            drafting_Coordinate_Position['Left View'][1], scale)
    draft.Right_View_Drawing(type, drafting_Coordinate_Position['Right View'][0],
                             drafting_Coordinate_Position['Right View'][1], scale)


# ------------標註------------
def balloons(i, l, h):
    # ---------圈碼圖----------
    if l == 1:
        BLOSTER1_center_Offset_Value = 1775
        GIB_Offset_value = -3000
        CLOCK_Offset_value = 2850
        BOLSTER1_Offset_value = 1750
        Fixture_offset_value = 2850
        CLUCTH_ASSEMBLY_offset_value = par.B[i] + 1650
    else:
        BLOSTER1_center_Offset_Value = 2625
        GIB_Offset_value = -4500
        CLOCK_Offset_value = 3500
        BOLSTER1_Offset_value = 2600
        Fixture_offset_value = 3250
        CLUCTH_ASSEMBLY_offset_value = par.B_15[i] + 2250
    BOLSTER1_XY_Z = -par.T[i]
    BOLSTER1_XY_T = par.Z[i] - par.T[i] + BOLSTER1_XY_Z
    JOINT_ALL_offset_value = CLUCTH_ASSEMBLY_offset_value
    CLOCK_Pointer = 3325
    FRAME_TOP_LEFT_X = (par.R[i] + 90 + 50) * par.cos45 + 50
    FRAME_TOP_LEFT_Y = FRAME_TOP_LEFT_X * par.sin30 / par.cos30
    # (將3D圖面尺寸轉換為2D尺寸, 將圖面x座標尺寸轉為R再將R轉為Y)
    # 引線點座標
    if l == 1:
        point_position = {'2': [-CLOCK_Offset_value * par.cos45,
                                -CLOCK_Offset_value * par.cos45 * par.sin30 / par.cos30],
                          '3': [(-1250) * par.cos45,
                                (-1250) * par.cos45 * par.sin30 / par.cos30],
                          '4': [(-BOLSTER1_Offset_value - par.R[i] / 2) * par.cos45,
                                (-BOLSTER1_Offset_value - par.R[i] / 2) * par.cos45 * par.sin30 / par.cos30],
                          '5': [(-BOLSTER1_Offset_value - par.R[i]) * par.cos45,
                                (-BOLSTER1_Offset_value - par.R[i]) * par.cos45 * par.sin30 / par.cos30 - par.S[i] +
                                par.T[i]],
                          '6': [(GIB_Offset_value - (par.R[i] + 145) / 2) * par.cos45,
                                (GIB_Offset_value - (par.R[i] + 145) / 2) * par.cos45 * par.sin30 / par.cos30]}
        # 圈圈座標`
        circle_position = {
            '2': ['2', point_position['2'][0] - FRAME_TOP_LEFT_X, point_position['2'][1] + FRAME_TOP_LEFT_Y],
            '3': ['3', point_position['3'][0] - FRAME_TOP_LEFT_X, point_position['3'][1] + FRAME_TOP_LEFT_Y],
            '4': ['4', point_position['4'][0] - FRAME_TOP_LEFT_X, point_position['4'][1] + FRAME_TOP_LEFT_Y],
            '5': ['5', point_position['5'][0] * 2 * 0.9 - FRAME_TOP_LEFT_X,
                  point_position['5'][0] * 2 * 0.9 * par.sin30 / par.cos30 + FRAME_TOP_LEFT_Y],
            '6': ['6', point_position['6'][0] - FRAME_TOP_LEFT_X, point_position['6'][1] + FRAME_TOP_LEFT_Y]}

        draft.balloons('Isometric view1', circle_position['6'][0], circle_position['6'][1], circle_position['6'][2],
                       point_position['6'][0], point_position['6'][1])
        draft.balloons('Isometric view1', circle_position['2'][0], circle_position['2'][1], circle_position['2'][2],
                       point_position['2'][0], point_position['2'][1])
        draft.balloons('Isometric view1', circle_position['4'][0], circle_position['4'][1], circle_position['4'][2],
                       point_position['4'][0], point_position['4'][1])
        draft.balloons('Isometric view1', circle_position['5'][0], circle_position['5'][1], circle_position['5'][2],
                       point_position['5'][0], point_position['5'][1])
        draft.balloons('Isometric view1', circle_position['3'][0], circle_position['3'][1], circle_position['3'][2],
                       point_position['3'][0], point_position['3'][1])

        # ------------中心線-------------
        draft.create_center_line('Isometric view1', 0, 0, -CLOCK_Pointer * par.cos45,
                                 -CLOCK_Pointer * par.cos45 * par.sin30 / par.cos30)
        draft.create_center_line('Isometric view1', -BLOSTER1_center_Offset_Value * par.cos45,
                                 -BLOSTER1_center_Offset_Value * par.cos45 * par.sin30 / par.cos30,
                                 -BLOSTER1_center_Offset_Value * par.cos45, -par.S[i] - par.Z[i] - par.T[i])
        draft.create_center_line('Isometric view2', -742.5 * par.cos45, -742.5 * par.cos45 * par.sin30 / par.cos30,
                                 -JOINT_ALL_offset_value * par.cos45,
                                 -JOINT_ALL_offset_value * par.cos45 * par.sin30 / par.cos30)
        draft.create_center_line('Isometric view2', -742.5 * par.cos45,
                                 -625,
                                 -3080,
                                 -(Fixture_offset_value + par.B[i]) * par.cos45 * par.sin30 / par.cos30 - 300)
    else:
        if h == 0:
            T_h = par.DH_S[i]
        elif h == 1:
            T_h = par.DH_H[i]
        else:
            T_h = par.DH_P[i]
        print(T_h)
        point_position = {'2': [-CLOCK_Offset_value * par.cos45,
                                -CLOCK_Offset_value * par.cos45 * par.sin30 / par.cos30],
                          '3': [(-1250) * par.cos45,
                                (-1250) * par.cos45 * par.sin30 / par.cos30],
                          '4': [(-BOLSTER1_Offset_value - par.P_15[i] / 3) * par.cos45,
                                (-BOLSTER1_Offset_value) * par.cos45 * par.sin30 / par.cos30 - par.S[i] + par.T[
                                    i] + T_h + 150],
                          '5': [(-BOLSTER1_Offset_value - par.E_15[i] / 2) * par.cos45,
                                (-BOLSTER1_Offset_value - par.E_15[i] / 2) * par.cos45 * par.sin30 / par.cos30 - par.S[
                                    i] + par.T[i]],
                          '6': [(GIB_Offset_value - (par.R_15[i] + 145) / 2) * par.cos45,
                                (GIB_Offset_value - (par.R_15[i] + 145) / 2) * par.cos45 * par.sin30 / par.cos30]}
        # 圈圈座標`
        circle_position = {
            '2': ['2', point_position['2'][0] - FRAME_TOP_LEFT_X, point_position['2'][1] + FRAME_TOP_LEFT_Y],
            '3': ['3', point_position['3'][0] - FRAME_TOP_LEFT_X, point_position['3'][1] + FRAME_TOP_LEFT_Y],
            '4': ['4', point_position['4'][0] - FRAME_TOP_LEFT_X * 2,
                  point_position['4'][1] - (par.P_15[i] / 3) * par.cos45 * par.sin30 / par.cos30 + par.S[i] - par.T[
                      i] - T_h - 150],
            '5': ['5', point_position['5'][0] * 2 * 0.825 - FRAME_TOP_LEFT_X,
                  point_position['5'][0] + par.T[i] / 2 + FRAME_TOP_LEFT_Y],
            '6': ['6', point_position['6'][0] - FRAME_TOP_LEFT_X, point_position['6'][1] + FRAME_TOP_LEFT_Y]}

        draft.balloons('Isometric view1', circle_position['6'][0], circle_position['6'][1], circle_position['6'][2],
                       point_position['6'][0], point_position['6'][1])
        draft.balloons('Isometric view1', circle_position['2'][0], circle_position['2'][1], circle_position['2'][2],
                       point_position['2'][0], point_position['2'][1])
        draft.balloons('Isometric view1', circle_position['4'][0], circle_position['4'][1], circle_position['4'][2],
                       point_position['4'][0], point_position['4'][1])
        draft.balloons('Isometric view1', circle_position['5'][0], circle_position['5'][1], circle_position['5'][2],
                       point_position['5'][0], point_position['5'][1])
        draft.balloons('Isometric view1', circle_position['3'][0], circle_position['3'][1], circle_position['3'][2],
                       point_position['3'][0], point_position['3'][1])
    Left_view_list = []

    for x in Left_view_list:
        Left_view_list.append(x)
    for x in Left_view_list:
        draft.add_dimension_to_view('Left view', str(x), Left_view_list[x][0], Left_view_list[x][1],
                                    Left_view_list[x][2],
                                    Left_view_list[x][3], Left_view_list[x][4], Left_view_list[x][5])
