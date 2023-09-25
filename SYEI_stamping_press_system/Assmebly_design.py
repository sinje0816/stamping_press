import main_program as mprog
import excel_parameter_change as epc
def assembly(i, apv, path, alpha, beta, zeta):
    #excel匯入
    excel = epc.ExcelOp('組立尺寸', 'Assembly_value')
    try:
        assmebly_par = excel.get_assmebly_sheet_par(i)
        print('Assembly_value Parameter change success')
    except BaseException:
        print('Assembly_value Parameter change error')
    #新增組立檔
    mprog.assembly_create()
    #將零件匯入組立檔
    part = epc.ExcelOp('尺寸整理表', '沖床機架零件清單')
    #讀取零件匯入數量
    part_name, part_quantity = part.get_assmebly_quantity(i)
    for name in range(len(part_name)):
        for x in range(int(part_quantity[name])):
            mprog.import_file_Part(path + '\\' + 'machining', part_name[name])
    #定義基準零件
    mprog.base_lock('FRAME1.1', 'FRAME1.1', 0)
    mprog.add_offset_assembly('FRAME2.1', 'FRAME1.1', 0, 'XY plane', 0, 4)
    mprog.add_offset_assembly('FRAME2.1', 'FRAME1.1', 0, 'YZ plane', 0, 5)
    mprog.add_offset_assembly('FRAME2.1', 'FRAME1.1', -(apv['FRAME22']['D']+(apv['FRAME1']['CC']/2)+(apv['FRAME2']['CC']/2)), 'ZX plane', 0, 6)
    mprog.add_offset_assembly('FRAME3.1', 'FRAME1.1', -(apv['FRAME1']['A'] - apv['FRAME3']['A'])-0.5*alpha-beta, 'XY plane', 0, 43)
    mprog.add_offset_assembly('FRAME3.1', 'FRAME1.1', -(apv['FRAME1']['F'] - assmebly_par['Ass_F'])-zeta, 'YZ plane', 0, 44)
    mprog.add_offset_assembly('FRAME3.1', 'FRAME1.1', -(apv['FRAME1']['CC'] / 2 + apv['FRAME3']['B']), 'ZX plane', 0,
                              45)
    mprog.add_offset_assembly('FRAME4.1', 'FRAME2.1', -(apv['FRAME2']['A'] - apv['FRAME4']['A'])-0.5*alpha-beta, 'XY plane', 0, 46)
    mprog.add_offset_assembly('FRAME4.1', 'FRAME2.1', -(apv['FRAME2']['F'] - assmebly_par['Ass_G'])-zeta, 'YZ plane', 0, 47)
    mprog.add_offset_assembly('FRAME4.1', 'FRAME2.1', apv['FRAME2']['CC'] / 2, 'ZX plane', 0, 48)
    mprog.add_offset_assembly('FRAME5.1', 'FRAME2.1',
                              -(apv['FRAME2']['A'] - apv['FRAME2']['bbbbb'] - assmebly_par['Ass_K'])-0.5*alpha-beta, 'XY plane', 0, 70)
    mprog.add_offset_assembly('FRAME5.1', 'FRAME2.1', apv['FRAME5']['A'] + assmebly_par['Ass_L'] + apv['FRAME32']['C'],
                              'YZ plane', 1, 71)
    mprog.add_offset_assembly('FRAME5.1', 'FRAME2.1', apv['FRAME2']['CC'] / 2, 'ZX plane', 0, 72)
    mprog.add_offset_assembly('FRAME7.1', 'FRAME2.1',
                              -((apv['FRAME2']['E'] - apv['FRAME2']['DD']) - apv['FRAME7']['A'] - apv['FRAME7']['Y']),
                              'XY plane', 0, 25)
    mprog.add_offset_assembly('FRAME7.1', 'FRAME2.1', -(apv['FRAME2']['B'] - assmebly_par['Ass_B'])-zeta, 'YZ plane', 0, 26)
    mprog.add_offset_assembly('FRAME7.1', 'FRAME2.1', assmebly_par['Ass_E'], 'ZX plane', 0, 27)
    mprog.add_offset_assembly('FRAME8.1', 'FRAME1.1', -assmebly_par['Ass_D'], 'XY plane', 0, 37)
    mprog.add_offset_assembly('FRAME8.1', 'FRAME1.1',
                              -(apv['FRAME1']['B'] - assmebly_par['Ass_B'] - apv['FRAME8']['B'])-zeta, 'YZ plane', 0, 38)
    mprog.add_offset_assembly('FRAME8.1', 'FRAME1.1', -assmebly_par['Ass_E'], 'ZX plane', 0, 39)
    mprog.add_offset_assembly('FRAME9.1', 'FRAME2.1', 0, 'XY plane', 0, 34)
    mprog.add_offset_assembly('FRAME9.1', 'FRAME2.1', 0, 'YZ plane', 0, 35)
    mprog.add_offset_assembly('FRAME9.1', 'FRAME2.1', apv['FRAME2']['CC'] / 2, 'ZX plane', 0, 36)
    mprog.add_offset_assembly('FRAME10.1', 'FRAME1.1', 0, 'XY plane', 1, 10)
    mprog.add_offset_assembly('FRAME10.1', 'FRAME1.1', -(apv['FRAME10']['B']), 'YZ plane', 0, 11)
    mprog.add_offset_assembly('FRAME10.1', 'FRAME1.1',
                              (apv['FRAME22']['D'] - assmebly_par['Ass_A']) / 2 + apv['FRAME1']['CC'] / 2, 'ZX plane',
                              1, 12)
    mprog.add_offset_assembly('FRAME10.2', 'FRAME2.1', apv['FRAME10']['K(2)'], 'XY plane', 0, 13)
    mprog.add_offset_assembly('FRAME10.2', 'FRAME2.1', -(apv['FRAME10']['B']), 'YZ plane', 0, 14)
    mprog.add_offset_assembly('FRAME10.2', 'FRAME2.1',
                              (apv['FRAME22']['D'] - assmebly_par['Ass_A']) / 2 + apv['FRAME2']['CC'] / 2, 'ZX plane',
                              0, 15)
    mprog.add_offset_assembly('FRAME12.1', 'FRAME3.1', 0, 'XY plane', 0, 136)
    mprog.add_offset_assembly('FRAME12.1', 'FRAME3.1', -(apv['FRAME3']['O']), 'YZ plane', 1, 137)
    mprog.add_offset_assembly('FRAME12.1', 'FRAME3.1', apv['FRAME19']['AH'], 'ZX plane', 0, 138)
    mprog.add_offset_assembly('FRAME13.1', 'FRAME2.1', -(
                apv['FRAME2']['A'] - apv['FRAME30']['A'] - (apv['FRAME17']['g3'] - apv['FRAME17']['g2']) -
                apv['FRAME13']['k'] - apv['FRAME13']['j'] - assmebly_par['Ass_V'] - assmebly_par['Ass_S'])-beta, 'XY plane',
                              0, 121)
    mprog.add_offset_assembly('FRAME13.1', 'FRAME2.1', apv['FRAME2']['F'] - apv['FRAME2']['G'] - apv['FRAME2']['FF']+zeta,
                              'YZ plane', 1, 122)
    mprog.add_offset_assembly('FRAME13.1', 'FRAME2.1', -apv['FRAME1']['CC'] / 2, 'ZX plane', 1, 123)
    mprog.add_offset_assembly('FRAME13.2', 'FRAME1.1', -(
                apv['FRAME2']['A'] - apv['FRAME30']['A'] - (apv['FRAME17']['g3'] - apv['FRAME17']['g2']) -
                apv['FRAME13']['k'] - apv['FRAME13']['j'] - assmebly_par['Ass_V'] - assmebly_par['Ass_S'])-beta, 'XY plane',
                              0, 124)
    mprog.add_offset_assembly('FRAME13.2', 'FRAME1.1', apv['FRAME2']['F'] - apv['FRAME2']['G'] - apv['FRAME2']['FF']+zeta,
                              'YZ plane', 1, 125)
    mprog.add_offset_assembly('FRAME13.2', 'FRAME1.1', apv['FRAME1']['CC'] / 2 + apv['FRAME13']['E'], 'ZX plane', 1,
                              126)
    mprog.add_offset_assembly('FRAME14.1', 'FRAME2.1', -(
                apv['FRAME2']['A'] - apv['FRAME30']['A'] - assmebly_par['Ass_S'] - apv['FRAME14']['g3'])-beta, 'XY plane', 0,
                              109)
    mprog.add_offset_assembly('FRAME14.1', 'FRAME2.1',
                              apv['FRAME2']['F'] - apv['FRAME2']['G'] - apv['FRAME14']['B'] - apv['FRAME2']['FF']+zeta,
                              'YZ plane', 1, 110)
    mprog.add_offset_assembly('FRAME14.1', 'FRAME2.1', apv['FRAME14']['A'] + apv['FRAME2']['CC'] / 2, 'ZX plane', 0,
                              111)
    mprog.add_offset_assembly('FRAME15.1', 'FRAME1.1', -(
                apv['FRAME1']['A'] - apv['FRAME30']['A'] - assmebly_par['Ass_S'] - apv['FRAME17']['g3'])-beta, 'XY plane', 0,
                              115)
    mprog.add_offset_assembly('FRAME15.1', 'FRAME1.1', apv['FRAME1']['F'] - apv['FRAME1']['G'] - assmebly_par['Ass_T']+zeta,
                              'YZ plane', 1, 116)
    mprog.add_offset_assembly('FRAME15.1', 'FRAME1.1',
                              apv['FRAME3']['B'] / 2 - assmebly_par['Ass_U'] + apv['FRAME1']['CC'] / 2 - apv['FRAME15'][
                                  'B'], 'ZX plane', 1, 117)
    mprog.add_offset_assembly('FRAME15.2', 'FRAME2.1', -(
                apv['FRAME2']['A'] - apv['FRAME30']['A'] - assmebly_par['Ass_S'] - apv['FRAME17']['g3'])-beta, 'XY plane', 0,
                              118)
    mprog.add_offset_assembly('FRAME15.2', 'FRAME2.1', apv['FRAME2']['F'] - apv['FRAME2']['G'] - assmebly_par['Ass_T']+zeta,
                              'YZ plane', 1, 119)
    mprog.add_offset_assembly('FRAME15.2', 'FRAME2.1',
                              -(apv['FRAME3']['B'] / 2 - assmebly_par['Ass_U'] + apv['FRAME1']['CC'] / 2), 'ZX plane',
                              1, 120)
    mprog.add_offset_assembly('FRAME17.1', 'FRAME1.1', -(
                apv['FRAME1']['A'] - apv['FRAME30']['A'] - assmebly_par['Ass_S'] - apv['FRAME17']['g3'])-beta, 'XY plane', 0,
                              112)
    mprog.add_offset_assembly('FRAME17.1', 'FRAME1.1',
                              apv['FRAME1']['F'] - apv['FRAME1']['G'] - apv['FRAME17']['B'] - apv['FRAME2']['FF']+zeta,
                              'YZ plane', 1, 113)
    mprog.add_offset_assembly('FRAME17.1', 'FRAME1.1', -(apv['FRAME2']['CC'] / 2), 'ZX plane', 0, 114)
    mprog.add_offset_assembly('FRAME18.1', 'FRAME1.1', -(
                apv['FRAME1']['A'] - apv['FRAME1']['bbbbb'] - assmebly_par['Ass_M'] - apv['FRAME18']['F'])-0.5*alpha-beta, 'XY plane',
                              0, 76)
    mprog.add_offset_assembly('FRAME18.1', 'FRAME1.1', -(apv['FRAME18']['A'] + assmebly_par['Ass_N']), 'YZ plane', 0,
                              77)
    mprog.add_offset_assembly('FRAME18.1', 'FRAME1.1', -(apv['FRAME1']['CC'] / 2), 'ZX plane', 0, 78)
    mprog.add_offset_assembly('FRAME19.1', 'FRAME1.1', -((apv['FRAME1']['E'] - apv['FRAME19']['D'])), 'XY plane', 0, 28)
    mprog.add_offset_assembly('FRAME19.1', 'FRAME1.1', -(apv['FRAME1']['B'] - assmebly_par['Ass_C']), 'YZ plane', 0, 29)
    mprog.add_offset_assembly('FRAME19.1', 'FRAME1.1', -(apv['FRAME1']['CC'] / 2 + apv['FRAME19']['AH']), 'ZX plane', 0,
                              30)
    mprog.add_offset_assembly('FRAME19.2', 'FRAME2.1', -((apv['FRAME1']['E'] - apv['FRAME19']['D'])), 'XY plane', 0, 31)
    mprog.add_offset_assembly('FRAME19.2', 'FRAME2.1', -(apv['FRAME2']['B'] - assmebly_par['Ass_C']), 'YZ plane', 0, 32)
    mprog.add_offset_assembly('FRAME19.2', 'FRAME2.1', apv['FRAME2']['CC'] / 2, 'ZX plane', 0, 33)
    mprog.add_offset_assembly('FRAME20.1', 'FRAME2.1', -(
                apv['FRAME1']['A'] - apv['FRAME1']['bbbbb'] - assmebly_par['Ass_R'] - apv['FRAME20']['H'])-0.5*alpha-beta, 'XY plane',
                              0, 100)
    mprog.add_offset_assembly('FRAME20.1', 'FRAME2.1', -(apv['FRAME20']['A'] + assmebly_par['Ass_N']), 'YZ plane', 0,
                              101)
    mprog.add_offset_assembly('FRAME20.1', 'FRAME2.1', apv['FRAME20']['B'] + apv['FRAME1']['CC'] / 2, 'ZX plane', 0,
                              102)
    mprog.add_offset_assembly('FRAME22.1', 'FRAME35.1', -(
                (apv['FRAME1']['E'] - apv['FRAME1']['DD']) - apv['FRAME22']['A(1)'] - apv['FRAME22']['A(2)']),
                              'XY plane', 1, 16)
    mprog.add_offset_assembly('FRAME22.1', 'FRAME35.1', apv['FRAME22']['O'], 'YZ plane', 0, 17)
    mprog.add_offset_assembly('FRAME22.1', 'FRAME1.1', -(apv['FRAME22']['D'] + apv['FRAME1']['CC'] / 2), 'ZX plane', 0,
                              18)
    mprog.add_offset_assembly('FRAME24.1', 'FRAME30.1', -(apv['FRAME30']['A'] - apv['FRAME24']['B']), 'XY plane', 0, 58)
    mprog.add_offset_assembly('FRAME24.1', 'FRAME30.1', apv['FRAME24']['A'] + apv['FRAME30']['M'] / 2, 'YZ plane', 0,
                              59)
    mprog.add_offset_assembly('FRAME24.1', 'FRAME30.1',
                              apv['FRAME30']['E'] - (apv['FRAME24']['I'] - apv['FRAME2']['CC']), 'ZX plane', 0, 60)
    mprog.add_offset_assembly('FRAME24_1.1', 'FRAME30.1', -(apv['FRAME30']['A'] - apv['FRAME24']['B']), 'XY plane', 0,
                              61)
    mprog.add_offset_assembly('FRAME24_1.1', 'FRAME30.1', apv['FRAME24']['A'] + apv['FRAME30']['M'] / 2, 'YZ plane', 0,
                              62)
    mprog.add_offset_assembly('FRAME24_1.1', 'FRAME30.1', (apv['FRAME24']['I'] - apv['FRAME2']['CC']), 'ZX plane', 0,
                              63)
    mprog.add_offset_assembly('FRAME26.1', 'FRAME18.1', apv['FRAME26']['B'], 'XY plane', 0, 79)
    mprog.add_offset_assembly('FRAME26.1', 'FRAME18.1', -assmebly_par['Ass_O'], 'YZ plane', 1, 80)
    mprog.add_offset_assembly('FRAME26.1', 'FRAME18.1', 0, 'ZX plane', 1, 81)
    mprog.add_offset_assembly('FRAME26.2', 'FRAME18.1', apv['FRAME26']['B'], 'XY plane', 0, 82)
    mprog.add_offset_assembly('FRAME26.2', 'FRAME18.1',
                              -(apv['FRAME18']['A'] - (assmebly_par['Ass_O'] + apv['FRAME26']['I'])), 'YZ plane', 1, 83)
    mprog.add_offset_assembly('FRAME26.2', 'FRAME18.1', 0, 'ZX plane', 1, 84)
    mprog.add_offset_assembly('FRAME26.3', 'FRAME20.1', apv['FRAME26']['B'], 'XY plane', 0, 103)
    mprog.add_offset_assembly('FRAME26.3', 'FRAME20.1', assmebly_par['Ass_O'] + apv['FRAME26']['I'], 'YZ plane', 0, 104)
    mprog.add_offset_assembly('FRAME26.3', 'FRAME20.1', -(apv['FRAME20']['B']), 'ZX plane', 0, 105)
    mprog.add_offset_assembly('FRAME26.4', 'FRAME20.1', apv['FRAME26']['B'], 'XY plane', 0, 106)
    mprog.add_offset_assembly('FRAME26.4', 'FRAME20.1', apv['FRAME20']['A'] - assmebly_par['Ass_O'], 'YZ plane', 0, 107)
    mprog.add_offset_assembly('FRAME26.4', 'FRAME20.1', -(apv['FRAME20']['B']), 'ZX plane', 0, 108)
    mprog.add_offset_assembly('FRAME28.1', 'FRAME30.1', 0, 'XY plane', 0, 145)
    mprog.add_offset_assembly('FRAME28.1', 'FRAME30.1', apv['FRAME30']['M'] / 2 + apv['FRAME28']['A'], 'YZ plane', 0,
                              146)
    mprog.add_offset_assembly('FRAME28.1', 'FRAME30.1', apv['FRAME27']['I'] + 1, 'ZX plane', 0, 147)
    mprog.add_offset_assembly('FRAME29.1', 'FRAME4.1', -(
                assmebly_par['Ass_Q'] + (apv['FRAME4']['A'] - apv['FRAME4']['a']) - apv['FRAME29']['A'] / 2),
                              'XY plane', 0, 94)
    mprog.add_offset_assembly('FRAME29.1', 'FRAME4.1', apv['FRAME29']['B'], 'YZ plane', 0, 95)
    mprog.add_offset_assembly('FRAME29.1', 'FRAME4.1',
                              apv['FRAME4']['B'] / 2 - apv['FRAME6']['e'] + apv['FRAME29']['A'] / 2, 'ZX plane', 0, 96)
    mprog.add_offset_assembly('FRAME29.2', 'FRAME4.1', -(
                assmebly_par['Ass_Q'] + (apv['FRAME4']['A'] - apv['FRAME4']['a']) - apv['FRAME29']['A'] / 2),
                              'XY plane', 0, 97)
    mprog.add_offset_assembly('FRAME29.2', 'FRAME4.1', apv['FRAME29']['B'], 'YZ plane', 0, 98)
    mprog.add_offset_assembly('FRAME29.2', 'FRAME4.1',
                              apv['FRAME4']['B'] / 2 + apv['FRAME6']['e'] + apv['FRAME29']['A'] / 2, 'ZX plane', 0, 99)
    mprog.add_offset_assembly('FRAME30.1', 'FRAME2.1', -(apv['FRAME2']['A'] - apv['FRAME30']['A'])-beta, 'XY plane', 0, 40)
    mprog.add_offset_assembly('FRAME30.1', 'FRAME2.1', -(apv['FRAME2']['F'] - (apv['FRAME30']['M'] / 2))-zeta, 'YZ plane', 0,
                              41)
    mprog.add_offset_assembly('FRAME30.1', 'FRAME2.1', apv['FRAME2']['CC'] / 2, 'ZX plane', 0, 42)
    mprog.add_offset_assembly('FRAME32.1', 'FRAME2.1', -(
                apv['FRAME2']['A'] - apv['FRAME2']['bbbbb'] - assmebly_par['Ass_K'] - apv['FRAME32']['A'])-0.5*alpha-beta, 'XY plane',
                              0, 73)
    mprog.add_offset_assembly('FRAME32.1', 'FRAME2.1', -(assmebly_par['Ass_L']), 'YZ plane', 0, 74)
    mprog.add_offset_assembly('FRAME32.1', 'FRAME2.1', apv['FRAME2']['CC'] / 2, 'ZX plane', 0, 75)
    mprog.add_offset_assembly('FRAME35.1', 'FRAME1.1', 0, 'XY plane', 1, 1)
    mprog.add_offset_assembly('FRAME35.1', 'FRAME1.1', -(apv['FRAME1']['B'])-zeta, 'YZ plane', 0, 2)
    mprog.add_offset_assembly('FRAME35.1', 'FRAME1.1',
                              (apv['FRAME22']['D'] - assmebly_par['Ass_A']) / 2 + (apv['FRAME1']['CC'] / 2), 'ZX plane',
                              1, 3)
    mprog.add_offset_assembly('FRAME35.2', 'FRAME2.1', apv['FRAME35']['K'] - apv['FRAME35']['L'], 'XY plane', 0, 7)
    mprog.add_offset_assembly('FRAME35.2', 'FRAME2.1', -(apv['FRAME2']['B'])-zeta, 'YZ plane', 0, 8)
    mprog.add_offset_assembly('FRAME35.2', 'FRAME2.1',
                              (apv['FRAME22']['D'] - assmebly_par['Ass_A']) / 2 + apv['FRAME2']['CC'] / 2, 'ZX plane',
                              0, 9)
    mprog.add_offset_assembly('FRAME38.1', 'FRAME19.2',
                              -(apv['FRAME19']['A'] - assmebly_par['Ass_W'] - apv['FRAME38']['A'] - apv['FRAME3']['C']),
                              'XY plane', 0, 133)
    mprog.add_offset_assembly('FRAME38.1', 'FRAME19.2', assmebly_par['Ass_X'], 'YZ plane', 1, 134)
    mprog.add_offset_assembly('FRAME38.1', 'FRAME19.2', -(apv['FRAME19']['AH'] + apv['FRAME38']['B']), 'ZX plane', 1,
                              135)
    mprog.add_offset_assembly('FRAME38.2', 'FRAME19.1',
                              -(apv['FRAME19']['A'] - assmebly_par['Ass_W'] - apv['FRAME38']['A'] - apv['FRAME3']['C']),
                              'XY plane', 0, 139)
    mprog.add_offset_assembly('FRAME38.2', 'FRAME19.1', -(assmebly_par['Ass_X'] + apv['FRAME38']['C']), 'YZ plane', 0,
                              140)
    mprog.add_offset_assembly('FRAME38.2', 'FRAME19.1', -(apv['FRAME38']['B']), 'ZX plane', 0, 141)

    if i == 0:
        mprog.add_offset_assembly('FRAME33.1', 'FRAME1.1', 0, 'XY plane', 0, 19)
        mprog.add_offset_assembly('FRAME33.1', 'FRAME1.1', -(apv['FRAME1']['B']-apv['FRAME33']['B'])-zeta, 'YZ plane', 0, 20)
        mprog.add_offset_assembly('FRAME33.1', 'FRAME1.1', -(apv['FRAME33']['K']+apv['FRAME1']['CC']/2), 'ZX plane', 0, 21)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME2.1', 0, 'XY plane', 0, 22)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME2.1', -(apv['FRAME2']['B']-apv['FRAME34']['B'])-zeta, 'YZ plane', 0, 23)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME2.1', apv['FRAME2']['CC']/2, 'ZX plane', 0, 24)
        mprog.add_offset_assembly('FRAME31.1', 'FRAME4.1', -(apv['FRAME4']['A']-apv['FRAME4']['a']-apv['FRAME4']['b']/2), 'XY plane', 0, 49)
        mprog.add_offset_assembly('FRAME31.1', 'FRAME4.1', -(apv['FRAME4']['C']-assmebly_par['Ass_H']), 'YZ plane', 0, 50)
        mprog.add_offset_assembly('FRAME31.1', 'FRAME4.1', apv['FRAME4']['B']/2-apv['FRAME4']['b']/2, 'ZX plane', 0, 51)
        mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', -(apv['FRAME3']['A']-apv['FRAME3']['E']-apv['FRAME11']['A']/2), 'XY plane', 0, 64)
        mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', apv['FRAME3']['O']+assmebly_par['Ass_I'], 'YZ plane', 0, 65)
        mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', apv['FRAME3']['B']/2+apv['FRAME11']['A']/2, 'ZX plane', 0, 66)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME4.1', -((assmebly_par['Ass_J']-apv['FRAME36']['A']/2)+(apv['FRAME4']['A']-apv['FRAME4']['a'])), 'XY plane', 0, 67)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME4.1', -(apv['FRAME4']['K']), 'YZ plane', 0, 68)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME4.1', apv['FRAME4']['B']-apv['FRAME36']['D'], 'ZX plane', 0, 69)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME4.1', -((assmebly_par['Ass_J']-apv['FRAME25']['B'])+(apv['FRAME4']['A']-apv['FRAME4']['a'])), 'XY plane', 0, 85)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME4.1', -(apv['FRAME4']['K']), 'YZ plane', 0, 86)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME4.1', assmebly_par['Ass_P']+apv['FRAME4']['B']/2, 'ZX plane', 0, 87)
        mprog.add_offset_assembly('FRAME27.1', 'FRAME1.1', -(apv['FRAME1']['A'] - apv['FRAME27']['A']) - beta,
                                  'XY plane', 0,
                                  52)
        mprog.add_offset_assembly('FRAME27.1', 'FRAME1.1', -(
                apv['FRAME1']['F'] - apv['FRAME27']['B'] - apv['FRAME30']['M'] - apv['FRAME27']['J']) - zeta,
                                  'YZ plane',
                                  0, 53)
        mprog.add_offset_assembly('FRAME27.1', 'FRAME1.1', -(apv['FRAME1']['CC'] / 2), 'ZX plane', 0, 54)
        mprog.add_offset_assembly('FRAME27_1.1', 'FRAME2.1', -(apv['FRAME2']['A'] - apv['FRAME27_1']['A']) - beta,
                                  'XY plane',
                                  0, 55)
        mprog.add_offset_assembly('FRAME27_1.1', 'FRAME2.1', -(
                apv['FRAME2']['F'] - apv['FRAME27_1']['B'] - apv['FRAME30']['M'] - apv['FRAME27_1']['J']) - zeta,
                                  'YZ plane', 0, 56)
        mprog.add_offset_assembly('FRAME27_1.1', 'FRAME2.1', apv['FRAME27_1']['I'] + apv['FRAME2']['CC'] / 2,
                                  'ZX plane', 0, 57)
        mprog.add_offset_assembly('FRAME6.1', 'FRAME4.1', assmebly_par['Ass_Q']-(apv['FRAME4']['A']-apv['FRAME4']['a'])-apv['FRAME6']['E']/2, 'XY plane', 0, 88)
        mprog.add_offset_assembly('FRAME6.1', 'FRAME4.1', 0, 'YZ plane', 1, 89)
        mprog.add_offset_assembly('FRAME6.1', 'FRAME4.1', -(apv['FRAME4']['B']/2-apv['FRAME6']['f']), 'ZX plane', 1, 90)
        mprog.add_offset_assembly('FRAME6.2', 'FRAME4.1', assmebly_par['Ass_Q']-(apv['FRAME4']['A']-apv['FRAME4']['a'])-apv['FRAME6']['E']/2, 'XY plane', 0, 91)
        mprog.add_offset_assembly('FRAME6.2', 'FRAME4.1', apv['FRAME6']['i'], 'YZ plane', 0, 92)
        mprog.add_offset_assembly('FRAME6.2', 'FRAME4.1', apv['FRAME4']['B']/2+apv['FRAME6']['f'], 'ZX plane', 0, 93)
        mprog.add_offset_assembly('FRAME21.1', 'FRAME1.1',apv['FRAME2']['A']-apv['FRAME30']['A']-(apv['FRAME17']['g3']-apv['FRAME17']['g2'])-apv['FRAME13']['k']-apv['FRAME13']['j']-assmebly_par['Ass_V']-assmebly_par['Ass_S']+beta, 'XY plane', 1, 127)
        mprog.add_offset_assembly('FRAME21.1', 'FRAME1.1',-(apv['FRAME2']['F']-apv['FRAME2']['G']-assmebly_par['Ass_T']-apv['FRAME21']['D'])-zeta, 'YZ plane', 0, 128)
        mprog.add_offset_assembly('FRAME21.1', 'FRAME1.1',-(apv['FRAME3']['B']/2-assmebly_par['Ass_U']+apv['FRAME1']['CC']/2-apv['FRAME21']['C']), 'ZX plane', 0, 129)
        mprog.add_offset_assembly('FRAME21.2', 'FRAME2.1',apv['FRAME2']['A']-apv['FRAME30']['A']-(apv['FRAME17']['g3']-apv['FRAME17']['g2'])-apv['FRAME13']['k']-apv['FRAME13']['j']-assmebly_par['Ass_V']-assmebly_par['Ass_S']+beta, 'XY plane', 1, 130)
        mprog.add_offset_assembly('FRAME21.2', 'FRAME2.1',-(apv['FRAME2']['F']-apv['FRAME2']['G']-assmebly_par['Ass_T']-apv['FRAME21']['D'])-zeta, 'YZ plane', 0, 131)
        mprog.add_offset_assembly('FRAME21.2', 'FRAME2.1',(apv['FRAME3']['B']/2-assmebly_par['Ass_U']+apv['FRAME1']['CC']/2), 'ZX plane', 0, 132)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME2.1', -(apv['FRAME2']['A']-apv['FRAME2']['ffff']-apv['FRAME37']['A']/2), 'XY plane', 0, 142)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME2.1', -(apv['FRAME2']['F']-apv['FRAME2']['eeee']-apv['FRAME37']['A']/2)-zeta, 'YZ plane', 0, 143)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME2.1', apv['FRAME2']['CC']/2, 'ZX plane', 0, 144)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME1.1',
                                  -(apv["FRAME1"]["E"] - apv["FRAME1"]["DD"] - apv["FRAME23"]["I"]), 'XY plane', 0, 145)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME1.1', -(apv['FRAME1']['B'] - apv['FRAME23']['B']) - zeta,
                                  'YZ plane', 0, 146)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME1.1', -(apv['FRAME1']['CC'] / 2 + apv['FRAME23']['A']), 'ZX plane',
                                  1, 147)
        mprog.add_offset_assembly('FRAME23.2', 'FRAME1.1',
                                  -(apv["FRAME1"]["E"] - apv["FRAME1"]["DD"] - apv["FRAME23"]["I"]), 'XY plane', 0, 148)
        mprog.add_offset_assembly('FRAME23.2', 'FRAME1.1', -(apv["FRAME1"]["B"]-apv["FRAME23"]["B"] - assmebly_par['Ass_Y'])-zeta,
                                  'YZ plane', 0, 149)
        mprog.add_offset_assembly('FRAME23.2', 'FRAME1.1', -(apv['FRAME1']['CC'] / 2 + apv['FRAME23']['A']), 'ZX plane',
                                  1, 150)
        mprog.add_offset_assembly('FRAME23.3', 'FRAME2.1',
                                  -(apv["FRAME1"]["E"] - apv["FRAME1"]["DD"] - apv["FRAME23"]["I"]), 'XY plane', 0, 151)
        mprog.add_offset_assembly('FRAME23.3', 'FRAME2.1', (apv['FRAME1']['B']) + zeta,
                                  'YZ plane', 1, 152)
        mprog.add_offset_assembly('FRAME23.3', 'FRAME2.1', -(apv['FRAME1']['CC'] / 2 + apv['FRAME23']['A']), 'ZX plane',
                                  0, 153)
        mprog.add_offset_assembly('FRAME23.4', 'FRAME2.1',
                                  -(apv["FRAME1"]["E"] - apv["FRAME1"]["DD"] - apv["FRAME23"]["I"]), 'XY plane', 0, 154)
        mprog.add_offset_assembly('FRAME23.4', 'FRAME2.1', (apv["FRAME1"]["B"] - assmebly_par['Ass_Y'])+zeta,
                                  'YZ plane', 1, 155)
        mprog.add_offset_assembly('FRAME23.4', 'FRAME2.1', -(apv['FRAME2']['CC'] / 2 + apv['FRAME23']['A']), 'ZX plane',
                                  0, 156)
    elif i == 1:
        mprog.add_offset_assembly('FRAME33.1', 'FRAME1.1', 0, 'XY plane', 0, 19)
        mprog.add_offset_assembly('FRAME33.1', 'FRAME1.1', -(apv['FRAME1']['B']-apv['FRAME33']['B'])-zeta, 'YZ plane', 0, 20)
        mprog.add_offset_assembly('FRAME33.1', 'FRAME1.1', -(apv['FRAME33']['K']+apv['FRAME1']['CC']/2), 'ZX plane', 0, 21)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME2.1', 0, 'XY plane', 0, 22)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME2.1', -(apv['FRAME2']['B']-apv['FRAME34']['B'])-zeta, 'YZ plane', 0, 23)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME2.1', apv['FRAME2']['CC']/2, 'ZX plane', 0, 24)
        mprog.add_offset_assembly('FRAME31.1', 'FRAME4.1', -(apv['FRAME4']['A']-apv['FRAME4']['a']-apv['FRAME4']['b']/2), 'XY plane', 0, 49)
        mprog.add_offset_assembly('FRAME31.1', 'FRAME4.1', -(apv['FRAME4']['C']-assmebly_par['Ass_H']), 'YZ plane', 0, 50)
        mprog.add_offset_assembly('FRAME31.1', 'FRAME4.1', apv['FRAME4']['B']/2-apv['FRAME4']['b']/2, 'ZX plane', 0, 51)
        mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', -(apv['FRAME3']['A']-apv['FRAME3']['E']-apv['FRAME11']['A']/2), 'XY plane', 0, 64)
        mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', apv['FRAME3']['O']+assmebly_par['Ass_I'], 'YZ plane', 0, 65)
        mprog.add_offset_assembly('FRAME11.1', 'FRAME3.1', apv['FRAME3']['B']/2+apv['FRAME11']['A']/2, 'ZX plane', 0, 66)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME4.1', -((assmebly_par['Ass_J']-apv['FRAME36']['A']/2)+(apv['FRAME4']['A']-apv['FRAME4']['a'])), 'XY plane', 0, 67)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME4.1', -(apv['FRAME4']['K']), 'YZ plane', 0, 68)
        mprog.add_offset_assembly('FRAME36.1', 'FRAME4.1', apv['FRAME4']['B']-apv['FRAME36']['D'], 'ZX plane', 0, 69)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME4.1', -((assmebly_par['Ass_J']-apv['FRAME25']['B'])+(apv['FRAME4']['A']-apv['FRAME4']['a'])), 'XY plane', 0, 85)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME4.1', -(apv['FRAME4']['K']), 'YZ plane', 0, 86)
        mprog.add_offset_assembly('FRAME25.1', 'FRAME4.1', assmebly_par['Ass_P']+apv['FRAME4']['B']/2, 'ZX plane', 0, 87)
        mprog.add_offset_assembly('FRAME27.1', 'FRAME1.1', -(apv['FRAME1']['A'] - apv['FRAME27']['A']) - beta,
                                  'XY plane', 0,
                                  52)
        mprog.add_offset_assembly('FRAME27.1', 'FRAME1.1', -(
                apv['FRAME1']['F'] - apv['FRAME27']['B'] - apv['FRAME30']['M'] - apv['FRAME27']['J']) - zeta,
                                  'YZ plane',
                                  0, 53)
        mprog.add_offset_assembly('FRAME27.1', 'FRAME1.1', -(apv['FRAME1']['CC'] / 2), 'ZX plane', 0, 54)
        mprog.add_offset_assembly('FRAME27_1.1', 'FRAME2.1', -(apv['FRAME2']['A'] - apv['FRAME27_1']['A']) - beta,
                                  'XY plane',
                                  0, 55)
        mprog.add_offset_assembly('FRAME27_1.1', 'FRAME2.1', -(
                apv['FRAME2']['F'] - apv['FRAME27_1']['B'] - apv['FRAME30']['M'] - apv['FRAME27_1']['J']) - zeta,
                                  'YZ plane', 0, 56)
        mprog.add_offset_assembly('FRAME27_1.1', 'FRAME2.1', apv['FRAME27_1']['I'] + apv['FRAME2']['CC'] / 2,
                                  'ZX plane', 0, 57)
        mprog.add_offset_assembly('FRAME6.1', 'FRAME4.1', assmebly_par['Ass_Q']-(apv['FRAME4']['A']-apv['FRAME4']['a'])-apv['FRAME6']['E']/2, 'XY plane', 0, 88)
        mprog.add_offset_assembly('FRAME6.1', 'FRAME4.1', 0, 'YZ plane', 1, 89)
        mprog.add_offset_assembly('FRAME6.1', 'FRAME4.1', -(apv['FRAME4']['B']/2-apv['FRAME6']['f']), 'ZX plane', 1, 90)
        mprog.add_offset_assembly('FRAME6.2', 'FRAME4.1', assmebly_par['Ass_Q']-(apv['FRAME4']['A']-apv['FRAME4']['a'])-apv['FRAME6']['E']/2, 'XY plane', 0, 91)
        mprog.add_offset_assembly('FRAME6.2', 'FRAME4.1', apv['FRAME6']['i'], 'YZ plane', 0, 92)
        mprog.add_offset_assembly('FRAME6.2', 'FRAME4.1', apv['FRAME4']['B']/2+apv['FRAME6']['f'], 'ZX plane', 0, 93)
        mprog.add_offset_assembly('FRAME21.1', 'FRAME1.1',apv['FRAME2']['A']-apv['FRAME30']['A']-(apv['FRAME17']['g3']-apv['FRAME17']['g2'])-apv['FRAME13']['k']-apv['FRAME13']['j']-assmebly_par['Ass_V']-assmebly_par['Ass_S']+beta, 'XY plane', 1, 127)
        mprog.add_offset_assembly('FRAME21.1', 'FRAME1.1',-(apv['FRAME2']['F']-apv['FRAME2']['G']-assmebly_par['Ass_T']-apv['FRAME21']['D'])-zeta, 'YZ plane', 0, 128)
        mprog.add_offset_assembly('FRAME21.1', 'FRAME1.1',-(apv['FRAME3']['B']/2-assmebly_par['Ass_U']+apv['FRAME1']['CC']/2-apv['FRAME21']['C']), 'ZX plane', 0, 129)
        mprog.add_offset_assembly('FRAME21.2', 'FRAME2.1',apv['FRAME2']['A']-apv['FRAME30']['A']-(apv['FRAME17']['g3']-apv['FRAME17']['g2'])-apv['FRAME13']['k']-apv['FRAME13']['j']-assmebly_par['Ass_V']-assmebly_par['Ass_S']+beta, 'XY plane', 1, 130)
        mprog.add_offset_assembly('FRAME21.2', 'FRAME2.1',-(apv['FRAME2']['F']-apv['FRAME2']['G']-assmebly_par['Ass_T']-apv['FRAME21']['D'])-zeta, 'YZ plane', 0, 131)
        mprog.add_offset_assembly('FRAME21.2', 'FRAME2.1',(apv['FRAME3']['B']/2-assmebly_par['Ass_U']+apv['FRAME1']['CC']/2), 'ZX plane', 0, 132)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME2.1', -(apv['FRAME2']['A']-apv['FRAME2']['ffff']-apv['FRAME37']['A']/2), 'XY plane', 0, 142)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME2.1', -(apv['FRAME2']['F']-apv['FRAME2']['eeee']-apv['FRAME37']['A']/2)-zeta, 'YZ plane', 0, 143)
        mprog.add_offset_assembly('FRAME37.1', 'FRAME2.1', apv['FRAME2']['CC']/2, 'ZX plane', 0, 144)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME1.1',
                                  -(apv["FRAME1"]["E"] - apv["FRAME1"]["DD"] - apv["FRAME23"]["I"]), 'XY plane', 0, 145)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME1.1', -(apv['FRAME1']['B'] - apv['FRAME23']['B']) - zeta,
                                  'YZ plane', 0, 146)
        mprog.add_offset_assembly('FRAME23.1', 'FRAME1.1', -(apv['FRAME1']['CC'] / 2 + apv['FRAME23']['A']), 'ZX plane',
                                  1, 147)
        mprog.add_offset_assembly('FRAME23.2', 'FRAME1.1',
                                  -(apv["FRAME1"]["E"] - apv["FRAME1"]["DD"] - apv["FRAME23"]["I"]), 'XY plane', 0, 148)
        mprog.add_offset_assembly('FRAME23.2', 'FRAME1.1', -(apv["FRAME1"]["B"]-apv["FRAME23"]["B"] - assmebly_par['Ass_Y'])-zeta,
                                  'YZ plane', 0, 149)
        mprog.add_offset_assembly('FRAME23.2', 'FRAME1.1', -(apv['FRAME1']['CC'] / 2 + apv['FRAME23']['A']), 'ZX plane',
                                  1, 150)
        mprog.add_offset_assembly('FRAME23.3', 'FRAME2.1',
                                  -(apv["FRAME1"]["E"] - apv["FRAME1"]["DD"] - apv["FRAME23"]["I"]), 'XY plane', 0, 151)
        mprog.add_offset_assembly('FRAME23.3', 'FRAME2.1', (apv['FRAME1']['B']) + zeta,
                                  'YZ plane', 1, 152)
        mprog.add_offset_assembly('FRAME23.3', 'FRAME2.1', -(apv['FRAME1']['CC'] / 2 + apv['FRAME23']['A']), 'ZX plane',
                                  0, 153)
        mprog.add_offset_assembly('FRAME23.4', 'FRAME2.1',
                                  -(apv["FRAME1"]["E"] - apv["FRAME1"]["DD"] - apv["FRAME23"]["I"]), 'XY plane', 0, 154)
        mprog.add_offset_assembly('FRAME23.4', 'FRAME2.1', (apv["FRAME1"]["B"] - assmebly_par['Ass_Y'])+zeta,
                                  'YZ plane', 1, 155)
        mprog.add_offset_assembly('FRAME23.4', 'FRAME2.1', -(apv['FRAME2']['CC'] / 2 + apv['FRAME23']['A']), 'ZX plane',
                                  0, 156)

    #更新
    mprog.update()
    mprog.Close_All()
    # 儲存Product檔
    mprog.saveas(path, 'Product1', '.CATProduct')