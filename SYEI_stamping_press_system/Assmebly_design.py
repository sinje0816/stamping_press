import main_program as mprog
import excel_parameter_change as epc
def assembly(i, apv):
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
    part = epc.ExcelOp("尺寸整理表", "沖床機架零件清單")
    #讀取零件匯入數量
    part_name, part_quantity = part.get_assmebly_quantity(i)
    for name in range(len(part_name)):
        for x in range(int(part_quantity[name])):
            mprog.import_file_Part('Z:\\SN1-25_17_17_42_37\\machining', part_name[name])
    #定義基準零件
    mprog.base_lock('FRAME1.1', 'FRAME1.1', 0)
    if i == 0:
        mprog.add_offset_assembly('FRAME35.1', 'FRAME1.1', 0, 'xy plane', 1, 1)
        mprog.add_offset_assembly('FRAME35.1', 'FRAME1.1', apv["FRAME1"]["B"], 'yz plane', 0, 2)
        mprog.add_offset_assembly('FRAME35.1', 'FRAME1.1', (apv['FRAME22']['D']-assmebly_par["Ass_A"])/2+(apv['FRAME1']['CC']/2), 'yz plane', 1, 3)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME1.1', 0, 'xy plane', 0, 4)
        mprog.add_offset_assembly('FRAME2.1', 'FRAME1.1', 0, 'yz plane', 0, 5)
        print(apv['FRAME22']['D'])
        print(apv['FRAME22']['D']+(apv['FRAME1']['CC']/2)+(apv['FRAME2']['CC']/2))
        mprog.add_offset_assembly('FRAME2.1', 'FRAME1.1', apv['FRAME22']['D']+(apv['FRAME1']['CC']/2)+(apv['FRAME2']['CC']/2), 'zx plane', 0, 6)
        mprog.add_offset_assembly('FRAME35.2', 'FRAME2.1', apv['FRAME35']['K']-apv['FRAME35']['L'], 'xy plane', 0, 7)
        mprog.add_offset_assembly('FRAME35.2', 'FRAME2.1', -(apv['FRAME2']['B']), 'yz plane', 0, 8)
        mprog.add_offset_assembly('FRAME35.2', 'FRAME2.1', (apv['FRAME22']['D']-assmebly_par["Ass_A"])/2+apv['FRAME2']['CC']/2, 'yz plane', 0, 9)
        mprog.add_offset_assembly('FRAME10.1', 'FRAME1.1', 0, 'xy plane', 1, 10)
        mprog.add_offset_assembly('FRAME10.1', 'FRAME1.1', -(apv['FRAME10']['B']), 'yz plane', 0, 11)
        mprog.add_offset_assembly('FRAME10.1', 'FRAME1.1', (apv['FRAME22']['D']-assmebly_par["Ass_A"])/2+apv['FRAME1']['CC']/2, 'zx plane', 1, 12)
        mprog.add_offset_assembly('FRAME10.2', 'FRAME2.1', apv["FRAME10"]["K(2)"], 'xy plane', 0, 13)
        mprog.add_offset_assembly('FRAME10.2', 'FRAME2.1', -(apv['FRAME10']['B']), 'yz plane', 0, 14)
        mprog.add_offset_assembly('FRAME10.2', 'FRAME2.1', (apv['FRAME22']['D']-assmebly_par["Ass_A"])/2+apv['FRAME2']['CC']/2, 'zx plane', 0, 15)
        mprog.add_offset_assembly('FRAME22.1', 'FRAME35.1', (apv['FRAME1']['E']-apv['FRAME1']['DD'])-apv['FRAME22']['A(1)']-apv['FRAME22']['A(2)'], 'xy plane', 1, 16)
        mprog.add_offset_assembly('FRAME22.1', 'FRAME35.1', apv['FRAME22']['O'], 'yz plane', 0, 17)
        mprog.add_offset_assembly('FRAME22.1', 'FRAME35.1', apv['FRAME22']['D']+apv['FRAME1']['CC']/2, 'zx plane', 0, 18)
        mprog.add_offset_assembly('FRAME33.1', 'FRAME1.1', 0, 'xy plane', 0, 19)
        mprog.add_offset_assembly('FRAME33.1', 'FRAME1.1', -(apv['FRAME1']['B']-apv['FRAME33']['B']), 'yz plane', 0, 20)
        mprog.add_offset_assembly('FRAME33.1', 'FRAME1.1', apv['FRAME33']['K']+apv['FRAME1']['CC']/2, 'zx plane', 0, 21)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME2.1', 0, 'xy plane', 0, 22)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME2.1', -(apv['FRAME2']['B']-apv['FRAME34']['B']), 'yz plane', 0, 23)
        mprog.add_offset_assembly('FRAME34.1', 'FRAME2.1', apv['FRAME2']['CC']/2, 'zx plane', 0, 24)
        mprog.add_offset_assembly('FRAME7.1', 'FRAME2.1', -((apv['FRAME2']['E']-apv['FRAME2']['DD'])-apv['FRAME7']['A']-apv['FRAME7']['Y']), 'xy plane', 0, 25)
        mprog.add_offset_assembly('FRAME7.1', 'FRAME2.1', -(apv["FRAME2"]["B"]-assmebly_par["Ass_B"]), 'yz plane', 0, 26)
        mprog.add_offset_assembly('FRAME7.1', 'FRAME2.1', assmebly_par["Ass_E"], 'zx plane', 0, 27)
        mprog.add_offset_assembly('FRAME19.1', 'FRAME1.1', -((apv['FRAME1']['E']-apv['FRAME1']['DD'])-apv['FRAME19']['D']), 'xy plane', 0, 28)
        mprog.add_offset_assembly('FRAME19.1', 'FRAME1.1', -(apv['FRAME1']['B']-assmebly_par["Ass_C"]), 'yz plane', 0, 29)
        mprog.add_offset_assembly('FRAME19.1', 'FRAME1.1', -(apv['FRAME1']['CC']/2+apv['FRAME19']['AH']), 'zx plane', 0, 30)
        mprog.add_offset_assembly('FRAME19.2', 'FRAME2.1', -((apv['FRAME1']['E']-apv['FRAME1']['DD'])-apv['FRAME19']['D']), 'xy plane', 0, 31)
        mprog.add_offset_assembly('FRAME19.2', 'FRAME2.1', -(apv['FRAME2']['B']-assmebly_par["Ass_C"]), 'yz plane', 0, 32)
        mprog.add_offset_assembly('FRAME19.2', 'FRAME2.1', apv['FRAME2']['CC']/2, 'zx plane', 0, 33)
        # mprog.add_offset_assembly('FRAME8.1', 'FRAME1.1', -assmebly_par["Ass_E"], 'zx plane', 1, 39)