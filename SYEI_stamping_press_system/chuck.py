import main_program as mprog
import excel_parameter_change as epc


def chuck_parameter_design_part(chuck_type):
    excel = epc.ExcelOp('鎖模夾頭尺寸表', '鎖模夾頭')
    if chuck_type == '302B16':
        chuck_number = 0
    if chuck_type == '322B16':
        chuck_number = 1
    if chuck_type == '342B16':
        chuck_number = 2
    if chuck_type == '372B16':
        chuck_number = 3
    if chuck_type == '392B16':
        chuck_number = 4
    if chuck_type == '412B16':
        chuck_number = 5
    if chuck_type == '43B0016':
        chuck_number = 6
    if chuck_type == '45B0016':
        chuck_number = 7
    if chuck_type == '472B0016':
        chuck_number = 8

    try:
        parameter_name, parameter_value = excel.get_sheet_par('chuck', chuck_number)
        print('chuck parameter change success')
    except:
        print('chuck parameter change error')
    try:
        if chuck_number != 0:
            mprog.partbodyfeatureactivate('EdgeFillet.1')
        else:
            mprog.partdeactivate('EdgeFillet.1')
        print('chuck edgefillet success')
    except:
        print('chuck edgefillet error')
    try:
        if chuck_number == 5:
            mprog.partbodyfeatureactivate('Pocket.4')
        else:
            mprog.partdeactivate('Pocket.4')
        print('chuck rectangle success')
    except:
        print('chuck rectangle error')
    finally:
        try:
            mprog.update()
            print('chuck update success')
        except:
            print('chuck update error')


chuck_parameter_design_part('412B16')
