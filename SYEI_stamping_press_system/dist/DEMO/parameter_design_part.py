import excel_parameter_change as epc
import main_program as mprog
import parameter as par


def padchange(stamping_press_type):
    excel = epc.ExcelOp('平板', '平板')
    plate_parameter_name, plate_parameter_value = excel.get_sheet_par('plate', stamping_press_type)
    if stamping_press_type == 0 or stamping_press_type == 1 or stamping_press_type == 2 or stamping_press_type == 3:  # 60N
        mprog.activatefeature('Bottom_60N', 0)
    elif stamping_press_type == 4 or stamping_press_type == 5 or stamping_press_type ==  6 or stamping_press_type ==  7 or stamping_press_type ==  8:  # 80N
        mprog.activatefeature('Bottom_80N', 0)
        mprog.partbodyfeatureactivate('80N')
        mprog.partbodyfeatureactivate('80N_EdgeFillet')
    mprog.Update()
    return plate_parameter_name, plate_parameter_value
