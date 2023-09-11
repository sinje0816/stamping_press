import main_program as mprog
import excel_parameter_change as epc

def clamping_bolt_parameter_design_part(i):
    excel = epc.ExcelOp('螺栓尺寸表', '鎖模螺栓')
    try:
        parameter_name, paramter_value = excel.get_sheet_par('clamping_bolt', i)
        print('clamping_bolt parameter change success')
    except:
        print('clamping_bolt parameter change error')
    finally:
        try:
            mprog.update()
            print('bolt update success')
        except:
            print('bolt update error')

clamping_bolt_parameter_design_part(8)

