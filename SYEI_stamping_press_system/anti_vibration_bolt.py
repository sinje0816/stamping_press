import main_program as mprog
import excel_parameter_change as epc


def anti_vibration_bolt_parameter_design_part(anti_vibration_bolt_tpye):
    excel = epc.ExcelOp('螺栓尺寸表', '防震螺栓_SN1用')
    if anti_vibration_bolt_tpye == 'CRMA201506KK':
        anti_vibration_bolt_number = 0
    if anti_vibration_bolt_tpye == 'CRMA301808KK':
        anti_vibration_bolt_number = 1
    if anti_vibration_bolt_tpye == 'CRMA323710YS':
        anti_vibration_bolt_number = 2

    try:
        parameter_name, paramter_value = excel.get_sheet_par('anti_vibration_bolt', anti_vibration_bolt_number)
        print('anti_vibration_bolt parameter change success')
    except:
        print('anti_vibration_bolt parameter change error')
    finally:
        try:
            mprog.update()
            print('anti_vibration_bolt update success')
        except:
            print('anti_vibration_bolt update error')


anti_vibration_bolt_parameter_design_part('CRMA323710YS')
