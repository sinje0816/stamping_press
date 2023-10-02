import excel_parameter_change as epc
import main_program as mprog
import parameter as par


def padchange(i):
    excel = epc.ExcelOp('平板', '平板')
    try:
        plate_parameter_name, plate_parameter_value = excel.get_sheet_par('plate', i)
        print(plate_parameter_value, plate_parameter_name)
        print('plate Parameter change success')
    except:
        print('plate Parameter change error')
    try:
        if i == 4 or 5 or 6 or 7 or 8:  # 80N
            mprog.activatefeature('Bottom_80N', 0)
    except:
        print('plate Parameter activate error')
    finally:
        try:
            mprog.Update()
            print('plate Update success')
        except:
            print('plate Update error')

    return plate_parameter_name, plate_parameter_value
