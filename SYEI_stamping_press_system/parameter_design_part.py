import excel_parameter_change as epc
import main_program as mprog


def padchange(i):
    excel = epc.ExcelOp('平板', '平板')
    try:
        plate_parameter_name, plate_parameter_value = excel.get_sheet_par('plate', i)
        print('plate Parameter change success')
    except:
        print('plate Parameter change error')
    try:
        if i == 0:  # 25N
            mprog.partbodyfeatureactivate('CH_60N')
            mprog.partbodyfeatureactivate('V')
            mprog.activatefeature('f', 0)
            mprog.activatefeature('Bottom_60N', 0)
            mprog.activatefeature('C', 0)

        if i == 1:  # 35N
            mprog.partbodyfeatureactivate('CH_60N')
            mprog.partbodyfeatureactivate('V')
            mprog.activatefeature('f', 0)
            mprog.activatefeature('Bottom_60N', 0)
            mprog.activatefeature('C', 0)

        if i == 2:  # 45N
            mprog.partbodyfeatureactivate('CH_60N')
            mprog.partbodyfeatureactivate('V')
            mprog.activatefeature('f', 0)
            mprog.activatefeature('Bottom_60N', 0)
            mprog.activatefeature('C', 0)

        if i == 3:  # 60N
            mprog.partbodyfeatureactivate('CH_60N')
            mprog.partbodyfeatureactivate('V')
            mprog.activatefeature('f', 0)
            mprog.activatefeature('Bottom_60N', 0)
            mprog.activatefeature('C', 0)

        if i == 4:  # 80N
            mprog.partbodyfeatureactivate('V')
            mprog.activatefeature('fr', 0)
            mprog.activatefeature('CH_80N', 0)
            mprog.activatefeature('Bottom_60N', 0)
            mprog.activatefeature('Bottom_80N', 0)
            mprog.activatefeature('C', 0)

        if i == 5:  # 110N
            mprog.partbodyfeatureactivate('V')
            mprog.activatefeature('fr', 0)
            mprog.activatefeature('CH_80N', 0)
            mprog.activatefeature('Bottom_60N', 0)
            mprog.activatefeature('Bottom_80N', 0)
            mprog.activatefeature('C', 0)

        if i == 6:  # 160N
            mprog.partbodyfeatureactivate('V')
            mprog.activatefeature('fr', 0)
            mprog.activatefeature('CH_80N', 0)
            mprog.activatefeature('Bottom_60N', 0)
            mprog.activatefeature('Bottom_80N', 0)
            mprog.activatefeature('C', 0)

        if i == 7:  # 200N
            mprog.partbodyfeatureactivate('V')
            mprog.activatefeature('fr', 0)
            mprog.activatefeature('CH_80N', 0)
            mprog.activatefeature('Bottom_60N', 0)
            mprog.activatefeature('Bottom_80N', 0)
            mprog.activatefeature('C', 0)

        if i == 8:  # 250N
            mprog.partbodyfeatureactivate('V')
            mprog.activatefeature('fr', 0)
            mprog.activatefeature('CH_80N', 0)
            mprog.activatefeature('Bottom_60N', 0)
            mprog.activatefeature('Bottom_80N', 0)
            mprog.activatefeature('C', 0)

    except:
        print('plate Parameter activate error')
    finally:
        try:
            mprog.Update()
            print('plate Update success')
        except:
            print('plate Update error')

    return plate_parameter_name, plate_parameter_value