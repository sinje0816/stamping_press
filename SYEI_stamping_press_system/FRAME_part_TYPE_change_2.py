import win32com.client as win32
import main_program as mprog
import excel_parameter_change_1 as epc
def FRAME_change_parameter2(name, i):
    if name== 'FRAME3':
        excel = epc.ExcelOp('FRAME3')
        try:
            excel.part_parameter('FRAME3', i)
            print('FRAME3 Parameter change success')
        except:
            print('FRAME3 Parameter change error')
        try:
            if i== 0:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
            elif i== 1:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')

            elif i== 2:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
            elif i == 3:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
            elif i == 4:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
            elif i== 5:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
            elif i == 6:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
            elif i == 7:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
            elif i == 8:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
        except:
            print('FRAME3 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME3 Update success')
            except:
                print('FRAME3 Update error')
    elif name== 'FRAME4':#已更改
        excel = epc.ExcelOp('FRAME4')
        try:
            excel.part_parameter('FRAME4', i)
            print('FRAME4 Parameter change success')
        except:
            print('FRAME4 Parameter change error')
        try:
            if i== 0:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
            elif i== 1:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
            elif i== 2:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
            elif i== 3:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 2)
            elif i== 4:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 5:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 6:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 7:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 8:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 4)
                mprog.activatefeature('Hole_4', 0)
        except:
            print('FRAME4 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME4 Update success')
            except:
                print('FRAME4 Update error')
    elif name== 'FRAME6':
        excel = epc.ExcelOp('FRAME6')
        try:
            excel.part_parameter('FRAME6', i)
            print('FRAME6 Parameter change success')
        except:
            print('FRAME6 Parameter change error')
        try:
            if i==0:
                mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1_i', 0)
            elif i== 1:
                mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1_i', 0)
            elif i== 2:
                mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1_i', 0)
            elif i == 3:
                mprog.activatefeature('FRAME_SN1_60_Body', 0)
                mprog.activatefeature('Hole_1_h FRAME9_L', 0)
            elif i == 4:
                mprog.activatefeature('FRAME_SN1_80110_Body', 0)
                mprog.activatefeature('Hole_1_h FRAME9_L', 0)
            elif i == 5:
                mprog.activatefeature('FRAME_SN1_80110_Body', 0)
                mprog.activatefeature('Hole_1_h FRAME9_L', 0)
            elif i == 7:
                mprog.activatefeature('FRAME_SN1_200_Body', 0)
                mprog.activatefeature('Hole_1_h FRAME9_L', 0)
        except:
            print('FRAME6 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME6 Update success')
            except:
                print('FRAME6 Update error')
    elif name== 'FRAME7':#已更改
        excel = epc.ExcelOp('FRAME7')
        try:
            excel.part_parameter('FRAME7', i)
            print('FRAME7 Parameter change success')
        except:
            print('FRAME7 Parameter change error')
        try:
            if i== 0:
                mprog.activatefeature('SN1_2560_Body', 1)
                mprog.partbodyfeatureactivate('SN1_25_VWX')
            elif i== 1:
                mprog.activatefeature('SN1_2560_Body', 1)
            elif i== 2:
                mprog.activatefeature('SN1_2560_Body', 1)
            elif i== 3:
                mprog.activatefeature('SN1_2560_Body', 1)
            elif i== 4:
                mprog.activatefeature('SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i== 5:
                mprog.activatefeature('SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i== 6:
                mprog.activatefeature('SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i== 7:
                mprog.activatefeature('SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i== 8:
                mprog.activatefeature('SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
        except:
            print('FRAME7 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME7 Update success')
            except:
                print('FRAME7 Update error')

FRAME_change_parameter2('FRAME7' , 7)