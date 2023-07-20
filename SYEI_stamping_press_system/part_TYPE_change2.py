import win32com.client as win32
import main_program as mprog
import excel_parameter_change_1 as epc
def change_parameter2(name, i):
    if name== 'FRAME3':
        excel = epc.ExcelOp('FRAME3')
        try:
            excel.part_parameter('FRAME3', i)
            print('FRAME3 Parameter change success')
        except:
            print('FRAME3 Parameter change error')
        try:
            if i== 0:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
            elif i== 1:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
            elif i== 2:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
            elif i == 3:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
            elif i == 4:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
            elif i== 5:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
            elif i == 6:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
            elif i == 7:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
            elif i == 8:
                mprog.activatefeature('SN1_25250_Body', 4)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
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
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 3)
                mprog.activatefeature('Hole_2', 0)
            elif i== 1:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 2)
                mprog.activatefeature('Hole_2', 0)
            elif i== 2:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
            elif i== 3:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 2)
            elif i== 4:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 5:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 6:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 7:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 8:
                mprog.activatefeature('SN1_25250_Body', 0)
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
                mprog.activatefeature('SN1_2545_Body', 0)
            elif i== 1:
                mprog.activatefeature('SN1_2545_Body', 0)
            elif i== 2:
                mprog.activatefeature('SN1_2545_Body', 0)
            elif i == 3:
                mprog.activatefeature('SN1_60_Body', 0)
            elif i == 4:
                mprog.activatefeature('SN1_80110_Body', 0)
            elif i == 5:
                mprog.activatefeature('SN1_80110_Body', 0)
            elif i == 7:
                mprog.activatefeature('SN1_200_Body', 0)
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
                mprog.activatefeature('SN1_2560_Body', 0)
            elif i== 1:
                mprog.activatefeature('SN1_2560_Body', 0)
            elif i== 2:
                mprog.activatefeature('SN1_2560_Body', 0)
            elif i== 3:
                mprog.activatefeature('SN1_2560_Body', 0)
            elif i== 4:
                mprog.activatefeature('SN1_80250_Body', 7)
                mprog.activatefeature('Hole_1', 0)
            elif i== 5:
                mprog.activatefeature('SN1_80250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 6:
                mprog.activatefeature('SN1_80250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 7:
                mprog.activatefeature('SN1_80250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 8:
                mprog.activatefeature('SN1_80250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
        except:
            print('FRAME7 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME7 Update success')
            except:
                print('FRAME7 Update error')
    elif name== 'FRAME9':
        excel = epc.ExcelOp('FRAME9')
        try:
            excel.part_parameter('FRAME9', i)
            print('FRAME9 Parameter change success')
        except:
            print('FRAME9 Parameter change error')
        try:
            if i== 0:
                mprog.activatefeature('SN1_2580_Body', 0)
            elif i== 1:
                mprog.activatefeature('SN1_2580_Body', 0)
            elif i== 2:
                mprog.activatefeature('SN1_2580_Body', 0)
            elif i== 3:
                mprog.activatefeature('SN1_2580_Body', 0)
            elif i== 4:
                mprog.activatefeature('channel steel', 0)
            elif i== 5:
                mprog.activatefeature('channel steel', 0)
            elif i== 6:
                mprog.activatefeature('channel steel', 0)
            elif i== 7:
                mprog.activatefeature('channel steel', 0)
            elif i== 8:
                mprog.activatefeature('channel steel', 0)
        except:
            print('FRAME9 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME9 Update success')
            except:
                print('FRAME9 Update error')
    elif name== 'FRAME10':#已更改
        excel = epc.ExcelOp('FRAME10')
        try:
            excel.part_parameter('FRAME10', i)
            print('FRAME10 Parameter change success')
        except:
            print('FRAME10 Parameter change error')
        try:
            if i== 0:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 1:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 2:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 3:
                mprog.activatefeature('SN1_60_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 4:
                mprog.activatefeature('SN1_80110_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 5:
                mprog.activatefeature('SN1_80110_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 6:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 7:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 8:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
        except:
            print('FRAME10 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME10 Update success')
            except:
                print('FRAME10 Update error')
    elif name== 'FRAME13':
        excel = epc.ExcelOp('FRAME13')
        try:
            excel.part_parameter('FRAME13', i)
            print('FRAME13 Parameter change success')
        except:
            print('FRAME13 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_25_45_CD')
                mprog.partbodyfeatureactivate('SN1_25_45_G')
                mprog.partbodyfeatureactivate('F')
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_25_45_CD')
                mprog.partbodyfeatureactivate('SN1_25_45_G')
                mprog.partbodyfeatureactivate('F')
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 8:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('Hole_4', 0)
        except:
            print('FRAME13 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME13 Update success')
            except:
                print('FRAME13 Update error')
    elif name== 'FRAME19':#已更改
        excel = epc.ExcelOp('FRAME19')
        try:
            excel.part_parameter('FRAME19', i)
            print('FRAME19 Parameter change success')
        except:
            print('FRAME19 Parameter change error')
        try:
            if i== 0:
                mprog.activatefeature('SN1_2545_Body', 8)
                mprog.partbodyfeatureactivate('SN1_25_X')
                mprog.partbodyfeatureactivate('SN1_2535_Y')
            elif i == 1:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.partbodyfeatureactivate('SN1_2535_Y')
            elif i== 2:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.partbodyfeatureactivate('SN1_45_Y')
                mprog.partbodyfeatureactivate('SN1_45_X')
            elif i== 3:
                mprog.activatefeature('SN1_60_Body', 0)
            elif i== 4:
                mprog.activatefeature('SN1_80110_Body', 0)
            elif i== 5:
                mprog.activatefeature('SN1_80110_Body', 0)
            elif i== 6:
                mprog.activatefeature('FRMAE_SN1_160_Body', 0)
                mprog.activatefeature('FRAME_34_Hole_1', 0)
            elif i== 7:
                mprog.activatefeature('SN1_200_Body', 0)
            elif i== 8:
                mprog.activatefeature('SN1_250_Body', 0)
        except:
            print('FRAME19 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME19 Update success')
            except:
                print('FRAME19 Update error')
    elif name== 'FRAME20':
        excel = epc.ExcelOp('FRAME20')
        try:
            excel.part_parameter('FRAME20', i)
            print('FRAME20 Parameter change success')
        except:
            print('FRAME20 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_45250_C')
                mprog.activatefeature('Hole_1', 0)
            elif i== 4:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.partbodyfeatureactivate('SN1_80250_DEF')
                mprog.partbodyfeatureactivate('SN1_80250_G')
                mprog.activatefeature('Hole_1', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.partbodyfeatureactivate('SN1_80250_DEF')
                mprog.partbodyfeatureactivate('SN1_80250_G')
                mprog.partbodyfeatureactivate('SN1_110160_ijk')
                mprog.activatefeature('Hole_2', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.partbodyfeatureactivate('SN1_80250_DEF')
                mprog.partbodyfeatureactivate('SN1_80250_G')
                mprog.partbodyfeatureactivate('SN1_110160_ijk')
                mprog.activatefeature('Hole_2', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.partbodyfeatureactivate('SN1_80250_DEF')
                mprog.partbodyfeatureactivate('SN1_80250_G')
                mprog.partbodyfeatureactivate('SN1_200250_hij')
                mprog.activatefeature('Hole_2', 1)
            elif i== 8:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.partbodyfeatureactivate('SN1_80250_DEF')
                mprog.partbodyfeatureactivate('SN1_80250_G')
                mprog.partbodyfeatureactivate('SN1_200250_hij')
                mprog.activatefeature('Hole_2', 1)
        except:
            print('FRAME20 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME20 Update success')
            except:
                print('FRAME20 Update error')
    elif name== 'FRAME22':#已更改
        excel = epc.ExcelOp('FRAME22')
        try:
            excel.part_parameter('FRAME22', i)
            print('FRAME22 Parameter change success')
        except:
            print('FRAME22 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.partbodyfeatureactivate('SN1_2545_GH')
                mprog.activatefeature('Hole_1', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.partbodyfeatureactivate('SN1_2545_GH')
                mprog.activatefeature('Hole_1', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.partbodyfeatureactivate('SN1_2545_GH')
                mprog.partbodyfeatureactivate('SN1_45_G(45)H(45)')
                mprog.activatefeature('Hole_1', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_60_I')
                mprog.partbodyfeatureactivate('SN1_60_JK')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i== 4:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                mprog.partbodyfeatureactivate('SN1_80250_JK')
                mprog.partbodyfeatureactivate('SN1_80250_GH')
                mprog.partbodyfeatureactivate('SN1_80250_IN')
                mprog.partbodyfeatureactivate('SN1_80250_LM1')
                mprog.partbodyfeatureactivate('SN1_80250_LM2')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                mprog.partbodyfeatureactivate('SN1_80250_JK')
                mprog.partbodyfeatureactivate('SN1_80250_GH')
                mprog.partbodyfeatureactivate('SN1_80250_IN')
                mprog.partbodyfeatureactivate('SN1_80250_LM1')
                mprog.partbodyfeatureactivate('SN1_80250_LM2')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                mprog.partbodyfeatureactivate('SN1_80250_JK')
                mprog.partbodyfeatureactivate('SN1_80250_GH')
                mprog.partbodyfeatureactivate('SN1_80250_IN')
                mprog.partbodyfeatureactivate('SN1_80250_LM1')
                mprog.partbodyfeatureactivate('SN1_80250_LM2')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                mprog.partbodyfeatureactivate('SN1_80250_JK')
                mprog.partbodyfeatureactivate('SN1_80250_GH')
                mprog.partbodyfeatureactivate('SN1_80250_IN')
                mprog.partbodyfeatureactivate('SN1_80250_LM1')
                mprog.partbodyfeatureactivate('SN1_80250_LM2')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i== 8:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                mprog.partbodyfeatureactivate('SN1_80250_JK')
                mprog.partbodyfeatureactivate('SN1_80250_GH')
                mprog.partbodyfeatureactivate('SN1_80250_IN')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.activatefeature('Hole_3', 0)
        except:
            print('FRAME22 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME22 Update success')
            except:
                print('FRAME22 Update error')
    elif name== 'FRAME25':
        excel = epc.ExcelOp('FRAME25')
        try:
            excel.part_parameter('FRAME25', i)
            print('FRAME25 Parameter change success')
        except:
            print('FRAME25 Parameter change error')
        try:
            if i==0:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i==1:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i==2:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i==3:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
        except:
            print('FRAME25 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME25 Update success')
            except:
                print('FRAME25 Update error')
    elif name== 'FRAME30':#已更改
        excel = epc.ExcelOp('FRAME30')
        try:
            excel.part_parameter('FRAME30', i)
            print('FRAME30 Parameter change success')
        except:
            print('FRAME30 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 4:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 8:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)


        except:
            print('FRAME30 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME30 Update success')
            except:
                print('FRAME30 Update error')
    elif name== 'FRAME31':
        excel = epc.ExcelOp('FRAME31')
        try:
            excel.part_parameter('FRAME31', i)
            print('FRAME31 Parameter change success')
        except:
            print('FRAME31 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('Body')
                mprog.activatefeature('FRAME_Body', 2)
            elif i== 1:
                mprog.partbodyfeatureactivate('Body')
                mprog.activatefeature('FRAME_Body', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('Body')
                mprog.activatefeature('FRAME_Body', 0)
        except:
            print('FRAME31 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME31 Update success')
            except:
                print('FRAME31 Update error')
    elif name== "FRAME33":
        excel = epc.ExcelOp('FRAME33')
        try:
            excel.part_parameter('FRAME33', i)
            print('FRAME33 Parameter change success')
        except:
            print('FRAME33 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.partbodyfeatureactivate('SN1_2545_J')
                mprog.activatefeature('Hole_1', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.partbodyfeatureactivate('SN1_2545_J')
                mprog.activatefeature('Hole_1', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.partbodyfeatureactivate('SN1_2545_J')
                mprog.activatefeature('Hole_1', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
        except:
            print('FRAME33 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME33 Update success')
            except:
                print('FRAME33 Update error')
    elif name== 'FRAME34':
        excel = epc.ExcelOp('FRAME34')
        try:
            excel.part_parameter('FRAME34', i)
            print('FRAME34 Parameter change success')
        except:
            print('FRAME34 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.partbodyfeatureactivate('SN1_2545_L')
                mprog.partbodyfeatureactivate('SN1_2535_J')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.partbodyfeatureactivate('SN1_2545_L')
                mprog.partbodyfeatureactivate('SN1_2535_J')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.partbodyfeatureactivate('SN1_2545_L')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60_NM')
                mprog.partbodyfeatureactivate('SN1_60_L')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)

        except:
            print('FRAME34 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME34 Update success')
            except:
                print('FRAME34 Update error')
    elif name== 'FRAME35':#已更改
        excel = epc.ExcelOp('FRAME35')
        try:
            excel.part_parameter('FRAME35', i)
            print('FRAME35 Parameter change success')
        except:
            print('FRAME35 Parameter change error')
        try:
            if i== 0:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
            elif i== 1:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
            elif i== 2:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
            elif i== 3:
                mprog.activatefeature('SN1_60_Body', 0)
                mprog.activatefeature('Hole', 2)
            elif i== 4:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
            elif i== 5:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
            elif i== 6:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
            elif i== 7:
                mprog.activatefeature('SN1_80200_Body', 6)
                mprog.activatefeature('Hole', 2)
            elif i== 8:
                mprog.activatefeature('SN1_250_Body', 0)
                mprog.activatefeature('Hole', 3)

        except:
            print('FRAME35 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME35 Update success')
            except:
                print('FRAME35 Update error')
    elif name== 'FRAME38':
        excel = epc.ExcelOp('FRAME38')
        try:
            excel.part_parameter('FRAME38', i)
            print('FRAME38 Parameter change success')
        except:
            print('FRAME38 Parameter change error')
        try:
            if i==0:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('Hole_1', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('Hole_1', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('Hole_1', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('Hole_1', 0)
            elif i== 4:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('Hole_1', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('Hole_1', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('Hole_1', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('Hole_1', 0)
            elif i== 8:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('Hole_1', 0)
        except:
            print('FRAME38 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME38 Update success')
            except:
                print('FRAME38 Update error')
    elif name== 'FRAME49':#已更改
        excel = epc.ExcelOp('FRAME49')
        try:
            excel.part_parameter('FRAME49', i)
            print('FRAME49 Parameter change success')
        except:
            print('FRAME49 Parameter change error')
        try:
            if i== 4:
                mprog.activatefeature('body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 5:
                mprog.activatefeature('body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 6:
                mprog.activatefeature('body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 7:
                mprog.activatefeature('body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 8:
                mprog.activatefeature('body', 0)
                mprog.activatefeature('Hole_1', 0)
        except:
            print('FRAME49 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME49 Update success')
            except:
                print('FRAME49 Update error')
    elif name== 'FRAME50':#已更改
        excel = epc.ExcelOp('FRAME50')
        try:
            excel.part_parameter('FRAME50', i)
            print('FRAME50 Parameter change success')
        except:
            print('FRAME50 Parameter change error')
        try:
            if i == 4:
                mprog.activatefeature('Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i == 5:
                mprog.activatefeature('Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 6:
                mprog.activatefeature('Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 7:
                mprog.activatefeature('Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i == 8:
                mprog.activatefeature('Body', 0)
                mprog.activatefeature('Hole_1', 0)
        except:
            print('FRAME50 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME50 Update success')
            except:
                print('FRAME50 Update error')



change_parameter2('FRAME19', 0)



