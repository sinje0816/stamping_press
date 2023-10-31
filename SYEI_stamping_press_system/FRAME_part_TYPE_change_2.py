import win32com.client as win32
import main_program as mprog
import excel_parameter_change_1 as epc
def FRAME_change_parameter2(name, i):
    if name == 'FRAME3':
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
    elif name == 'FRAME4':#已更改
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
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
            elif i== 1:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
            elif i== 2:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
            elif i== 3:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
            elif i== 4:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 5:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 6:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 7:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i== 8:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
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
    elif name == 'FRAME6':
        excel = epc.ExcelOp('FRAME6')
        try:
            excel.part_parameter('FRAME6', i)
            print('FRAME6 Parameter change success')
        except:
            print('FRAME6 Parameter change error')
        try:
            if i==0:
                mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                mprog.activatefeature('M16_i', 0)
            elif i== 1:
                mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                mprog.activatefeature('M16_i', 0)
            elif i== 2:
                mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                mprog.activatefeature('M16_i', 0)
            elif i == 3:
                mprog.activatefeature('FRAME_SN1_60_Body', 0)
                mprog.activatefeature('M16_h FRAME9_L', 0)
            elif i == 4:
                mprog.activatefeature('FRAME_SN1_80110_Body', 0)
                mprog.activatefeature('M16_h FRAME9_L', 0)
            elif i == 5:
                mprog.activatefeature('FRAME_SN1_80110_Body', 0)
                mprog.activatefeature('M16_h FRAME9_L', 0)
            elif i == 7:
                mprog.activatefeature('FRAME_SN1_200_Body', 0)
                mprog.activatefeature('M16_h FRAME9_L', 0)
        except:
            print('FRAME6 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME6 Update success')
            except:
                print('FRAME6 Update error')
    elif name == 'FRAME7':#已更改
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
                mprog.activatefeature('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i== 5:
                mprog.activatefeature('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i== 6:
                mprog.activatefeature('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i== 7:
                mprog.activatefeature('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i== 8:
                mprog.activatefeature('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
        except:
            print('FRAME7 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME7 Update success')
            except:
                print('FRAME7 Update error')
    elif name == 'FRAME9':
        excel = epc.ExcelOp('FRAME9')
        try:
            excel.part_parameter('FRAME9', i)
            print('FRAME9 Parameter change success')
        except:
            print('FRAME9 Parameter change error')
        try:
            if i== 0:
                mprog.activatefeature('SN1_2580_Body', 1)
            elif i== 1:
                mprog.activatefeature('SN1_2580_Body', 1)
            elif i== 2:
                mprog.activatefeature('SN1_2580_Body', 1)
            elif i== 3:
                mprog.activatefeature('SN1_2580_Body', 1)
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
    elif name == 'FRAME10':#已更改
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
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i== 1:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i== 2:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i== 3:
                mprog.activatefeature('SN1_60_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i== 4:
                mprog.activatefeature('SN1_80110_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i== 5:
                mprog.activatefeature('SN1_80110_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i== 6:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i== 7:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i== 8:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
        except:
            print('FRAME10 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME10 Update success')
            except:
                print('FRAME10 Update error')
    elif name == 'FRAME13':
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
                mprog.activatefeature('GIB_OIL_HOLE', 1)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 1)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_25_45_CD')
                mprog.partbodyfeatureactivate('SN1_25_45_G')
                mprog.activatefeature('GIB_OIL_HOLE', 1)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 1)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i== 8:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 0)
                mprog.activatefeature('Hole_4', 0)
        except:
            print('FRAME13 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME13 Update success')
            except:
                print('FRAME13 Update error')
    elif name == 'FRAME19':#已更改
        excel = epc.ExcelOp('FRAME19')
        try:
            excel.part_parameter('FRAME19', i)
            print('FRAME19 Parameter change success')
        except:
            print('FRAME19 Parameter change error')
        try:
            if i== 0:
                mprog.activatefeature('SN1_2545_Body', 6)
                mprog.partbodyfeatureactivate('SN1_25_X')
                mprog.partbodyfeatureactivate('SN1_2535_Y')
            elif i == 1:
                mprog.activatefeature('SN1_2545_Body', 6)
                mprog.partbodyfeatureactivate('SN1_2535_Y')
                mprog.partbodyfeatureactivate("SN1_3545_AD")
            elif i== 2:
                mprog.activatefeature('SN1_2545_Body', 6)
                mprog.partbodyfeatureactivate('SN1_45_Y')
                mprog.partbodyfeatureactivate('SN1_45_X')
                mprog.partbodyfeatureactivate("SN1_3545_AD")
            elif i== 3:
                mprog.activatefeature('SN1_60_Body', 8)
            elif i== 4:
                mprog.activatefeature('SN1_80110_Body', 5)
            elif i== 5:
                mprog.activatefeature('SN1_80110_Body', 5)
            elif i== 6:
                mprog.activatefeature('FRMAE_SN1_160_Body', 0)
                mprog.activatefeature('FRAME_34_Hole_1', 0)
            elif i== 7:
                mprog.activatefeature('SN1_200_Body', 4)
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
    elif name == 'FRAME20':
        excel = epc.ExcelOp('FRAME20')
        try:
            excel.part_parameter('FRAME20', i)
            print('FRAME20 Parameter change success')
        except:
            print('FRAME20 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
            elif i== 4:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.partbodyfeatureactivate('SN1_80250_DEF')
                mprog.partbodyfeatureactivate('SN1_80250_G')
                mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
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
    elif name == 'FRAME22':#已更改
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
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i== 3:
                mprog.activatefeature('FRAME_SN1_60_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i== 4:
                mprog.activatefeature('FRAME_SN1_80250_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i== 5:
                mprog.activatefeature('FRAME_SN1_80250_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i== 6:
                mprog.activatefeature('FRAME_SN1_80250_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i== 7:
                mprog.activatefeature('FRAME_SN1_80250_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i== 8:
                mprog.activatefeature('FRAME_SN1_80250_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
        except:
            print('FRAME22 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME22 Update success')
            except:
                print('FRAME22 Update error')
    elif name == 'FRAME25':
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
                mprog.activatefeature('3_M8通', 0)
            elif i==1:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('3_M8通', 0)
            elif i==2:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('3_M8通', 0)
            elif i==3:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('3_M8通', 0)
        except:
            print('FRAME25 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME25 Update success')
            except:
                print('FRAME25 Update error')
    elif name == 'FRAME28':
        excel = epc.ExcelOp('FRAME28')
        try:
            excel.part_parameter('FRAME28', i)
            print('FRAME28 Parameter change success')
        except:
            print('FRAME28 Parameter change error')
        try:
            if i==0:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i==1:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i==2:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i==3:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i==4:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i==5:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i==6:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i==7:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i==8:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
        except:
            print('FRAME28 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME28 Update success')
            except:
                print('FRAME28 Update error')
    elif name == 'FRAME30':#已更改
        excel = epc.ExcelOp('FRAME30')
        try:
            excel.part_parameter('FRAME30', i)
            print('FRAME30 Parameter change success')
        except:
            print('FRAME30 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i== 4:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i== 8:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)


        except:
            print('FRAME30 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME30 Update success')
            except:
                print('FRAME30 Update error')
    elif name == 'FRAME31':
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
    elif name == 'FRAME32':
        excel = epc.ExcelOp('FRAME32')
        try:
            excel.part_parameter('FRAME32', i)
            print('FRAME32 Parameter change success')
        except:
            print('FRAME32 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('Body')
            elif i== 1:
                mprog.partbodyfeatureactivate('Body')
            elif i== 2:
                mprog.partbodyfeatureactivate('Body')
            elif i== 3:
                mprog.partbodyfeatureactivate('Body')
            elif i== 4:
                mprog.partbodyfeatureactivate('Body')
            elif i== 5:
                mprog.partbodyfeatureactivate('Body')
            elif i== 6:
                mprog.partbodyfeatureactivate('Body')
            elif i== 7:
                mprog.partbodyfeatureactivate('Body')
            elif i== 8:
                mprog.partbodyfeatureactivate('Body')
        except:
            print('FRAME32 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME32 Update success')
            except:
                print('FRAME32 Update error')
    elif name == "FRAME33":
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
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('FRAME_2545_LMN')
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('FRAME_2545_LMN')
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('FRAME_2545_LMN')
            elif i== 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('FRAME_60_LMN')
        except:
            print('FRAME33 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME33 Update success')
            except:
                print('FRAME33 Update error')
    elif name == 'FRAME34':
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
                mprog.partbodyfeatureactivate('SN1_2535_J')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_1_2545_PQR')
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.partbodyfeatureactivate('SN1_2535_J')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_1_2545_PQR')
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_1_2545_PQR')
            elif i== 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60_NM')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_1_60_PQR')
        except:
            print('FRAME34 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME34 Update success')
            except:
                print('FRAME34 Update error')
    elif name == 'FRAME35':#已更改
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
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i== 1:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i== 2:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i== 3:
                mprog.activatefeature('SN1_60_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i== 4:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i== 5:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i== 6:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i== 7:
                mprog.activatefeature('SN1_80200_Body', 6)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
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
    elif name == 'FRAME36':#已更改
        excel = epc.ExcelOp('FRAME36')
        try:
            excel.part_parameter('FRAME36', i)
            print('FRAME36 Parameter change success')
        except:
            print('FRAME36 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('FRAME_SN1_25_Body')
            elif i== 1:
                mprog.partbodyfeatureactivate('FRAME_SN1_3560_Body')
            elif i== 2:
                mprog.partbodyfeatureactivate('FRAME_SN1_3560_Body')
            elif i== 3:
                mprog.partbodyfeatureactivate('FRAME_SN1_3560_Body')
        except:
            print('FRAME36 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME36 Update success')
            except:
                print('FRAME36 Update error')
    elif name == 'FRAME37':#已更改
        excel = epc.ExcelOp('FRAME37')
        try:
            excel.part_parameter('FRAME37', i)
            print('FRAME37 Parameter change success')
        except:
            print('FRAME37 Parameter change error')
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('Body')
                mprog.partbodyfeatureactivate('FRAME_Body')
            elif i== 1:
                mprog.partbodyfeatureactivate('Body')
                mprog.partbodyfeatureactivate('FRAME_Body')
            elif i== 2:
                mprog.partbodyfeatureactivate('Body')
                mprog.partbodyfeatureactivate('FRAME_Body')
            elif i== 3:
                mprog.partbodyfeatureactivate('Body')
                mprog.partbodyfeatureactivate('FRAME_Body')
        except:
            print('FRAME37 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME37 Update success')
            except:
                print('FRAME37 Update error')
    elif name == 'FRAME38':
        excel = epc.ExcelOp('FRAME38')
        try:
            excel.part_parameter('FRAME38', i)
            print('FRAME38 Parameter change success')
        except:
            print('FRAME38 Parameter change error')
        try:
            if i==0:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i== 4:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i== 8:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
        except:
            print('FRAME38 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME38 Update success')
            except:
                print('FRAME38 Update error')
    elif name == 'FRAME49':#已更改
        excel = epc.ExcelOp('FRAME49')
        try:
            excel.part_parameter('FRAME49', i)
            print('FRAME49 Parameter change success')
        except:
            print('FRAME49 Parameter change error')
        try:
            if i== 4:
                mprog.activatefeature('body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                mprog.activatefeature('M12通', 0)
            elif i== 5:
                mprog.activatefeature('body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                mprog.activatefeature('M12通', 0)
            elif i== 6:
                mprog.activatefeature('body', 0)
                mprog.activatefeature('M12通', 0)
            elif i== 7:
                mprog.activatefeature('body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                mprog.activatefeature('M12通', 0)
            elif i== 8:
                mprog.activatefeature('body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                mprog.activatefeature('M12通', 0)
        except:
            print('FRAME49 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME49 Update success')
            except:
                print('FRAME49 Update error')
    elif name == 'FRAME50':#已更改
        excel = epc.ExcelOp('FRAME50')
        try:
            excel.part_parameter('FRAME50', i)
            print('FRAME50 Parameter change success')
        except:
            print('FRAME50 Parameter change error')
        try:
            if i == 4:
                mprog.activatefeature('Body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_c')
                mprog.activatefeature('Hole_1', 0)
            elif i == 5:
                mprog.activatefeature('Body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_c')
                mprog.activatefeature('Hole_1', 0)
            elif i== 6:
                mprog.activatefeature('Body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_c')
                mprog.activatefeature('Hole_1', 0)
            elif i== 7:
                mprog.activatefeature('Body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_c')
                mprog.activatefeature('Hole_1', 0)
            elif i == 8:
                mprog.activatefeature('Body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_c')
                mprog.activatefeature('Hole_1', 0)
        except:
            print('FRAME50 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME50 Update success')
            except:
                print('FRAME50 Update error')
    elif name == 'FRAME54':#已更改
        excel = epc.ExcelOp('FRAME54')
        try:
            excel.part_parameter('FRAME54', i)
            print('FRAME54 Parameter change success')
        except:
            print('FRAME54 Parameter change error')
        try:
            if i== 6:
                mprog.partbodyfeatureactivate('Body')
            elif i== 7:
                mprog.partbodyfeatureactivate('Body')
            elif i == 8:
                mprog.partbodyfeatureactivate('Body')
        except:
            print('FRAME54 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME54 Update success')
            except:
                print('FRAME54 Update error')
def BALL_SCREW_change(BALL_SCREW_name):
    excel = epc.ExcelOp('BALL_SCREW')
    if BALL_SCREW_name == '302B07S02':
        BALL_SCREW_number = 0
    elif BALL_SCREW_name == '302B07S03':
        BALL_SCREW_number = 1
    elif BALL_SCREW_name == '322B07S02':
        BALL_SCREW_number = 2
    elif BALL_SCREW_name == '322B07S03':
        BALL_SCREW_number = 3
    elif BALL_SCREW_name == '342B07S01':
        BALL_SCREW_number = 4
    elif BALL_SCREW_name == '342B07':
        BALL_SCREW_number = 5
    elif BALL_SCREW_name == '372B07S02':
        BALL_SCREW_number = 6
    elif BALL_SCREW_name == '372B07S04':
        BALL_SCREW_number = 7
    elif BALL_SCREW_name == '372B07':
        BALL_SCREW_number = 8
    elif BALL_SCREW_name == '392B07S06':
        BALL_SCREW_number = 9
    elif BALL_SCREW_name == '392B07S07':
        BALL_SCREW_number = 10
    elif BALL_SCREW_name == '392B07S08':
        BALL_SCREW_number = 11
    elif BALL_SCREW_name == '412B07S01':
        BALL_SCREW_number = 12
    elif BALL_SCREW_name == '412B07S02':
        BALL_SCREW_number = 13
    elif BALL_SCREW_name == '412B07':
        BALL_SCREW_number = 14
    elif BALL_SCREW_name == '432B07S01':
        BALL_SCREW_number = 15
    elif BALL_SCREW_name == '432B07S02':
        BALL_SCREW_number = 16
    elif BALL_SCREW_name == '432B07':
        BALL_SCREW_number = 17
    elif BALL_SCREW_name == '452B07S03':
        BALL_SCREW_number = 18
    elif BALL_SCREW_name == '452B07S04':
        BALL_SCREW_number = 19
    elif BALL_SCREW_name == '452B07S05':
        BALL_SCREW_number = 20
    elif BALL_SCREW_name == '472B07S03':
        BALL_SCREW_number = 21
    elif BALL_SCREW_name == '472B07S04':
        BALL_SCREW_number = 22
    elif BALL_SCREW_name == '472B07S05':
        BALL_SCREW_number = 23
    try:
        excel.part_parameter('BALL_SCREW', BALL_SCREW_number)
        print('BALL_SCREW Parameter change success')
    except:
        print('BALL_SCREW Parameter change error')
    try:
        if BALL_SCREW_name == "302B07S02":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_T')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "302B07S03":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_T')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "322B07S02":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_T')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "322B07S03":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_T')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "342B07S01":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_T')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "342B07":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_T')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "372B07S02":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "372B07S04":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "372B07":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "392B07S06":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "392B07S07":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "392B07S08":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "412B07S01":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "412B07S02":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "412B07":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "432B07S01":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "432B07S02":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "432B07":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "452B07S03":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('Groove_AY')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "452B07S04":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('Groove_AY')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "452B07S05":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('Groove_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('Groove_AY')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "472B07S03":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('EdgeFillet_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
            mprog.partbodyfeatureactivate('EdgeFillet_AZ')
        elif BALL_SCREW_name == "472B07S04":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('EdgeFillet_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
        elif BALL_SCREW_name == "472B07S05":
            mprog.partbodyfeatureactivate('BALL_SCREW_Body')
            mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
            mprog.partbodyfeatureactivate('AE')
            mprog.partbodyfeatureactivate('AB')
            mprog.partbodyfeatureactivate('Hole_AAAC')
            mprog.partbodyfeatureactivate('EdgeFillet_O')
            mprog.partbodyfeatureactivate('Chamfer_P')
            mprog.partbodyfeatureactivate('EdgeFillet_V')
            mprog.partbodyfeatureactivate('EdgeFillet_X')
            mprog.partbodyfeatureactivate('EdgeFillet_Y')
            mprog.partbodyfeatureactivate('EdgeFillet_AD')
            mprog.partbodyfeatureactivate('Chamfer_AUAV')
            mprog.partbodyfeatureactivate('EdgeFillet_AW')
            mprog.partbodyfeatureactivate('EdgeFillet_AX')
            mprog.partbodyfeatureactivate('AFAH')
            mprog.partbodyfeatureactivate('Hole_AJ')
    except:
        print('BALL_SCREW Parameter activate error')
    finally:
        try:
            mprog.Update()
            print('BALL_SCREW Update success')
        except:
            print('BALL_SCREW Update error')



FRAME_change_parameter2("FRAME20", 0)