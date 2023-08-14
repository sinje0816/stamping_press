import win32com.client as win32
import main_program as mprog
import excel_parameter_change_1 as epc
def change_parameter2(name, i):
    if name == 'FRAME3':
        try:
            mprog.bodydeactivate('Hole_1', 0)
            mprog.bodydeactivate('Hole_2', 0)
        except:
            print('FRAME3 Parameter change error')
        try:
            if i == 0:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
                mprog.partdeactivate('FRAME_SN1_25250_PQ')
                mprog.partdeactivate('FRAME_SN1_2560_S')
                mprog.partdeactivate('FRAME_SN1_2560_U')
                mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
            elif i == 1:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
                mprog.partdeactivate('FRAME_SN1_25250_PQ')
                mprog.partdeactivate('FRAME_SN1_2560_S')
                mprog.partdeactivate('FRAME_SN1_2560_U')
                mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
            elif i == 2:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
                mprog.partdeactivate('FRAME_SN1_25250_PQ')
                mprog.partdeactivate('FRAME_SN1_2560_S')
                mprog.partdeactivate('FRAME_SN1_2560_U')
                mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
            elif i == 3:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
                mprog.partdeactivate('FRAME_SN1_25250_PQ')
                mprog.partdeactivate('FRAME_SN1_2560_S')
                mprog.partdeactivate('FRAME_SN1_2560_U')
                mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
            elif i == 4:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
                mprog.partdeactivate('FRAME_SN1_25250_PQ')
                mprog.partdeactivate('FRAME_SN1_80250_R')
                mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
            elif i == 5:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
                mprog.partdeactivate('FRAME_SN1_25250_PQ')
                mprog.partdeactivate('FRAME_SN1_80250_R')
                mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
            elif i == 6:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
                mprog.partdeactivate('FRAME_SN1_25250_PQ')
                mprog.partdeactivate('FRAME_SN1_80250_R')
                mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
            elif i == 7:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
                mprog.partdeactivate('FRAME_SN1_25250_PQ')
                mprog.partdeactivate('FRAME_SN1_80250_R')
                mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
            elif i == 8:
                mprog.activatefeature('SN1_25250_Body', 4)
                mprog.activatefeature('SN1_25250_M', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 1)
                mprog.partdeactivate('FRAME_SN1_25250_PQ')
                mprog.partdeactivate('FRAME_SN1_80250_R')
                mprog.bodydeactivate('M10X20L(不可貫穿)', 0)
        except:
            print('FRAME3 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME3 Update error')
    elif name == 'FRAME4':#已更改
        try:
            if i == 0:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 3)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
            elif i == 1:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 2)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
            elif i == 2:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
            elif i == 3:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
            elif i == 4:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i == 5:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i == 6:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i == 7:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i == 8:
                mprog.activatefeature('SN1_25250_Body', 0)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 4)
                mprog.activatefeature('Hole_4', 0)
        except:
            print('FRAME4 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME4 Update error')
    elif name == 'FRAME6':
        try:
            if i == 0:
                mprog.bodydeactivate('FRAME_SN1_2545_Body', 0)
                mprog.bodydeactivate('M16_i', 0)
                mprog.activatefeature('SN1_2545_Body', 0)
            elif i == 1:
                mprog.bodydeactivate('FRAME_SN1_2545_Body', 0)
                mprog.bodydeactivate('M16_i', 0)
                mprog.activatefeature('SN1_2545_Body', 0)
            elif i == 2:
                mprog.bodydeactivate('FRAME_SN1_2545_Body', 0)
                mprog.bodydeactivate('M16_i', 0)
                mprog.activatefeature('SN1_2545_Body', 0)
            elif i == 3:
                mprog.bodydeactivate('FRAME_SN1_60_Body', 0)
                mprog.bodydeactivate('M16_h FRAME9_L', 0)
                mprog.activatefeature('SN1_60_Body', 0)
            elif i == 4:
                mprog.bodydeactivate('FRAME_SN1_80110_Body', 0)
                mprog.bodydeactivate('M16_h FRAME9_L', 0)
                mprog.activatefeature('SN1_80110_Body', 0)
            elif i == 5:
                mprog.bodydeactivate('FRAME_SN1_80110_Body', 0)
                mprog.bodydeactivate('M16_h FRAME9_L', 0)
                mprog.activatefeature('SN1_80110_Body', 0)
            elif i == 7:
                mprog.bodydeactivate('FRAME_SN1_200_Body', 0)
                mprog.bodydeactivate('M16_h FRAME9_L', 0)
                mprog.activatefeature('SN1_200_Body', 0)
        except:
            print('FRAME6 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME6 Update error')
    elif name == 'FRAME7':
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
                mprog.bodydeactivate('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('SN1_80250_Body', 7)
                mprog.activatefeature('Hole_1', 0)
            elif i== 5:
                mprog.bodydeactivate('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('SN1_80250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 6:
                mprog.bodydeactivate('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('SN1_80250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 7:
                mprog.bodydeactivate('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('SN1_80250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
            elif i== 8:
                mprog.bodydeactivate('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('SN1_80250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
        except:
            print('FRAME7 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME7 Update error')
    elif name == 'FRAME9':
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
            except:
                print('FRAME9 Update error')
    elif name == 'FRAME10':
        try:
            if i== 0:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.partdeactivate('SN1_25250_K')
            elif i== 1:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.partdeactivate('SN1_25250_K')
            elif i== 2:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.partdeactivate('SN1_25250_K')
            elif i== 3:
                mprog.activatefeature('SN1_60_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partdeactivate('SN1_25250_K')
            elif i== 4:
                mprog.activatefeature('SN1_80110_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partdeactivate('SN1_25250_K')
            elif i== 5:
                mprog.activatefeature('SN1_80110_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partdeactivate('SN1_25250_K')
            elif i== 6:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partdeactivate('SN1_25250_K')
            elif i== 7:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partdeactivate('SN1_25250_K')
            elif i== 8:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partdeactivate('SN1_25250_K')
        except:
            print('FRAME10 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME10 Update error')
    elif name == 'FRAME13':
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_25_45_CD')
                mprog.partbodyfeatureactivate('SN1_25_45_G')
                mprog.partbodyfeatureactivate('F')
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.bodydeactivate('Hole_4', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 1)
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.bodydeactivate('Hole_4', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 1)
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_25_45_CD')
                mprog.partbodyfeatureactivate('SN1_25_45_G')
                mprog.partbodyfeatureactivate('F')
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.bodydeactivate('Hole_4', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 1)
            elif i == 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.bodydeactivate('Hole_4', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 1)
            elif i == 4:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.bodydeactivate('Hole_4', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.bodydeactivate('Hole_4', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.bodydeactivate('Hole_4', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.bodydeactivate('Hole_4', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 0)
            elif i== 8:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('F')
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.bodydeactivate('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.bodydeactivate('Hole_4', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 0)
        except:
            print('FRAME13 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME13 Update error')
    elif name == 'FRAME15':
        try:
            if i == 0:  # 25N
                    mprog.activatefeature('Hole', 0)
            elif i == 1:  # 35N
                    mprog.activatefeature('Hole', 0)
            elif i == 2:  # 45N
                    mprog.activatefeature('Hole', 0)
            elif i == 3:  # 60N
                    mprog.activatefeature('Hole', 0)
            elif i == 4:  # 80N
                    mprog.activatefeature('Hole', 0)
            elif i == 5:  # 110N
                    mprog.activatefeature('Hole', 0)
            elif i == 6:  # 160N
                    mprog.activatefeature('Hole', 0)
            elif i == 7:  # 200N
                    mprog.partbodyfeatureactivate('SN1_200250')
                    mprog.activatefeature('Hole', 0)
            elif i == 8:  # 250N
                    mprog.partbodyfeatureactivate('SN1_200250')
                    mprog.activatefeature('Hole', 0)
        except:
            print('FRAME15 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME15 Update error')
    elif name == 'FRAME19':
        try:
            if i== 0:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.partbodyfeatureactivate('SN1_25_X')
                mprog.partbodyfeatureactivate('SN1_2535_Y')
            elif i == 1:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.partbodyfeatureactivate('SN1_2535_Y')
                mprog.partbodyfeatureactivate("SN1_3545_AD")
            elif i== 2:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.partbodyfeatureactivate('SN1_45_Y')
                mprog.partbodyfeatureactivate('SN1_45_X')
                mprog.partbodyfeatureactivate("SN1_3545_AD")
            elif i== 3:
                mprog.activatefeature('SN1_60_Body', 0)
            elif i== 4:
                mprog.activatefeature('SN1_80110_Body', 0)
            elif i== 5:
                mprog.activatefeature('SN1_80110_Body', 0)
            elif i== 6:
                mprog.activatefeature('SN1_160_Body', 0)
                mprog.bodydeactivate('FRMAE_SN1_160_Body', 0)
                mprog.bodydeactivate('FRAME_34_Hole_1', 0)
            elif i== 7:
                mprog.activatefeature('SN1_200_Body', 0)
            elif i== 8:
                mprog.activatefeature('SN1_250_Body', 0)
        except:
            print('FRAME19 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME19 Update error')
    elif name == 'FRAME20':
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
            except:
                print('FRAME20 Update error')
    elif name == 'FRAME22':
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.partbodyfeatureactivate('SN1_2545_GH')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.bodydeactivate('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.partbodyfeatureactivate('SN1_2545_GH')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.bodydeactivate('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.partbodyfeatureactivate('SN1_2545_GH')
                mprog.partbodyfeatureactivate('SN1_45_G(45)H(45)')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.bodydeactivate('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_60_I')
                mprog.partbodyfeatureactivate('SN1_60_JK')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.bodydeactivate('FRAME_SN1_60_Body', 0)
                mprog.bodydeactivate('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i== 4:
                mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                mprog.partbodyfeatureactivate('SN1_80250_JK')
                mprog.partbodyfeatureactivate('SN1_80250_GH')
                mprog.partbodyfeatureactivate('SN1_80250_IN')
                mprog.partbodyfeatureactivate('SN1_80250_LM1')
                mprog.partbodyfeatureactivate('SN1_80250_LM2')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.bodydeactivate('FRAME_SN1_80250_Body', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                mprog.partbodyfeatureactivate('SN1_80250_JK')
                mprog.partbodyfeatureactivate('SN1_80250_GH')
                mprog.partbodyfeatureactivate('SN1_80250_IN')
                mprog.partbodyfeatureactivate('SN1_80250_LM1')
                mprog.partbodyfeatureactivate('SN1_80250_LM2')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.bodydeactivate('FRAME_SN1_80250_Body', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                mprog.partbodyfeatureactivate('SN1_80250_JK')
                mprog.partbodyfeatureactivate('SN1_80250_GH')
                mprog.partbodyfeatureactivate('SN1_80250_IN')
                mprog.partbodyfeatureactivate('SN1_80250_LM1')
                mprog.partbodyfeatureactivate('SN1_80250_LM2')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.bodydeactivate('FRAME_SN1_80250_Body', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                mprog.partbodyfeatureactivate('SN1_80250_JK')
                mprog.partbodyfeatureactivate('SN1_80250_GH')
                mprog.partbodyfeatureactivate('SN1_80250_IN')
                mprog.partbodyfeatureactivate('SN1_80250_LM1')
                mprog.partbodyfeatureactivate('SN1_80250_LM2')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.bodydeactivate('FRAME_SN1_80250_Body', 0)
            elif i== 8:
                mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                mprog.partbodyfeatureactivate('SN1_80250_JK')
                mprog.partbodyfeatureactivate('SN1_80250_GH')
                mprog.partbodyfeatureactivate('SN1_80250_IN')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.bodydeactivate('FRAME_SN1_80250_Body', 0)
        except:
            print('FRAME22 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME22 Update error')
    elif name == 'FRAME25':
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
            except:
                print('FRAME25 Update error')
    elif name == 'FRAME28':
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
            except:
                print('FRAME28 Update error')
    elif name == 'FRAME30':
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partdeactivate('SN1_25250_L')
                mprog.partdeactivate('SN1_25250_HIJ')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.activatefeature('CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.bodydeactivate('4_M14通', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partdeactivate('SN1_25250_L')
                mprog.partdeactivate('SN1_25250_HIJ')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.activatefeature('CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.bodydeactivate('4_M14通', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partdeactivate('SN1_25250_L')
                mprog.partdeactivate('SN1_25250_HIJ')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.activatefeature('CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.bodydeactivate('4_M14通', 0)
            elif i== 3:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partdeactivate('SN1_25250_L')
                mprog.partdeactivate('SN1_25250_HIJ')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.activatefeature('CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.bodydeactivate('4_M14通', 0)
            elif i== 4:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partdeactivate('SN1_25250_L')
                mprog.partdeactivate('SN1_25250_HIJ')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.activatefeature('CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.bodydeactivate('4_M14通', 0)
            elif i== 5:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partdeactivate('SN1_25250_L')
                mprog.partdeactivate('SN1_25250_HIJ')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.activatefeature('CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.bodydeactivate('4_M14通', 0)
            elif i== 6:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partdeactivate('SN1_25250_L')
                mprog.partdeactivate('SN1_25250_HIJ')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.activatefeature('CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.bodydeactivate('4_M14通', 0)
            elif i== 7:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partdeactivate('SN1_25250_L')
                mprog.partdeactivate('SN1_25250_HIJ')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.activatefeature('CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.bodydeactivate('4_M14通', 0)
            elif i== 8:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partdeactivate('SN1_25250_L')
                mprog.partdeactivate('SN1_25250_HIJ')
                mprog.partbodyfeatureactivate('SN1_25250_CD')
                mprog.partbodyfeatureactivate('SN1_25250_FG')
                mprog.activatefeature('CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.bodydeactivate('4_M14通', 0)


        except:
            print('FRAME30 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME30 Update error')
    elif name == 'FRAME31':
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('Body')
                mprog.bodydeactivate('FRAME_Body', 0)
            elif i== 1:
                mprog.partbodyfeatureactivate('Body')
                mprog.bodydeactivate('FRAME_Body', 0)
            elif i== 2:
                mprog.partbodyfeatureactivate('Body')
                mprog.bodydeactivate('FRAME_Body', 0)
        except:
            print('FRAME31 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME31 Update error')
    elif name == 'FRAME32':
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
            except:
                print('FRAME32 Update error')
    elif name == "FRAME33":
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.partbodyfeatureactivate('SN1_2545_J')
                mprog.activatefeature('Hole_1', 0)
                mprog.partdeactivate('FRAME_2545_LMN')
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.partbodyfeatureactivate('SN1_2545_J')
                mprog.activatefeature('Hole_1', 0)
                mprog.partdeactivate('FRAME_2545_LMN')
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.partbodyfeatureactivate('SN1_2545_J')
                mprog.activatefeature('Hole_1', 0)
                mprog.partdeactivate('FRAME_2545_LMN')
            elif i== 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.partdeactivate('FRAME_60_LMN')
        except:
            print('FRAME33 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME33 Update success')
            except:
                print('FRAME33 Update error')
    elif name == 'FRAME34':
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.partbodyfeatureactivate('SN1_2545_L')
                mprog.partbodyfeatureactivate('SN1_2535_J')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partdeactivate('FRAME_1_2545_PQR')
            elif i== 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.partbodyfeatureactivate('SN1_2545_L')
                mprog.partbodyfeatureactivate('SN1_2535_J')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partdeactivate('FRAME_1_2545_PQR')
            elif i== 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.partbodyfeatureactivate('SN1_2545_L')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partdeactivate('FRAME_1_2545_PQR')
            elif i== 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60_NM')
                mprog.partbodyfeatureactivate('SN1_60_L')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partdeactivate('FRAME_1_60_PQR')

        except:
            print('FRAME34 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME34 Update success')
            except:
                print('FRAME34 Update error')
    elif name == 'FRAME35':
        try:
            if i== 0:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partdeactivate('SN1_25250_L')
            elif i== 1:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partdeactivate('SN1_25250_L')
            elif i== 2:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partdeactivate('SN1_25250_L')
            elif i== 3:
                mprog.activatefeature('SN1_60_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partdeactivate('SN1_25250_L')
            elif i== 4:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partdeactivate('SN1_25250_L')
            elif i== 5:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partdeactivate('SN1_25250_L')
            elif i== 6:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partdeactivate('SN1_25250_L')
            elif i== 7:
                mprog.activatefeature('SN1_80200_Body', 6)
                mprog.activatefeature('Hole', 2)
                mprog.partdeactivate('SN1_25250_L')
            elif i== 8:
                mprog.activatefeature('SN1_250_Body', 0)
                mprog.activatefeature('Hole', 3)
                mprog.partdeactivate('SN1_25250_L')

        except:
            print('FRAME35 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME35 Update error')
    elif name == 'FRAME36':
        try:
            if i== 0:
                mprog.partbodyfeatureactivate('SN1_2535_Body')
                mprog.partdeactivate('FRAME_SN1_25_Body')
            elif i== 1:
                mprog.partbodyfeatureactivate('SN1_2535_Body')
                mprog.partdeactivate('FRAME_SN1_3560_Body')
            elif i== 2:
                mprog.partbodyfeatureactivate('SN1_4560_Body')
                mprog.partdeactivate('FRAME_SN1_3560_Body')
            elif i== 3:
                mprog.partbodyfeatureactivate('SN1_4560_Body')
                mprog.partdeactivate('FRAME_SN1_3560_Body')
        except:
            print('FRAME36 Parameter activate error')
        finally:
            try:
                mprog.Update()
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
                mprog.partdeactivate('FRAME_Body')
            elif i== 1:
                mprog.partbodyfeatureactivate('Body')
                mprog.partdeactivate('FRAME_Body')
            elif i== 2:
                mprog.partbodyfeatureactivate('Body')
                mprog.partdeactivate('FRAME_Body')
            elif i== 3:
                mprog.partbodyfeatureactivate('Body')
                mprog.partdeactivate('FRAME_Body')
        except:
            print('FRAME37 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME37 Update success')
            except:
                print('FRAME37 Update error')
    elif name == 'FRAME38':
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
            except:
                print('FRAME38 Update error')
    elif name == 'FRAME45':
        try:
            if i == 3:  # 60N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
            elif i == 4:  # 80N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
            elif i == 5:  # 110N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
            elif i == 6:  # 160N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
            elif i == 7:  # 200N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
            elif i == 8:  # 250N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
        except:
            print('FRAME45 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME45 Update error')
    elif name == 'FRAME49':
        try:
            if i== 4:
                mprog.activatefeature('body', 0)
                mprog.partdeactivate('FRAME_SN1_80250_c')
                mprog.bodydeactivate('M12通', 0)
            elif i== 5:
                mprog.activatefeature('body', 0)
                mprog.partdeactivate('FRAME_SN1_80250_c')
                mprog.bodydeactivate('M12通', 0)
            elif i== 6:
                mprog.activatefeature('body', 0)
                mprog.bodydeactivate('M12通', 0)
            elif i== 7:
                mprog.activatefeature('body', 0)
                mprog.partdeactivate('FRAME_SN1_80250_c')
                mprog.bodydeactivate('M12通', 0)
            elif i== 8:
                mprog.activatefeature('body', 0)
                mprog.partdeactivate('FRAME_SN1_80250_c')
                mprog.bodydeactivate('M12通', 0)
        except:
            print('FRAME49 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME49 Update error')
    elif name == 'FRAME50':
        try:
            if i == 4:
                mprog.activatefeature('Body', 0)
                mprog.partdeactivate('FRAME_SN1_c')
                mprog.bodydeactivate('Hole_1', 0)
            elif i == 5:
                mprog.activatefeature('Body', 0)
                mprog.partdeactivate('FRAME_SN1_c')
                mprog.bodydeactivate('Hole_1', 0)
            elif i== 6:
                mprog.activatefeature('Body', 0)
                mprog.partdeactivate('FRAME_SN1_c')
                mprog.bodydeactivate('Hole_1', 0)
            elif i== 7:
                mprog.activatefeature('Body', 0)
                mprog.partdeactivate('FRAME_SN1_c')
                mprog.bodydeactivate('Hole_1', 0)
            elif i == 8:
                mprog.activatefeature('Body', 0)
                mprog.partdeactivate('FRAME_SN1_c')
                mprog.bodydeactivate('Hole_1', 0)
        except:
            print('FRAME50 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME50 Update error')
    elif name == 'FRAME54':
        try:
            if i== 6:
                mprog.partbodyfeatureactivate('Body')
                mprog.partbodyfeatureactivate('Chamfer_E')
            elif i== 7:
                mprog.partbodyfeatureactivate('Body')
                mprog.partbodyfeatureactivate('Chamfer_E')
            elif i == 8:
                mprog.partbodyfeatureactivate('Body')
                mprog.partbodyfeatureactivate('Chamfer_E')
        except:
            print('FRAME54 Parameter activate error')
        finally:
            try:
                mprog.Update()
            except:
                print('FRAME54 Update error')



change_parameter2('FRAME20', 0)



