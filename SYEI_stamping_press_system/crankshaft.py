import win32com.client as win32
import main_program as mprog, main_GUI
import excel_parameter_change as epc

def change_parameter(name, file_sheet_name, i):
    if name == 'crankshaft_S':
        excel = epc.ExcelOp('尺寸整理表_曲軸', 'crankshaft_S')
        try:
            crankshaft_S_parameter_name, crankshaft_S_parameter_value = excel.get_sheet_par(file_sheet_name, i)
            print('%s Parameter change success' % file_sheet_name)
        except:
            print('%s Parameter change error' % file_sheet_name)

        try:
            if i == 0:      # 302CC7S
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外1', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_6M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 1:    # 322CC7S
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('BALLr外1', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_6M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 2:    # 342CC7S
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外2', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_6M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 3:    # 372CC7S
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外2', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 4:    # 395CC7S
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.partbodyfeatrueactivate('CALLr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
            elif i == 5:    # 412CC7S
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外2', 0)
                mprog.partbodyfeatrueactivate('BALLr內2', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.partbodyfeatrueactivate('Cex(SN1_110_S)', 0)
                mprog.activatefeatrue('Cex孔(SN1_110_S)', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
                mprog.activatefeatrue('Dex(SN1_110_S)', 0)
            elif i == 6:    # 435CC7S
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
            elif i == 7:    # 455CC7S
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
            elif i == 8:    # 475CC7S
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外2', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
        except:
            print('crankshaft_S Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('crankshaft_S Update success')
            except:
                print('crankshaft_S Update error')

    elif name == 'crankshaft_H':
        excel = epc.ExcelOp('尺寸整理表_曲軸', 'crankshaft_H')
        try:
            crankshaft_H_parameter_name, crankshaft_H_parameter_value = excel.get_sheet_par(file_sheet_name, i)
            print('%s Parameter change success' % file_sheet_name)
        except:
            print('%s Parameter change error' % file_sheet_name)

        try:
            if i == 0:      # 302CC7H
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外1', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_6M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 1:    # 322CC7H
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('BALLr外1', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_6M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 2:    # 342CC7H
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外2', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_6M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 3:    # 372CC7H
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外1', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 4:    # 395CC7H
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.partbodyfeatrueactivate('CALLr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
            elif i == 5:    # 415CC7H
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外2', 0)
                mprog.partbodyfeatrueactivate('BALLr內2', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
            elif i == 6:    # 435CC7H
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
            elif i == 7:    # 455CC7H
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
            elif i == 8:    # 475CC7H
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外2', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
        except:
            print('crankshaft_H Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('crankshaft_H Update success')
            except:
                print('crankshaft_H Update error')

    elif name == 'crankshaft_P':
        excel = epc.ExcelOp('尺寸整理表_曲軸', 'crankshaft_P')
        try:
            crankshaft_P_parameter_name, crankshaft_P_parameter_value = excel.get_sheet_par(file_sheet_name, i)
            print('%s Parameter change success' % file_sheet_name)
        except:
            print('%s Parameter change error' % file_sheet_name)

        try:
            if i == 0:      # 302CC7P
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外1', 0)
                mprog.partbodyfeatrueactivate('BALLr內2', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_6M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 1:    # 322CC7P
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('BALLr外1', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_6M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 2:    # 342CC7P
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外2', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_6M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 3:    # 372CC7P
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Aex(SN1_60_P)', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外1', 0)
                mprog.partbodyfeatrueactivate('BALLr內1', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_1M', 0)
            elif i == 4:    # 395CC7P
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.partbodyfeatrueactivate('CALLr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
            # elif i == 5:    # 415CC7P
            #     mprog.partbodyfeatrueactivate('Bex', 0)
            #     mprog.partbodyfeatrueactivate('C5', 0)
            #     mprog.partbodyfeatrueactivate('AALLr', 0)
            #     mprog.partbodyfeatrueactivate('BALLr外2', 0)
            #     mprog.partbodyfeatrueactivate('BALLr內2', 0)
            #     mprog.partbodyfeatrueactivate('CALLr1', 0)
            #     mprog.partbodyfeatrueactivate('Cr2', 0)
            #     mprog.activatefeatrue('D1吊掛孔', 0)
            #     mprog.activatefeatrue('D2_4M', 0)
            #     mprog.activatefeatrue('D3_4M', 0)
            #     mprog.activatefeatrue('D4_4M', 0)
            elif i == 6:    # 435CC7P
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
            elif i == 7:    # 455CC7P
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('CALLr1', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
            elif i == 8:    # 475CC7P
                mprog.partbodyfeatrueactivate('Bex', 0)
                mprog.partbodyfeatrueactivate('C5', 0)
                mprog.partbodyfeatrueactivate('Ar1', 0)
                mprog.partbodyfeatrueactivate('AALLr', 0)
                mprog.partbodyfeatrueactivate('BALLr外2', 0)
                mprog.partbodyfeatrueactivate('Cr2', 0)
                mprog.activatefeatrue('D1吊掛孔', 0)
                mprog.activatefeatrue('D2_4M', 0)
                mprog.activatefeatrue('D3_4M', 0)
                mprog.activatefeatrue('D4_4M', 0)
        except:
            print('crankshaft_P Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('crankshaft_P Update success')
            except:
                print('crankshaft_P Update error')

change_parameter('crankshaft_S', 'crankshaft', 0)