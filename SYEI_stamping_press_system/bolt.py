import main_program as mprog
import excel_parameter_change as epc

def bolt_parameter_design_part(i):
    excel = epc.ExcelOp('螺栓尺寸表', '夾頭螺栓')  #讀取excel檔案(括號內填入(分頁名稱)
    try:
        parameter_name, paramter_value = excel.get_sheet_par('bolt', i)
        print('bolt parameter change success')
    except:
        print('bolt parameter change error')
    try:
        mprog.partbodyfeatureactivate('Pad.1')
        mprog.partbodyfeatureactivate('Thread.1')
        mprog.partbodyfeatureactivate('Thread.2')
        mprog.partbodyfeatureactivate('Chamfer.1')
        mprog.partbodyfeatureactivate('Pocket.1')
        mprog.partbodyfeatureactivate('Shaft.1')
        print('bolt machining success')
    except:
        print('bolt machining error')
    finally:
        try:
            mprog.update()
            print('bolt update success')
        except:
            print('bolt update error')

bolt_parameter_design_part(8)