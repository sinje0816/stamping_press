import main_program as mprog
import STP_input as S_i
import excel_parameter_change as epc
import file_path as fp
#油面計(衝頭)XY跟YZ互換
excel = epc.ExcelOp('組立尺寸', 'STP_Assembly_value')
S_assmebly_par = excel.get_assmebly_sheet_par(5)

mprog.add_offset_assembly('FRAME30.1', 'FRAME2.1', -(2910 - 680) - 50, 'XY plane', 0,
                          40)
#FRAME30的高度alpha要拿掉



#待補齊
#crankshaft的YZ距離為暫定
#PANEL 的YZ距離為暫定
#POINTER 的YZ距離為暫定