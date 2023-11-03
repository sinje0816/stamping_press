import main_program as mprog
import STP_input as S_i
import excel_parameter_change as epc

excel = epc.ExcelOp('組立尺寸', 'STP_Assembly_value')
try:
    S_assmebly_par = excel.get_assmebly_sheet_par(8)
    print('STP_Assembly_value Parameter change success')
except BaseException:
    print('STP_Assembly_value Parameter change error')

mprog.Close_All()
#待修正
#FRAME43.2高度要調整

#待討論
#曲軸合銅座建零件



#待補齊
#crankshaft的YZ距離為暫定
#PANEL 的YZ距離為暫定
#POINTER 的YZ距離為暫定