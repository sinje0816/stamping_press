import main_program as mprog
import STP_input as S_i
import excel_parameter_change as epc
import file_path as fp
#油面計(衝頭)XY跟YZ互換
excel = epc.ExcelOp('組立尺寸', 'STP_Assembly_value')
S_assmebly_par = excel.get_assmebly_sheet_par(5)

print(-(346+32.5-15-150-106-149))
#待修正
#FRAME43.2高度要調整

#待討論
#曲軸合銅座建零件



#待補齊
#crankshaft的YZ距離為暫定
#PANEL 的YZ距離為暫定
#POINTER 的YZ距離為暫定