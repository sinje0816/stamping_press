import main_program as mprog
import STP_input as S_i
import excel_parameter_change as epc
import file_path as fp
mprog.add_offset_assembly('plate.1', 'FRAME22.1', apv['FRAME22']['A(1)']-apv['FRAME22']['A(2)'], 'XY plane', 0, 305)
mprog.add_offset_assembly('plate.1', 'FRAME22.1', -(S_assmebly_par['BH_YZ']), 'YZ plane', 0, 306)
mprog.add_offset_assembly('plate.1', 'FRAME22.1', apv['FRAME30']['E'] / 2, 'ZX plane', 0, 307)

mprog.Close_All()
#待修正
#FRAME43.2高度要調整

#待討論
#曲軸合銅座建零件



#待補齊
#crankshaft的YZ距離為暫定
#PANEL 的YZ距離為暫定
#POINTER 的YZ距離為暫定