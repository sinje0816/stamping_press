import main_program as mprog
import STP_input as S_i
import excel_parameter_change as epc
import file_path as fp

mprog.add_offset_assembly(S_i.OIL_LEVEL_GAUGE_list_SLIDE[8] + '.1',
                          S_i.SLIDE_list_normal[8] + '.1', 100, 'XY plane', 0,
                          317)
#油面計(衝頭)XY跟YZ互換

mprog.Close_All()
#待修正
#FRAME43.2高度要調整

#待討論
#曲軸合銅座建零件



#待補齊
#crankshaft的YZ距離為暫定
#PANEL 的YZ距離為暫定
#POINTER 的YZ距離為暫定