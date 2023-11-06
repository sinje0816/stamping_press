import main_program as mprog
import STP_input as S_i
import excel_parameter_change as epc
import file_path as fp
stamping_press_type = 8
PORTABLE_STAND_list = ['EWR12S06_PORTABLE_STAND', 'EWR12S06_PORTABLE_STAND' , 'EWR12S06_PORTABLE_STAND', 'EWR12S03_PORTABLE_STAND', 'EWR12S03_PORTABLE_STAND', 'EWR12S03_PORTABLE_STAND', 'EWR12S02_PORTABLE_STAND', 'EWR12S02_PORTABLE_STAND', 'EWR12S02_PORTABLE_STAND']
name = 'PORTABLE_STAND'
if name == 'PORTABLE_STAND':
    mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PORTABLE_STAND_list[stamping_press_type])

mprog.Close_All()
#待修正
#FRAME43.2高度要調整

#待討論
#曲軸合銅座建零件



#待補齊
#crankshaft的YZ距離為暫定
#PANEL 的YZ距離為暫定
#POINTER 的YZ距離為暫定