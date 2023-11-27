import main_program as mprog
import file_path as fp

PANEL_list = ['32H8302_PANEL', '322CC10_PANEL', '37H8302_PANEL', '37H8302_PANEL', '41H8302_PANEL', '552CC10_PANEL', '45H8302_PANEL', '45H8302_PANEL', '47H0001_PANEL']
CON_ROD_list = ["302B04_CON_ROD", "322B04_CON_ROD", "342B04S01_CON_ROD", "372B04S01_CON_ROD", "392B04S01_CON_ROD", "412B04S03_CON_ROD", "432B04S01_CON_ROD", "452B04S02_CON_ROD", "472B04S02_CON_ROD"]
CON_ROD_BASE_list = ["302B03_CON_ROD_BASE", "322B03_CON_ROD_BASE", "342B03_CON_ROD_BASE", "372B03_CON_ROD_BASE", "392B03S01_CON_ROD_BASE", "412B03S01P10_CON_ROD_BASE", "432B03S01_CON_ROD_BASE", "452B03S02_CON_ROD_BASE", "47B03S02_CON_ROD_BASE"]
CON_ROD_CAP_list = ["302B02_CON_ROD_CAP", "322B02_CON_ROD_CAP", "342B02_CON_ROD_CAP", "372B02_CON_ROD_CAP", "392B02P10_CON_ROD_CAP", "412B02S01P10_CON_ROD_CAP", "432B02S01_CON_ROD_CAP", "452B02S01_CON_ROD_CAP", "47B02S01P10_CON_ROD_CAP"]
INVERTERBRACKET_list = ['EWR12S01_INVERTERBRACKET', 'EWR12S01_INVERTERBRACKET', 'EWR12S01_INVERTERBRACKET', 'EWR14S01_INVERTERBRACKET']
POINTER_list = ['32H8304_POINTER', '322CC11_POINTER', '37H8304_POINTER', '37H8304_POINTER', "41H8304_POINTER" ,"41H8304_POINTER", "45H8304_POINTER", "45H8304_POINTER", "45H8304_POINTER"]
COVER_list = ['302A03S01_COVER', '322A03S01_COVER', '342A02S01_COVER', '372A03S01_COVER', '392A03_COVER', '412A03_COVER', '432A03_COVER', '452A03_COVER', '47A0002_COVER']
PLUG_list = '32A0016-PLUG'
feeding_shaft_cover_list = ['392K01_feeding_shaft_cover', '392K01_feeding_shaft_cover', '392K01_feeding_shaft_cover', '392K01_feeding_shaft_cover']
OIL_LEVEL_GAUGE_list_FRAME = 'OGASKL060_OIL_LEVEL_GAUGE'
OIL_LEVEL_GAUGE_list_CON_ROD = ['','','', 'OGASKL150W_OIL_LEVEL_GAUGE', 'OGASKL150W_OIL_LEVEL_GAUGE', 'OGASKL150W_OIL_LEVEL_GAUGE', 'OGASKL170W_OIL_LEVEL_GAUGE', 'OGASKL254W_OIL_LEVEL_GAUGE', 'OGASKL254W_OIL_LEVEL_GAUGE', 'OGASKL254W_OIL_LEVEL_GAUGE']
OIL_LEVEL_GAUGE_list_SLIDE = ['OGC13901HA_OIL_LEVEL_GAUGE', 'OGC13901HA_OIL_LEVEL_GAUGE', 'OGC14001HA_OIL_LEVEL_GAUGE', 'OGC14001HA_OIL_LEVEL_GAUGE', 'OGC14001HA_OIL_LEVEL_GAUGE', 'OGC14001HA_OIL_LEVEL_GAUGE', 'OGC14001HA_OIL_LEVEL_GAUGE', 'OGC14001HA_OIL_LEVEL_GAUGE', 'OGC14001HA_OIL_LEVEL_GAUGE']
slide_gib_list_right = ['302B22S02_slide_gib', '322B22S02_slide_gib', '342B22S02_slide_gib', '372B22S02_slide_gib', '392B22S02_slide_gib', '412B22S02_slide_gib', '432B22S02_slide_gib', '452B22S02_slide_gib', '472B22S03_slide_gib']
slide_gib_list_left = ['302B21S02_slide_gib', '322B21S02_slide_gib', '342B21S02_slide_gib', '372B21S02_slide_gib', '392B21S02_slide_gib', '412B21S02_slide_gib', '432B21S02_slide_gib', '452B21S02_slide_gib', '472B21S03_slide_gib']
CONTROL_PANEL_list_normal = ['01A0541785Q_CONTROL_PANEL', '01A0541785Q_CONTROL_PANEL', '01A0541785Q_CONTROL_PANEL', '01A0541785Q_CONTROL_PANEL', '01A0541785Q_CONTROL_PANEL', '01A0541785Q_CONTROL_PANEL', '01A0541785Q_CONTROL_PANEL', '01A0541795Q_CONTROL_PANEL', '01A0541795Q_CONTROL_PANEL']
CONTROL_PANEL_list_special = '01A0543735P_CONTROL_PANEL'
PANEL_BOX_list_normal = ['01A0532335Q_PANEL_BOX', '01A0532335Q_PANEL_BOX', '01A0532335Q_PANEL_BOX', 'SLSR01S01_PANEL_BOX', 'SLSR01S01_PANEL_BOX', 'SLSR01S01_PANEL_BOX', 'SLR05_PANEL_BOX', '94452S028_PANEL_BOX', '94452S028_PANEL_BOX']
PANEL_BOX_list_normal_SLIDE = ['SLSR01S01_PANEL_BOX', 'SLSR01S01_PANEL_BOX', 'SLSR01S01_PANEL_BOX', 'SLSR01S01_PANEL_BOX', 'SLSR01S01_PANEL_BOX', 'SLSR01S01_PANEL_BOX']
PANEL_BOX_list_normal_slide_gib = ['', '', 'SLR05_PANEL_BOX', 'SLR05_PANEL_BOX', 'SLR05_PANEL_BOX', 'SLR05_PANEL_BOX', 'SLR05_PANEL_BOX', '94452S028_PANEL_BOX', '94452S028_PANEL_BOX']
PANEL_BOX_list_normal_Throat = ['', '', '', '', '', 'SLSR01S01_PANEL_BOX', 'SLSR01S01_PANEL_BOX', '01A0532505Q_PANEL_BOX', '01A0532515Q_PANEL_BOX']
PANEL_BOX_list_special = ['MZW8740_PANEL_BOX', 'MZW8740_PANEL_BOX', 'MZW8740_PANEL_BOX', '01A0535755P_PANEL_BOX', '01A0535755P_PANEL_BOX', '01A0535755P_PANEL_BOX', '53R04A_PANEL_BOX', '452R23_PANEL_BOX', '472R03_PANEL_BOX']
PANEL_BOX_list_special_SLIDE = ['01A0535755P_PANEL_BOX', '01A0535755P_PANEL_BOX', '01A0535755P_PANEL_BOX', '01A0535755P_PANEL_BOX', '01A0535755P_PANEL_BOX', '01A0535755P_PANEL_BOX']
PANEL_BOX_list_special_slide_gib = ['', '', '53R04A_PANEL_BOX', '53R04A_PANEL_BOX', '53R04A_PANEL_BOX', '53R04A_PANEL_BOX', '53R04A_PANEL_BOX', '53R04A_PANEL_BOX', '53R04A_PANEL_BOX']
PANEL_BOX_list_special_Throat = ['', '', '', '', '','01A0535755P_PANEL_BOX', '01A0535755P_PANEL_BOX', 'EWR16S01_PANEL_BOX', 'EWR16S01_PANEL_BOX']
PANEL_BOX_BRACKET_list_normal = ['', '', '', '', '392R24S02_PANEL_BOX_BRACKET', '412R24S03_PANEL_BOX_BRACKET', '432R24_PANEL_BOX_BRACKET', '45R0002_PANEL_BOX_BRACKET', '47R0002_PANEL_BOX_BRACKET']
PANEL_BOX_BRACKET_list_normal_SILDE = ['302R24_PANEL_BOX_BRACKET', '322R24_PANEL_BOX_BRACKET', '342R24_PANEL_BOX_BRACKET', '372R24_PANEL_BOX_BRACKET', '392R24_PANEL_BOX_BRACKET', '412R24_PANEL_BOX_BRACKET', '432R24_PANEL_BOX_BRACKET', '45R0002_PANEL_BOX_BRACKET', '47R0002_PANEL_BOX_BRACKET']
PANEL_BOX_BRACKET_list_normal_slide_gib = ['', '', '01A0553735P_PANEL_BOX_BRACKET', '01A0553745P_PANEL_BOX_BRACKET', '01A0553755P_PANEL_BOX_BRACKET', '412R68S01_PANEL_BOX_BRACKET', '432R68S01_PANEL_BOX_BRACKET', '452R68S01_PANEL_BOX_BRACKET', '472R68S01_PANEL_BOX_BRACKET']
PANEL_BOX_BRACKET_list_normal_Throat = ['', '', '', '', '', '412R24_PANEL_BOX_BRACKET', '432R24_PANEL_BOX_BRACKET', '452R02S01_PANEL_BOX_BRACKET', '472R02S01_PANEL_BOX_BRACKET']
ELECTRIC_BOX_list_normal = ['01A0534405P_ELECTRIC_BOX', '01A0534405P_ELECTRIC_BOX', '01A0534405P_ELECTRIC_BOX', '01A0534395P_ELECTRIC_BOX', '01A0534375P_ELECTRIC_BOX', '01A0534375P_ELECTRIC_BOX', '01A0534375P_ELECTRIC_BOX', '01A0534375P_ELECTRIC_BOX', '01A0534375P_ELECTRIC_BOX']
GUARD_FLYWHEEL_list_normal = ['302N06S04_GUARD_FLYWHEEL', '32N6603_GUARD_FLYWHEEL', '342N06S03_GUARD_FLYWHEEL', '372N06S01_GUARD_FLYWHEEL', '392N06S06_GUARD_FLYWHEEL', '412N06S08_GUARD_FLYWHEEL', '432N06S05_GUARD_FLYWHEEL', '452N06S04_GUARD_FLYWHEEL', '472N06S04_GUARD_FLYWHEEL']
NAME_PLATE_list_normal = ['302N52_NAME_PLATE', '322N52_NAME_PLATE', '342N52_NAME_PLATE', '372N52_NAME_PLATE', '392N52_NAME_PLATE', '412N52_NAME_PLATE', '432N52_NAME_PLATE', '452N52_NAME_PLATE', '472N52_NAME_PLATE']
TRADEMARK_NAMEPLATE_list_normal = ['MZD8276_TRADEMARK_NAMEPLATE', 'MZD8276_TRADEMARK_NAMEPLATE', '34N19_TRADEMARK_NAMEPLATE', '34N19_TRADEMARK_NAMEPLATE', 'MZD8275_TRADEMARK_NAMEPLATE', 'MZD8275_TRADEMARK_NAMEPLATE', 'MZD8274_TRADEMARK_NAMEPLATE', 'MZD8274_TRADEMARK_NAMEPLATE', '47N46_TRADEMARK_NAMEPLATE']
OPERATION_BOX_list = '01A061186RP_OPERATION_BOX'
PORTABLE_STAND_list = ['EWR12S06_PORTABLE_STAND', 'EWR12S06_PORTABLE_STAND' , 'EWR12S06_PORTABLE_STAND', 'EWR12S03_PORTABLE_STAND', 'EWR12S03_PORTABLE_STAND', 'EWR12S03_PORTABLE_STAND', 'EWR12S02_PORTABLE_STAND', 'EWR12S02_PORTABLE_STAND', 'EWR12S02_PORTABLE_STAND']
OPERATION_BOX_list_normal = ['01A061186RP_OPERATION_BOX', '01A061186RP_OPERATION_BOX', '01A061186RP_OPERATION_BOX', '01A061186RP_OPERATION_BOX', '01A061186RP_OPERATION_BOX', '01A061186RP_OPERATION_BOX', '01A061186RP_OPERATION_BOX', '01A061186RP_OPERATION_BOX', '01A061186RP_OPERATION_BOX']
BEARING_HOUSING_list = ['302CC4S02_BEARING_HOUSING', '322CC4S02_BEARING_HOUSING', '342CC4S02_BEARING_HOUSING', '372CC4S02_BEARING_HOUSING', '395CC4S02_BEARING_HOUSING', '412CC4S02_BEARING_HOUSING', '432CC4S02_BEARING_HOUSING', '455CC4S02_BEARING_HOUSING', '475CC4S02_BEARING_HOUSING']
SLIDE_list_normal = ['', '', '', 'punch_3_normal', 'punch_4_normal', 'punch_5_normal', 'punch_6_normal', 'punch_7_normal', 'punch_8_normal']
BALANCER_list = ['FAC063160W_BALANCER', 'FAC063170W_BALANCER', 'FAC070205W_BALANCER', 'FAC080235W_BALANCER', 'FAC100280W_BALANCER', 'FAC120320W_BALANCER', 'FAC140360W_BALANCER', 'FAC180390WA_BALANCER', 'FAC200450WS01_BALANCER']
MOTOR_list = ['5HP_3.7kW_MOTOR', '5HP_3.7kW_MOTOR', '5HP_3.7kW_MOTOR', '7.5HP_5.5kW_MOTOR', '10HP_7.5kW_MOTOR', '10HP_7.5kW_MOTOR', '15HP_11kW_MOTOR', '20HP_15kW_MOTOR', '25HP_18.5kW_MOTOR']
MOTOR_BRACKET_S_list = ['302M01S02_MOTOR_BASE', '322M01S02_MOTOR_BASE', '34M4701_MOTOR_BASE', '392M01_MOTOR_BASE', '392M01S03_MOTOR_BASE', '41M4701_MOTOR_BASE', '34M4701_MOTOR_BASE', '452M01_MOTOR_BASE', '472M01_MOTOR_BASE']
MOTOR_BRACKET_HP_list = ['302M01S02_MOTOR_BASE', '322M01S02_MOTOR_BASE', '34M4701_MOTOR_BASE', '392M01_MOTOR_BASE', '392M01S03_MOTOR_BASE', '412M01S03_MOTOR_BASE', '432M01S04_MOTOR_BASE', '94452S053_MOTOR_BASE', '472M01_MOTOR_BASE']
# 穎益
GUM_E_list = ['L7_18-20_ANTI_VIBRATION_GUM', 'L7_18-20_ANTI_VIBRATION_GUM', 'L9_21_24_ANTI_VIBRATION_GUM', 'L9_21_24_ANTI_VIBRATION_GUM', 'L19_25_24_ANTI_VIBRATION_GUM', 'L29_29_24_ANTI_VIBRATION_GUM', 'L39_32_30_ANTI_VIBRATION_GUM', 'L69_38_39_ANTI_VIBRATION_GUM', 'L89_45_39_ANTI_VIBRATION_GUM']
# UNISORB
GUM_U_list = ['PM_61_M16_N90_ANTI_VIBRATION_GUM', 'PM_61_M16_N90_ANTI_VIBRATION_GUM', 'PM_81_M24_N75_ANTI_VIBRATION_GUM', 'PM_81_M24_N90_ANTI_VIBRATION_GUM', 'PM_81_M24_N90_ANTI_VIBRATION_GUM', 'PM_101_M30(150)_TF24_ANTI_VIBRATION_GUM', 'PM_121_M36(180)_TF25_ANTI_VIBRATION_GUM', 'PM_121_M36(180)_TF26_ANTI_VIBRATION_GUM', 'PM_161_M36(180)_TF2631_ANTI_VIBRATION_GUM']
# 商定
GUM_S_list = ['KM_145_ANTI_VIBRATION_GUM', 'KM_145_ANTI_VIBRATION_GUM', 'KM_145_ANTI_VIBRATION_GUM', 'KM_200_ANTI_VIBRATION_GUM', 'KM_200_ANTI_VIBRATION_GUM', 'KM_200_ANTI_VIBRATION_GUM', 'KM_250_ANTI_VIBRATION_GUM', 'KM_250_ANTI_VIBRATION_GUM', 'KM_250_ANTI_VIBRATION_GUM']
#東元
INVERTER_D_S_list =['', '', '', '', '', 'SAV0102B1T_INVERTER', '', '', 'SAV0252B1T_INVERTER']
INVERTER_D_HP_list =['', '', '', '', '', 'SAV0152B1T_INVERTER', '', '', 'SAV0252B1T_INVERTER']
#台達
INVERTER_T_S_list =['', '', '', '', '', 'SAV0102BUS_INVERTER', '', '', 'SAV0252BUS_INVERTER']
INVERTER_T_HP_list =['', '', '', '', '', 'SAV0152BUS_INVERTER', '', '', 'SAV0252BUS_INVERTER']
def STP(name, stamping_press_type, machining, travel_type, GUM, INVERTER):
    if name == 'PANEL':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_list[stamping_press_type])
        mprog.save_file_stp(machining, PANEL_list[stamping_press_type])
        mprog.save_stpfile_part(machining, PANEL_list[stamping_press_type])
        mprog.close_file(PANEL_list[stamping_press_type])
    elif name == 'CON_ROD':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), CON_ROD_list[stamping_press_type])
        mprog.save_file_stp(machining, CON_ROD_list[stamping_press_type])
        mprog.save_stpfile_part(machining, CON_ROD_list[stamping_press_type])
        mprog.close_file(CON_ROD_list[stamping_press_type])
    elif name == 'CON_ROD_BASE':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), CON_ROD_BASE_list[stamping_press_type])
        mprog.save_file_stp(machining, CON_ROD_BASE_list[stamping_press_type])
        mprog.save_stpfile_part(machining, CON_ROD_BASE_list[stamping_press_type])
        mprog.close_file(CON_ROD_BASE_list[stamping_press_type])
    elif name == 'CON_ROD_CAP':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), CON_ROD_CAP_list[stamping_press_type])
        mprog.save_file_stp(machining, CON_ROD_CAP_list[stamping_press_type])
        mprog.save_stpfile_part(machining, CON_ROD_CAP_list[stamping_press_type])
        mprog.close_file(CON_ROD_CAP_list[stamping_press_type])
    elif name == 'INVERTERBRACKET':
        try:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), INVERTERBRACKET_list[stamping_press_type])
            mprog.save_file_stp(machining, INVERTERBRACKET_list[stamping_press_type])
            mprog.save_stpfile_part(machining, INVERTERBRACKET_list[stamping_press_type])
            mprog.close_file(INVERTERBRACKET_list[stamping_press_type])
        except:
            pass
    elif name == 'POINTER':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), POINTER_list[stamping_press_type])
        mprog.save_file_stp(machining, POINTER_list[stamping_press_type])
        mprog.save_stpfile_part(machining, POINTER_list[stamping_press_type])
        mprog.close_file(POINTER_list[stamping_press_type])
    elif name == 'COVER':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), COVER_list[stamping_press_type])
        mprog.save_file_stp(machining, COVER_list[stamping_press_type])
        mprog.save_stpfile_part(machining, COVER_list[stamping_press_type])
        mprog.close_file(COVER_list[stamping_press_type])
    elif name == 'PLUG':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PLUG_list)
        mprog.save_file_stp(machining, PLUG_list)
        mprog.save_stpfile_part(machining, PLUG_list)
        mprog.close_file(PLUG_list)
    elif name == 'feeding_shaft_cover':
        try:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), feeding_shaft_cover_list[stamping_press_type])
            mprog.save_file_stp(machining, feeding_shaft_cover_list[stamping_press_type])
            mprog.save_stpfile_part(machining, feeding_shaft_cover_list[stamping_press_type])
            mprog.close_file(feeding_shaft_cover_list[stamping_press_type])
        except:
            pass
    elif name == 'OIL_LEVEL_GAUGE':
        if stamping_press_type <= 2:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), 'OGASKL060_OIL_LEVEL_GAUGE')
            mprog.save_file_stp(machining, 'OGASKL060_OIL_LEVEL_GAUGE')
            mprog.save_stpfile_part(machining, 'OGASKL060_OIL_LEVEL_GAUGE')
            mprog.close_file('OGASKL060_OIL_LEVEL_GAUGE')
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), OIL_LEVEL_GAUGE_list_SLIDE[stamping_press_type])
            mprog.save_file_stp(machining, OIL_LEVEL_GAUGE_list_SLIDE[stamping_press_type])
            mprog.save_stpfile_part(machining, OIL_LEVEL_GAUGE_list_SLIDE[stamping_press_type])
            mprog.close_file(OIL_LEVEL_GAUGE_list_SLIDE[stamping_press_type])
        else:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), 'OGASKL060_OIL_LEVEL_GAUGE')
            mprog.save_file_stp(machining, 'OGASKL060_OIL_LEVEL_GAUGE')
            mprog.save_stpfile_part(machining, 'OGASKL060_OIL_LEVEL_GAUGE')
            mprog.close_file('OGASKL060_OIL_LEVEL_GAUGE')
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), OIL_LEVEL_GAUGE_list_SLIDE[stamping_press_type])
            mprog.save_file_stp(machining, OIL_LEVEL_GAUGE_list_SLIDE[stamping_press_type])
            mprog.save_stpfile_part(machining, OIL_LEVEL_GAUGE_list_SLIDE[stamping_press_type])
            mprog.close_file(OIL_LEVEL_GAUGE_list_SLIDE[stamping_press_type])
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), OIL_LEVEL_GAUGE_list_CON_ROD[stamping_press_type])
            mprog.save_file_stp(machining, OIL_LEVEL_GAUGE_list_CON_ROD[stamping_press_type])
            mprog.save_stpfile_part(machining, OIL_LEVEL_GAUGE_list_CON_ROD[stamping_press_type])
            mprog.close_file(OIL_LEVEL_GAUGE_list_CON_ROD[stamping_press_type])
    elif name == 'slide_gib':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), slide_gib_list_right[stamping_press_type])
        mprog.save_file_stp(machining, slide_gib_list_right[stamping_press_type])
        mprog.save_stpfile_part(machining, slide_gib_list_right[stamping_press_type])
        mprog.close_file(slide_gib_list_right[stamping_press_type])
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), slide_gib_list_left[stamping_press_type])
        mprog.save_file_stp(machining, slide_gib_list_left[stamping_press_type])
        mprog.save_stpfile_part(machining, slide_gib_list_left[stamping_press_type])
        mprog.close_file(slide_gib_list_left[stamping_press_type])
    elif name == 'ELECTRIC_BOX_PLATE':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), 'EWR60S01_ELECTRIC_BOX_PLATE')
        mprog.save_file_stp(machining, 'EWR60S01_ELECTRIC_BOX_PLATE')
        mprog.save_stpfile_part(machining, 'EWR60S01_ELECTRIC_BOX_PLATE')
        mprog.close_file('EWR60S01_ELECTRIC_BOX_PLATE')
    elif name == 'MOUNT_FILTER':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), 'EGPSSGD1000IS_MOUNT_FILTER')
        mprog.save_file_stp(machining, 'EGPSSGD1000IS_MOUNT_FILTER')
        mprog.save_stpfile_part(machining, 'EGPSSGD1000IS_MOUNT_FILTER')
        mprog.close_file('EGPSSGD1000IS_MOUNT_FILTER')
    elif name == 'CONTROL_PANEL':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), CONTROL_PANEL_list_normal[stamping_press_type])
        mprog.save_file_stp(machining, CONTROL_PANEL_list_normal[stamping_press_type])
        mprog.save_stpfile_part(machining, CONTROL_PANEL_list_normal[stamping_press_type])
        mprog.close_file(CONTROL_PANEL_list_normal[stamping_press_type])
    elif name == 'PANEL_BOX':
        if stamping_press_type == 0:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 1:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 2:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 3:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 4:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 5:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 6:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 7:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 8:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_list_normal[stamping_press_type])
    elif name == 'PANEL_BOX_BRACKET':
        if stamping_press_type == 0:
            pass
        elif stamping_press_type == 1:
            pass
        elif stamping_press_type == 2:
            pass
        elif stamping_press_type == 3:
            pass
        elif stamping_press_type == 4:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_BRACKET_list_normal[stamping_press_type])
        elif stamping_press_type == 5:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_BRACKET_list_normal[stamping_press_type])
        elif stamping_press_type == 6:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_BRACKET_list_normal[stamping_press_type])
        elif stamping_press_type == 7:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_BRACKET_list_normal[stamping_press_type])
        elif stamping_press_type == 8:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.save_file_stp(machining, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.save_stpfile_part(machining, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
            mprog.close_file(PANEL_BOX_BRACKET_list_normal[stamping_press_type])
    elif name == 'ELECTRIC_BOX':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), ELECTRIC_BOX_list_normal[stamping_press_type])
        mprog.save_file_stp(machining, ELECTRIC_BOX_list_normal[stamping_press_type])
        mprog.save_stpfile_part(machining, ELECTRIC_BOX_list_normal[stamping_press_type])
        mprog.close_file(ELECTRIC_BOX_list_normal[stamping_press_type])
    elif name == 'GUARD_FLYWHEEL':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), GUARD_FLYWHEEL_list_normal[stamping_press_type])
        mprog.save_file_stp(machining, GUARD_FLYWHEEL_list_normal[stamping_press_type])
        mprog.save_stpfile_part(machining, GUARD_FLYWHEEL_list_normal[stamping_press_type])
        mprog.close_file(GUARD_FLYWHEEL_list_normal[stamping_press_type])
    elif name == 'NAME_PLATE':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), NAME_PLATE_list_normal[stamping_press_type])
        mprog.save_file_stp(machining, NAME_PLATE_list_normal[stamping_press_type])
        mprog.save_stpfile_part(machining, NAME_PLATE_list_normal[stamping_press_type])
        mprog.close_file(NAME_PLATE_list_normal[stamping_press_type])
    elif name == 'TRADEMARK_NAMEPLATE':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), TRADEMARK_NAMEPLATE_list_normal[stamping_press_type])
        mprog.save_file_stp(machining, TRADEMARK_NAMEPLATE_list_normal[stamping_press_type])
        mprog.save_stpfile_part(machining, TRADEMARK_NAMEPLATE_list_normal[stamping_press_type])
        mprog.close_file(TRADEMARK_NAMEPLATE_list_normal[stamping_press_type])
    elif name == 'OPERATION_BOX':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), '01A061186RP_OPERATION_BOX')
        mprog.save_file_stp(machining, '01A061186RP_OPERATION_BOX')
        mprog.save_stpfile_part(machining, '01A061186RP_OPERATION_BOX')
        mprog.close_file('01A061186RP_OPERATION_BOX')
    elif name == 'PORTABLE_STAND':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), PORTABLE_STAND_list[stamping_press_type])
        mprog.save_file_stp(machining, PORTABLE_STAND_list[stamping_press_type])
        mprog.save_stpfile_part(machining, PORTABLE_STAND_list[stamping_press_type])
        mprog.close_file(PORTABLE_STAND_list[stamping_press_type])
    elif name == 'OPERATION_BOX':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), OPERATION_BOX_list_normal[stamping_press_type])
        mprog.save_file_stp(machining, OPERATION_BOX_list_normal[stamping_press_type])
        mprog.save_stpfile_part(machining, OPERATION_BOX_list_normal[stamping_press_type])
        mprog.close_file(OPERATION_BOX_list_normal[stamping_press_type])
    elif name == 'BEARING_HOUSING':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), BEARING_HOUSING_list[stamping_press_type])
        mprog.save_file_stp(machining, BEARING_HOUSING_list[stamping_press_type])
        mprog.save_stpfile_part(machining, BEARING_HOUSING_list[stamping_press_type])
        mprog.close_file(BEARING_HOUSING_list[stamping_press_type])
    elif name == 'SLIDE':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), SLIDE_list_normal[stamping_press_type])
        mprog.save_file_stp(machining, SLIDE_list_normal[stamping_press_type])
        mprog.save_stpfile_part(machining, SLIDE_list_normal[stamping_press_type])
        mprog.close_file(SLIDE_list_normal[stamping_press_type])
    elif name == 'BALANCER':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), BALANCER_list[stamping_press_type])
        mprog.save_file_stp(machining, BALANCER_list[stamping_press_type])
        mprog.save_stpfile_part(machining, BALANCER_list[stamping_press_type])
        mprog.close_file(BALANCER_list[stamping_press_type])
    elif name == 'MOTOR':
        mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), MOTOR_list[stamping_press_type])
        mprog.save_file_stp(machining, MOTOR_list[stamping_press_type])
        mprog.save_stpfile_part(machining, MOTOR_list[stamping_press_type])
        mprog.close_file(MOTOR_list[stamping_press_type])
    elif name == 'MOTOR_BRACKET':
        if travel_type == 1:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), MOTOR_BRACKET_S_list[stamping_press_type])
            mprog.save_file_stp(machining, MOTOR_BRACKET_S_list[stamping_press_type])
            mprog.save_stpfile_part(machining, MOTOR_BRACKET_S_list[stamping_press_type])
            mprog.close_file(MOTOR_BRACKET_S_list[stamping_press_type])
        else:
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), MOTOR_BRACKET_HP_list[stamping_press_type])
            mprog.save_file_stp(machining, MOTOR_BRACKET_HP_list[stamping_press_type])
            mprog.save_stpfile_part(machining, MOTOR_BRACKET_HP_list[stamping_press_type])
            mprog.close_file(MOTOR_BRACKET_HP_list[stamping_press_type])
    elif name == 'WIRE_CASING':
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), '43N6603_WIRE_CASING')
            mprog.save_file_stp(machining, '43N6603_WIRE_CASING')
            mprog.save_stpfile_part(machining, '43N6603_WIRE_CASING')
            mprog.close_file('43N6603_WIRE_CASING')
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), '43N6604_WIRE_CASING')
            mprog.save_file_stp(machining, '43N6604_WIRE_CASING')
            mprog.save_stpfile_part(machining, '43N6604_WIRE_CASING')
            mprog.close_file('43N6604_WIRE_CASING')
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), 'MZW8702_WIRE_CASING')
            mprog.save_file_stp(machining, 'MZW8702_WIRE_CASING')
            mprog.save_stpfile_part(machining, 'MZW8702_WIRE_CASING')
            mprog.close_file('MZW8702_WIRE_CASING')
    elif name == 'ANTI_VIBRATION_GUM':
        if GUM == '無':
            pass
        elif GUM == '穎益':
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), GUM_E_list[stamping_press_type])
            mprog.save_file_stp(machining, GUM_E_list[stamping_press_type])
            mprog.save_stpfile_part(machining, GUM_E_list[stamping_press_type])
            mprog.close_file(GUM_E_list[stamping_press_type])
        elif GUM == 'UNISORB':
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), GUM_U_list[stamping_press_type])
            mprog.save_file_stp(machining, GUM_U_list[stamping_press_type])
            mprog.save_stpfile_part(machining, GUM_U_list[stamping_press_type])
            mprog.close_file(GUM_U_list[stamping_press_type])
        elif GUM == '商定':
            mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), GUM_S_list[stamping_press_type])
            mprog.save_file_stp(machining, GUM_S_list[stamping_press_type])
            mprog.save_stpfile_part(machining, GUM_S_list[stamping_press_type])
            mprog.close_file(GUM_S_list[stamping_press_type])
    elif name == 'INVERTER':
        if INVERTER == '東元':
            if travel_type == 1:
                mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name),
                                  INVERTER_D_S_list[stamping_press_type])
                mprog.save_file_stp(machining, INVERTER_D_S_list[stamping_press_type])
                mprog.save_stpfile_part(machining, INVERTER_D_S_list[stamping_press_type])
                mprog.close_file(INVERTER_D_S_list[stamping_press_type])
            else:
                mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name),
                                  INVERTER_D_HP_list[stamping_press_type])
                mprog.save_file_stp(machining, INVERTER_D_HP_list[stamping_press_type])
                mprog.save_stpfile_part(machining, INVERTER_D_HP_list[stamping_press_type])
                mprog.close_file(INVERTER_D_HP_list[stamping_press_type])
        elif INVERTER == '台達':
            if travel_type == 1:
                mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), INVERTER_T_S_list[stamping_press_type])
                mprog.save_file_stp(machining, INVERTER_T_S_list[stamping_press_type])
                mprog.save_stpfile_part(machining, INVERTER_T_S_list[stamping_press_type])
                mprog.close_file(INVERTER_T_S_list[stamping_press_type])
            else:
                mprog.import_part(fp.system_root + fp.DEMO_part + "\\" + str(name), INVERTER_T_HP_list[stamping_press_type])
                mprog.save_file_stp(machining, INVERTER_T_HP_list[stamping_press_type])
                mprog.save_stpfile_part(machining, INVERTER_T_HP_list[stamping_press_type])
                mprog.close_file(INVERTER_T_HP_list[stamping_press_type])



def Assmebly(name,path, stamping_press_type, travel_type, GUM, INVERTER):
    if name == 'PANEL':
        mprog.import_file_Part(path, PANEL_list[stamping_press_type])
    elif name == 'CON_ROD':
        mprog.import_file_Part(path, CON_ROD_list[stamping_press_type])
    elif name == 'CON_ROD_BASE':
        mprog.import_file_Part(path, CON_ROD_BASE_list[stamping_press_type])
    elif name == 'CON_ROD_CAP':
        mprog.import_file_Part(path, CON_ROD_CAP_list[stamping_press_type])
    elif name == 'INVERTERBRACKET':
        mprog.import_file_Part(path, INVERTERBRACKET_list[stamping_press_type])
    elif name == 'POINTER':
        mprog.import_file_Part(path, POINTER_list[stamping_press_type])
    elif name == 'COVER':
        mprog.import_file_Part(path, COVER_list[stamping_press_type])
    elif name == 'PLUG':
        mprog.import_file_Part(path, PLUG_list)
    elif name == 'feeding_shaft_cover':
        mprog.import_file_Part(path, feeding_shaft_cover_list[stamping_press_type])
    elif name == 'OIL_LEVEL_GAUGE':
        mprog.import_file_Part(path, 'OGASKL060_OIL_LEVEL_GAUGE')
        mprog.import_file_Part(path, OIL_LEVEL_GAUGE_list_SLIDE[stamping_press_type])
        mprog.import_file_Part(path, OIL_LEVEL_GAUGE_list_SLIDE[stamping_press_type])
        mprog.import_file_Part(path, OIL_LEVEL_GAUGE_list_CON_ROD[stamping_press_type])
    elif name == 'slide_gib':
        mprog.import_file_Part(path, slide_gib_list_right[stamping_press_type])
        mprog.import_file_Part(path, slide_gib_list_left[stamping_press_type])
    elif name == 'ELECTRIC_BOX_PLATE':
        mprog.import_file_Part(path, 'EWR60S01_ELECTRIC_BOX_PLATE')
    elif name == 'MOUNT_FILTER':
        mprog.import_file_Part(path, 'EGPSSGD1000IS_MOUNT_FILTER')
    elif name == 'CONTROL_PANEL':
        mprog.import_file_Part(path, CONTROL_PANEL_list_normal[stamping_press_type])
    elif name == 'PANEL_BOX':
        if stamping_press_type == 0:
            mprog.import_file_Part(path, PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 1:
            mprog.import_file_Part(path, PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 2:
            mprog.import_file_Part(path, PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 3:
            mprog.import_file_Part(path, PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 4:
            mprog.import_file_Part(path, PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 5:
            mprog.import_file_Part(path, PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 6:
            mprog.import_file_Part(path, PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 7:
            mprog.import_file_Part(path, PANEL_BOX_list_normal[stamping_press_type])
        elif stamping_press_type == 8:
            mprog.import_file_Part(path, PANEL_BOX_list_normal[stamping_press_type])
    elif name == 'PANEL_BOX_BRACKET':
        if stamping_press_type == 0:
            pass
        elif stamping_press_type == 1:
            pass
        elif stamping_press_type == 2:
            pass
        elif stamping_press_type == 3:
            pass
        elif stamping_press_type == 4:
            mprog.import_file_Part(path, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
        elif stamping_press_type == 5:
            mprog.import_file_Part(path, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
        elif stamping_press_type == 6:
            mprog.import_file_Part(path, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
        elif stamping_press_type == 7:
            mprog.import_file_Part(path, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
        elif stamping_press_type == 8:
            mprog.import_file_Part(path, PANEL_BOX_BRACKET_list_normal[stamping_press_type])
    elif name == 'ELECTRIC_BOX':
        mprog.import_file_Part(path, ELECTRIC_BOX_list_normal[stamping_press_type])
    elif name == 'GUARD_FLYWHEEL':
        mprog.import_file_Part(path, GUARD_FLYWHEEL_list_normal[stamping_press_type])
    elif name == 'NAME_PLATE':
        mprog.import_file_Part(path, NAME_PLATE_list_normal[stamping_press_type])
    elif name == 'TRADEMARK_NAMEPLATE':
        mprog.import_file_Part(path, TRADEMARK_NAMEPLATE_list_normal[stamping_press_type])
    elif name == 'PORTABLE_STAND':
        mprog.import_file_Part(path, PORTABLE_STAND_list[stamping_press_type])
    elif name == 'OPERATION_BOX':
        mprog.import_file_Part(path, OPERATION_BOX_list_normal[stamping_press_type])
    elif name == 'BEARING_HOUSING':
        mprog.import_file_Part(path, BEARING_HOUSING_list[stamping_press_type])
    elif name == 'SLIDE':
        mprog.import_file_Part(path, SLIDE_list_normal[stamping_press_type])
    elif name == 'BALANCER':
        mprog.import_file_Part(path, BALANCER_list[stamping_press_type])
    elif name == 'MOTOR':
        mprog.import_file_Part(path, MOTOR_list[stamping_press_type])
    elif name == 'MOTOR_BRACKET':
        if travel_type == 1:
            mprog.import_file_Part(path, MOTOR_BRACKET_S_list[stamping_press_type])
        else:
            mprog.import_file_Part(path, MOTOR_BRACKET_HP_list[stamping_press_type])
    elif name == 'WIRE_CASING':
            mprog.import_file_Part(path, '43N6603_WIRE_CASING')
            mprog.import_file_Part(path, '43N6604_WIRE_CASING')
            mprog.import_file_Part(path,'MZW8702_WIRE_CASING')
    elif name == 'ANTI_VIBRATION_GUM':
        if GUM == '無':
            pass
        elif GUM == '穎益':
            mprog.import_file_Part(path, GUM_E_list[stamping_press_type])
        elif GUM == 'UNISORB':
            mprog.import_file_Part(path, GUM_U_list[stamping_press_type])
        elif GUM == '商定':
            mprog.import_file_Part(path, GUM_S_list[stamping_press_type])
    elif name == 'INVERTER':
        if INVERTER == '東元':
            if travel_type == 1:
                mprog.import_file_Part(path, INVERTER_D_S_list[stamping_press_type])
            else:
                mprog.import_file_Part(path, INVERTER_D_HP_list[stamping_press_type])
        elif INVERTER == '台達':
            if travel_type == 1:
                mprog.import_file_Part(path, INVERTER_T_S_list[stamping_press_type])
            else:
                mprog.import_file_Part(path, INVERTER_T_HP_list[stamping_press_type])



