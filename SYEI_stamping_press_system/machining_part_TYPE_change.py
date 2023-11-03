import main_program as mprog
import excel_parameter_change as epc
import sys


def modify_list_elements(enter_text, input_list):
    modified_list = [f'{item}%s' % (enter_text) for item in input_list]
    return modified_list


# 母檔變數變換
def change_machining_parameter(name, stamping_press_type, machiningdiepad, travel_type):
    all_part_list = epc.ExcelOp('尺寸整理表', '沖床機架零件清單').get_col_cell(1)
    try:
        if name == 'FRAME1':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME1')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME1', stamping_press_type)
                print('FRAME1 Parameter change success')
            except BaseException:
                print('FRAME1 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('110通孔', 0)
                    mprog.activatefeature('油面劑用', 0)
                    mprog.activatefeature('60通孔', 0)
                    mprog.activatefeature('電動黃油泵用', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 5)
                    mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 2)
                    mprog.activatefeature('兩點組合', 0)
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('6_M10X12L', 0)
                    mprog.partbodyfeatureactivate('JJ')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('110通孔', 0)
                    mprog.activatefeature('油面劑用', 0)
                    mprog.activatefeature('60通孔', 0)
                    mprog.activatefeature('電動黃油泵用', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 5)
                    mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 3)
                    mprog.activatefeature('兩點組合', 0)
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('6_M10X12L', 0)
                    mprog.partbodyfeatureactivate('JJ')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('110通孔', 0)
                    mprog.activatefeature('油面劑用', 0)
                    mprog.activatefeature('60通孔', 0)
                    mprog.activatefeature('電動黃油泵用', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 5)
                    mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 2)
                    mprog.activatefeature('兩點組合', 0)
                    mprog.activatefeature('6_M10X12L', 0)
                    mprog.partbodyfeatureactivate('JJ')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('110通孔', 0)
                    mprog.activatefeature('油面劑用', 0)
                    mprog.activatefeature('60通孔', 0)
                    mprog.activatefeature('電動黃油泵用', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 5)
                    mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 2)
                    mprog.activatefeature('兩點組合', 0)
                    mprog.activatefeature('6_M10X12L', 0)
                    mprog.partbodyfeatureactivate('JJ')
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.activatefeature('油面劑用', 0)
                    mprog.activatefeature('60通孔', 0)
                    mprog.activatefeature('電動黃油泵用', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 5)
                    mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 3)
                    mprog.activatefeature('兩點組合', 0)
                    mprog.activatefeature('3-M5X10L(配管用)(r1)', 1)
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.activatefeature('油面劑用', 0)
                    mprog.activatefeature('60通孔', 0)
                    mprog.activatefeature('電動黃油泵用', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 3)
                    mprog.activatefeature('兩點組合', 0)
                    mprog.activatefeature('3-M5X10L(配管用)(r1)', 2)
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.activatefeature('3-M5X10L(配管用)(r1)', 0)
                    mprog.activatefeature('油面劑用', 0)
                    mprog.activatefeature('60通孔', 0)
                    mprog.activatefeature('電動黃油泵用', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 5)
                    mprog.activatefeature('兩點組合', 0)
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.activatefeature('3-M5X10L(配管用)(r1)', 2)
                    mprog.activatefeature('油面劑用', 0)
                    mprog.activatefeature('60通孔', 0)
                    mprog.activatefeature('電動黃油泵用', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 5)
                    mprog.activatefeature('兩點組合', 0)
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('T250')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD_250')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.activatefeature('3-M5X10L(配管用)(r1)', 0)
                    mprog.activatefeature('油面劑用', 0)
                    mprog.activatefeature('60通孔', 0)
                    mprog.activatefeature('電動黃油泵用', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 0)
                    mprog.activatefeature('兩點組合', 0)
                print('FRAME1 machining success')
            except BaseException:
                print('FRAME1 machining error')
            try:
                mprog.update()
                print('FRAME1 update success')
            except BaseException:
                print('FRAME1 update error')
        elif name == 'FRAME2':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME2')
            try:
                parameter_name, parameter_value = excel.get_sheet_par(
                    'FRAME2', stamping_press_type)
                print('FRAME2 Parameter change success')
            except BaseException:
                print('FRAME2 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.activatefeature('電線穿線孔', 0)
                    mprog.activatefeature('電線護罩(長型)', 4)
                    mprog.activatefeature('馬達電線固定孔', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('風管、電線固定孔(吹氣吹料用)', 1)
                    mprog.activatefeature('風管、電線穿線孔(吹氣吹料用)', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('氣壓錶箱用', 0)
                    mprog.activatefeature('吹氣吹料(標配)', 0)
                    mprog.activatefeature('電氣箱用', 0)
                    mprog.activatefeature('安全開關', 0)
                    mprog.activatefeature('安全桿', 0)
                    mprog.activatefeature('解角器裝置線用', 0)
                    mprog.activatefeature('解角器用', 0)
                    mprog.activatefeature('變頻器用', 0)
                    mprog.activatefeature('解角器護蓋用', 0)
                    mprog.activatefeature('解碼器電線(蛇管)固定孔', 0)
                    mprog.partbodyfeatureactivate('part_S')
                    mprog.partbodyfeatureactivate('MM')
                    mprog.activatefeature('解角器固定架用', 0)
                    mprog.activatefeature('解角器固定架用_2', 0)
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('電線穿線孔', 0)
                    mprog.activatefeature('電線護罩(長型)', 4)
                    mprog.activatefeature('馬達電線固定孔', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('風管、電線固定孔(吹氣吹料用)', 1)
                    mprog.activatefeature('風管、電線穿線孔(吹氣吹料用)', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('氣壓錶箱用', 0)
                    mprog.activatefeature('吹氣吹料(標配)', 0)
                    mprog.activatefeature('電氣箱用', 0)
                    mprog.activatefeature('安全開關', 0)
                    mprog.activatefeature('安全桿', 0)
                    mprog.activatefeature('解角器裝置線用', 0)
                    mprog.activatefeature('解角器用', 0)
                    mprog.activatefeature('變頻器用', 0)
                    mprog.activatefeature('解角器護蓋用', 0)
                    mprog.activatefeature('解碼器電線(蛇管)固定孔', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.activatefeature('1-M5X10L底孔20L配管用', 0)
                    mprog.partbodyfeatureactivate('part_S')
                    mprog.partbodyfeatureactivate('MM')
                    mprog.activatefeature('解角器固定架用', 0)
                    mprog.activatefeature('解角器固定架用_2', 0)
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('電線穿線孔', 0)
                    mprog.activatefeature('電線護罩(長型)', 4)
                    mprog.activatefeature('馬達電線固定孔', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('風管、電線固定孔(吹氣吹料用)', 2)
                    mprog.activatefeature('風管、電線穿線孔(吹氣吹料用)', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('氣壓錶箱用', 0)
                    mprog.activatefeature('吹氣吹料(標配)', 0)
                    mprog.activatefeature('電氣箱用', 0)
                    mprog.activatefeature('安全開關', 0)
                    mprog.activatefeature('安全桿', 0)
                    mprog.activatefeature('解角器裝置線用', 0)
                    mprog.activatefeature('解角器用', 0)
                    mprog.activatefeature('變頻器用', 0)
                    mprog.activatefeature('解角器護蓋用', 0)
                    mprog.activatefeature('解碼器電線(蛇管)固定孔', 0)
                    mprog.activatefeature('1-M5X10L底孔20L配管用', 0)
                    mprog.partbodyfeatureactivate('part_S')
                    mprog.partbodyfeatureactivate('MM')
                    mprog.activatefeature('解角器固定架用', 0)
                    mprog.activatefeature('解角器固定架用_2', 0)
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('電線穿線孔', 0)
                    mprog.activatefeature('電線護罩(長型)', 4)
                    mprog.activatefeature('馬達電線固定孔', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('風管、電線固定孔(吹氣吹料用)', 2)
                    mprog.activatefeature('風管、電線穿線孔(吹氣吹料用)', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('氣壓錶箱用', 0)
                    mprog.activatefeature('吹氣吹料(標配)', 0)
                    mprog.activatefeature('電氣箱用', 0)
                    mprog.activatefeature('安全開關', 0)
                    mprog.activatefeature('安全桿', 0)
                    mprog.activatefeature('解角器裝置線用', 0)
                    mprog.activatefeature('解角器用', 0)
                    mprog.activatefeature('變頻器用', 0)
                    mprog.activatefeature('解角器護蓋用', 0)
                    mprog.activatefeature('解碼器電線(蛇管)固定孔', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.partbodyfeatureactivate('part_S')
                    mprog.partbodyfeatureactivate('MM')
                    mprog.activatefeature('解角器固定架用', 0)
                    mprog.activatefeature('解角器固定架用_2', 0)
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.partbodyfeatureactivate('jjj4')
                    mprog.partbodyfeatureactivate('mmm4')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('電線穿線孔', 0)
                    mprog.activatefeature('電線護罩(長型)', 4)
                    mprog.activatefeature('馬達電線固定孔', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('風管、電線固定孔(吹氣吹料用)', 2)
                    mprog.activatefeature('風管、電線穿線孔(吹氣吹料用)', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 5)
                    mprog.activatefeature('氣壓錶箱用', 0)
                    mprog.activatefeature('吹氣吹料(標配)', 0)
                    mprog.activatefeature('電氣箱用', 0)
                    mprog.activatefeature('安全開關', 0)
                    mprog.activatefeature('安全桿', 0)
                    mprog.activatefeature('解角器裝置線用', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.activatefeature('4-M8X16L', 0)
                    mprog.activatefeature('2-M5(配管用)', 0)
                    mprog.activatefeature('台達變頻器用10HP', 0)
                    mprog.activatefeature('東元變頻器用10HP', 0)
                    mprog.activatefeature('電氣箱連桿', 2)
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.partbodyfeatureactivate('dd4')
                    mprog.partbodyfeatureactivate('x4')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('電線穿線孔', 0)
                    mprog.activatefeature('電線護罩(長型)', 5)
                    mprog.activatefeature('馬達電線固定孔', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('風管、電線固定孔(吹氣吹料用)', 0)
                    mprog.activatefeature('風管、電線穿線孔(吹氣吹料用)', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('氣壓錶箱用', 0)
                    mprog.activatefeature('吹氣吹料(標配)', 0)
                    mprog.activatefeature('電氣箱用', 0)
                    mprog.activatefeature('安全開關', 0)
                    mprog.activatefeature('安全桿', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.activatefeature('4-M8X16L', 0)
                    mprog.activatefeature('2-M5(配管用)', 0)
                    mprog.activatefeature('電氣箱連桿', 0)
                    mprog.activatefeature('東元變頻器10,15HP440V用', 0)
                    mprog.activatefeature('台達變頻器10,15HP', 0)
                    mprog.activatefeature('東元變頻器,15HP220V', 0)
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.partbodyfeatureactivate('ggg4')
                    mprog.partbodyfeatureactivate('xx3')
                    mprog.partbodyfeatureactivate('iii4')
                    mprog.partbodyfeatureactivate('fff4')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('電線穿線孔', 0)
                    mprog.activatefeature('電線護罩(長型)', 5)
                    mprog.activatefeature('馬達電線固定孔', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('風管、電線固定孔(吹氣吹料用)', 2)
                    mprog.activatefeature('風管、電線穿線孔(吹氣吹料用)', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('氣壓錶箱用', 0)
                    mprog.activatefeature('吹氣吹料(標配)', 0)
                    mprog.activatefeature('電氣箱用', 0)
                    mprog.activatefeature('安全開關', 0)
                    mprog.activatefeature('安全桿', 0)
                    mprog.activatefeature('4-M8X16L', 0)
                    mprog.activatefeature('2-M5(配管用)', 0)
                    mprog.activatefeature('電氣箱連桿', 0)
                    mprog.activatefeature('東元變頻器440V,15,20HP', 0)
                    mprog.activatefeature('台達變頻器440V,15,20HP；220V15HP', 0)
                    mprog.activatefeature('台達變頻器220V,20HP', 0)
                    mprog.activatefeature('東元變頻器220V,15,20HP', 0)
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('before250_喉口')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.partbodyfeatureactivate('iii4')
                    mprog.partbodyfeatureactivate('ssss4')
                    mprog.partbodyfeatureactivate('sss4')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('電線穿線孔', 0)
                    mprog.activatefeature('電線護罩(長型)', 5)
                    mprog.activatefeature('馬達電線固定孔', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('風管、電線固定孔(吹氣吹料用)', 0)
                    mprog.activatefeature('風管、電線穿線孔(吹氣吹料用)', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('氣壓錶箱用', 0)
                    mprog.activatefeature('吹氣吹料(標配)', 0)
                    mprog.activatefeature('電氣箱用', 0)
                    mprog.activatefeature('安全開關', 0)
                    mprog.activatefeature('安全桿', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.activatefeature('4-M8X16L', 0)
                    mprog.activatefeature('2-M5(配管用)', 0)
                    mprog.activatefeature('電氣箱連桿', 0)
                    mprog.activatefeature('台達變頻器220V,20HP', 0)
                    mprog.activatefeature('東元變頻器220V/440V,20HP', 0)
                    mprog.activatefeature('台達變頻器440V,20HP', 0)
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('T250')
                    mprog.partbodyfeatureactivate('Z1')
                    mprog.partbodyfeatureactivate('machining_DD_250')
                    mprog.partbodyfeatureactivate('machining_E上方研磨')
                    mprog.activatefeature('machining_2x8_12通孔', 0)
                    mprog.activatefeature('machining_aaaaa2', 0)
                    mprog.activatefeature('machining_aaaaa4', 0)
                    mprog.activatefeature('電線穿線孔', 0)
                    mprog.activatefeature('電線護罩(長型)', 5)
                    mprog.activatefeature('馬達電線固定孔', 0)
                    mprog.activatefeature('吊孔', 0)
                    mprog.activatefeature('風管、電線固定孔(吹氣吹料用)', 2)
                    mprog.activatefeature('風管、電線穿線孔(吹氣吹料用)', 0)
                    mprog.activatefeature('光電裝置電線固定孔', 0)
                    mprog.activatefeature('氣壓錶箱用', 0)
                    mprog.activatefeature('吹氣吹料(標配)', 0)
                    mprog.activatefeature('電氣箱用', 0)
                    mprog.activatefeature('安全開關', 0)
                    mprog.activatefeature('安全桿', 0)
                    mprog.activatefeature('1-M5X10L(配管用)', 0)
                    mprog.activatefeature('4-M8X16L', 0)
                    mprog.activatefeature('2-M5(配管用)', 0)
                    mprog.activatefeature('電氣箱連桿', 0)
                    mprog.activatefeature('東元變頻器25HP', 0)
                    mprog.activatefeature('台達變頻器25HP', 0)
                print('FRAME2 machining success')
            except BaseException:
                print('FRAME2 Part activate error')
            finally:
                mprog.update()
                print('FRAME2 update success')
        elif name == 'FRAME3':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME3')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME3', stamping_press_type)
                print('FRAME3 Parameter change success')
            except:
                print('FRAME3 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                    mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                    mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
                    mprog.partbodyfeatureactivate('machining_AA')
                    mprog.partbodyfeatureactivate('machining_L')
                    mprog.partbodyfeatureactivate('machining_bbbbb')
                    mprog.activatefeature('M10X20L(不可貫穿)', 0)
                elif stamping_press_type == 1:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                    mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                    mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
                    mprog.partbodyfeatureactivate('machining_AA')
                    mprog.partbodyfeatureactivate('machining_L')
                    mprog.partbodyfeatureactivate('machining_bbbbb')
                    mprog.activatefeature('M10X20L(不可貫穿)', 0)
                elif stamping_press_type == 2:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                    mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                    mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
                    mprog.partbodyfeatureactivate('machining_AA')
                    mprog.partbodyfeatureactivate('machining_L')
                    mprog.partbodyfeatureactivate('machining_bbbbb')
                    mprog.activatefeature('M10X20L(不可貫穿)', 0)
                elif stamping_press_type == 3:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                    mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                    mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
                    mprog.partbodyfeatureactivate('machining_AA')
                    mprog.partbodyfeatureactivate('machining_L')
                    mprog.partbodyfeatureactivate('machining_bbbbb')
                    mprog.activatefeature('M10X20L(不可貫穿)', 0)
                elif stamping_press_type == 4:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                    mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
                    mprog.partbodyfeatureactivate('machining_AA')
                    mprog.partbodyfeatureactivate('machining_L')
                    mprog.partbodyfeatureactivate('machining_bbbbb')
                    mprog.activatefeature('M10X20L(不可貫穿)', 0)
                elif stamping_press_type == 5:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                    mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
                    mprog.partbodyfeatureactivate('machining_AA')
                    mprog.partbodyfeatureactivate('machining_L')
                    mprog.partbodyfeatureactivate('machining_bbbbb')
                    mprog.activatefeature('M10X20L(不可貫穿)', 0)
                elif stamping_press_type == 6:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                    mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
                    mprog.partbodyfeatureactivate('machining_AA')
                    mprog.partbodyfeatureactivate('machining_bbbbb')
                    mprog.partbodyfeatureactivate('machining_L')
                    mprog.activatefeature('M10X20L(不可貫穿)', 0)
                elif stamping_press_type == 7:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                    mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
                    mprog.partbodyfeatureactivate('machining_AA')
                    mprog.partbodyfeatureactivate('machining_bbbbb')
                    mprog.partbodyfeatureactivate('machining_L')
                elif stamping_press_type == 8:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                    mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
                    mprog.partbodyfeatureactivate('machining_AA')
                    mprog.partbodyfeatureactivate('machining_L')
                    mprog.partbodyfeatureactivate('machining_bbbbb')
                print('FRAME3 machining success')
            except:
                print('FRAME3 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME3 Update success')
                except:
                    print('FRAME3 Update error')
        elif name == 'FRAME4':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME4')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME4', stamping_press_type)
                print('FRAME4 Parameter change success')
            except:
                print('FRAME4 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.partbodyfeatureactivate("machining_P")
                    mprog.partbodyfeatureactivate("machining_bbbbb")
                elif stamping_press_type == 1:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.partbodyfeatureactivate("machining_P")
                    mprog.partbodyfeatureactivate("machining_O")
                    mprog.partbodyfeatureactivate("machining_bbbbb")
                elif stamping_press_type == 2:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.partbodyfeatureactivate("machining_P")
                    mprog.partbodyfeatureactivate("machining_O")
                    mprog.partbodyfeatureactivate("machining_bbbbb")
                elif stamping_press_type == 3:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 2)
                    mprog.partbodyfeatureactivate("machining_N")
                    mprog.partbodyfeatureactivate("machining_O")
                    mprog.partbodyfeatureactivate("machining_M")
                    mprog.partbodyfeatureactivate("machining_P")
                    mprog.partbodyfeatureactivate("machining_bbbbb")
                elif stamping_press_type == 4:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 2)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("machining_N")
                    mprog.partbodyfeatureactivate("machining_O")
                    mprog.partbodyfeatureactivate("machining_M")
                    mprog.partbodyfeatureactivate("machining_P")
                    mprog.partbodyfeatureactivate("machining_bbbbb")
                elif stamping_press_type == 5:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 2)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("machining_N")
                    mprog.partbodyfeatureactivate("machining_O")
                    mprog.partbodyfeatureactivate("machining_M")
                    mprog.partbodyfeatureactivate("machining_P")
                    mprog.partbodyfeatureactivate("machining_bbbbb")
                elif stamping_press_type == 6:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 2)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("machining_N")
                    mprog.partbodyfeatureactivate("machining_O")
                    mprog.partbodyfeatureactivate("machining_M")
                    mprog.partbodyfeatureactivate("machining_P")
                    mprog.partbodyfeatureactivate("machining_bbbbb")
                elif stamping_press_type == 7:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 2)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("machining_N")
                    mprog.partbodyfeatureactivate("machining_O")
                    mprog.partbodyfeatureactivate("machining_M")
                    mprog.partbodyfeatureactivate("machining_P")
                    mprog.partbodyfeatureactivate("machining_bbbbb")
                elif stamping_press_type == 8:
                    mprog.activatefeature('SN1_25250_Body', 1)
                    mprog.activatefeature('Hole_1', 1)
                    mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                    mprog.activatefeature('Hole_3', 4)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("machining_N")
                    mprog.partbodyfeatureactivate("machining_O")
                    mprog.partbodyfeatureactivate("machining_M")
                    mprog.partbodyfeatureactivate("machining_P")
                    mprog.partbodyfeatureactivate("machining_bbbbb")
                print('FRAME4 machining success')
            except:
                print('FRAME4 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME4 Update success')
                except:
                    print('FRAME4 Update error')
        elif name == 'FRAME5':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME5')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME5', stamping_press_type)
                print('FRAME5 Parameter change success')
            except:
                print('FRAME5 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('SN1-25_X')
                    mprog.partbodyfeatureactivate('CD')
                    mprog.partbodyfeatureactivate('SN1_25_H')
                    mprog.activatefeature('2-M8通孔', 0)
                    mprog.activatefeature('TY牌電磁閥用', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                    mprog.activatefeature('PT1/2', 0)
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('CD')
                    mprog.partbodyfeatureactivate("SN1_3545_H")
                    mprog.activatefeature('2-M8通孔', 0)
                    mprog.activatefeature('TY牌電磁閥用', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                    mprog.activatefeature('PT1/2', 0)
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('CD')
                    mprog.partbodyfeatureactivate("SN1_3545_H")
                    mprog.activatefeature('2-M8通孔', 0)
                    mprog.activatefeature('TY牌電磁閥用', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                    mprog.activatefeature('PT1/2', 0)
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('CD')
                    mprog.partbodyfeatureactivate('NM')
                    mprog.partbodyfeatureactivate('SN1_60_R')
                    mprog.activatefeature('2-M8通孔', 0)
                    mprog.activatefeature('TY牌電磁閥用', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('SN1_80250_OP')
                    mprog.partbodyfeatureactivate('SN1_80250_Q')
                    mprog.activatefeature('2-M8通孔', 0)
                    mprog.activatefeature('TY牌電磁閥用', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('SN1_80250_OP')
                    mprog.partbodyfeatureactivate('SN1_80250_Q')
                    mprog.partbodyfeatureactivate('6-M6通C5')
                    mprog.activatefeature('6-M6通', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('SN1_80250_OP')
                    mprog.partbodyfeatureactivate('SN1_80250_Q')
                    mprog.partbodyfeatureactivate('3-M8通_C5')
                    mprog.activatefeature('3-M8通', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('SN1_80250_OP')
                    mprog.partbodyfeatureactivate('SN1_80250_Q')
                    mprog.partbodyfeatureactivate('5-M8通_C5')
                    mprog.activatefeature('5-M8通', 0)
                    mprog.activatefeature('M6通孔LED', 0)
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('SN1_80250_OP')
                    mprog.partbodyfeatureactivate('SN1_80250_Q')
                    mprog.partbodyfeatureactivate('5-M8通_C5')
                    mprog.activatefeature('5-M8通', 0)
                    mprog.activatefeature('M6通孔LED', 0)
            except BaseException:
                print('FRAME5 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME5 Update success')
                except BaseException:
                    print('FRAME5 Update error')
        elif name == 'FRAME6':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME6')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME6', stamping_press_type)
                print('FRAME6 Parameter change success')
            except:
                print('FRAME6 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                    mprog.activatefeature('M16_i', 0)
                elif stamping_press_type == 1:
                    mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                    mprog.activatefeature('M16_i', 0)
                elif stamping_press_type == 2:
                    mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                    mprog.activatefeature('M16_i', 0)
                elif stamping_press_type == 3:
                    mprog.activatefeature('FRAME_SN1_60_Body', 0)
                    mprog.activatefeature('M16_h FRAME9_L', 0)
                elif stamping_press_type == 4:
                    mprog.activatefeature('FRAME_SN1_80110_Body', 0)
                    mprog.activatefeature('M16_h FRAME9_L', 0)
                elif stamping_press_type == 5:
                    mprog.activatefeature('FRAME_SN1_80110_Body', 0)
                    mprog.activatefeature('M16_h FRAME9_L', 0)
                elif stamping_press_type == 7:
                    mprog.activatefeature('FRAME_SN1_200_Body', 0)
                    mprog.activatefeature('M16_h FRAME9_L', 0)
                print('FRAME6 machining success')
            except:
                print('FRAME6 machining activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME6 Update success')
                except:
                    print('FRAME6 Update error')
        elif name == 'FRAME7':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME7')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME7', stamping_press_type)
                print('FRAME7 Parameter change success')
            except:
                print('FRAME7 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.activatefeature('SN1_2560_Body', 1)
                    mprog.partbodyfeatureactivate('SN1_25_VWX')
                    mprog.partbodyfeatureactivate('DD')
                elif stamping_press_type == 1:
                    mprog.activatefeature('SN1_2560_Body', 1)
                    mprog.partbodyfeatureactivate('DD')
                elif stamping_press_type == 2:
                    mprog.activatefeature('SN1_2560_Body', 1)
                    mprog.partbodyfeatureactivate('DD')
                elif stamping_press_type == 3:
                    mprog.activatefeature('SN1_2560_Body', 1)
                    mprog.partbodyfeatureactivate('DD')
                elif stamping_press_type == 4:
                    mprog.activatefeature('SN1_80250_Body', 4)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('DD')
                elif stamping_press_type == 5:
                    mprog.activatefeature('SN1_80250_Body', 4)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('DD')
                elif stamping_press_type == 6:
                    mprog.activatefeature('SN1_80250_Body', 4)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('DD')
                elif stamping_press_type == 7:
                    mprog.activatefeature('SN1_80250_Body', 4)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('DD')
                elif stamping_press_type == 8:
                    mprog.activatefeature('SN1_80250_Body', 4)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('DD')
                print('FRAME7 machining success')
            except:
                print('FRAME7 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME7 Update success')
                except:
                    print('FRAME7 Update error')
        elif name == 'FRAME8':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME8')
            try:
                parameter_name, parameter_value = excel.get_sheet_par(
                    'FRAME8', stamping_press_type)
                print('FRAME8 Parameter change success')
            except:
                print('FRAME8 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('上半倒角_25')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate("Pad.2")
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate("Pad.2")
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate("Pad.2")
                elif stamping_press_type == 4:
                    # mprog.partbodyfeatureactivate('上半倒角80')
                    mprog.partbodyfeatureactivate('R40特徵')
                    # mprog.partbodyfeatureactivate('下半開槽')
                    mprog.activatefeature('38孔通', 0)
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('R40特徵')
                    mprog.activatefeature('38孔通', 0)
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('R40特徵')
                    mprog.activatefeature('38孔通', 0)
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('R40特徵')
                    mprog.activatefeature('38孔通', 0)
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('R40特徵')
                    mprog.activatefeature('38孔通', 0)
                print('FRAME8 machining success')
            except BaseException:
                print('FRAME8 machining activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME8 Update success')
                except BaseException:
                    print('FRAME8 Update error')
        elif name == 'FRAME9':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME9')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME9', stamping_press_type)
                print('FRAME9 Parameter change success')
            except:
                print('FRAME9 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.activatefeature('SN1_2580_Body', 1)
                elif stamping_press_type == 1:
                    mprog.activatefeature('SN1_2580_Body', 1)
                elif stamping_press_type == 2:
                    mprog.activatefeature('SN1_2580_Body', 1)
                elif stamping_press_type == 3:
                    mprog.activatefeature('SN1_2580_Body', 1)
                elif stamping_press_type == 4:
                    mprog.activatefeature('channel steel', 0)
                elif stamping_press_type == 5:
                    mprog.activatefeature('channel steel', 0)
                elif stamping_press_type == 6:
                    mprog.activatefeature('channel steel', 0)
                elif stamping_press_type == 7:
                    mprog.activatefeature('channel steel', 0)
                elif stamping_press_type == 8:
                    mprog.activatefeature('channel steel', 0)
                print('FRAME9 machining success')
            except:
                print('FRAME9 machining activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME9 Update success')
                except:
                    print('FRAME9 Update error')
        elif name == 'FRAME10':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME10')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME10', stamping_press_type)
                print('FRAME10 Parameter change success')
            except:
                print('FRAME10 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('SN1_25250_K')
                elif stamping_press_type == 1:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('SN1_25250_K')
                elif stamping_press_type == 2:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('SN1_25250_K')
                elif stamping_press_type == 3:
                    mprog.activatefeature('SN1_60_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('SN1_25250_K')
                elif stamping_press_type == 4:
                    mprog.activatefeature('SN1_80110_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('SN1_25250_K')
                elif stamping_press_type == 5:
                    mprog.activatefeature('SN1_80110_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('SN1_25250_K')
                elif stamping_press_type == 6:
                    mprog.activatefeature('SN1_160250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('SN1_25250_K')
                elif stamping_press_type == 7:
                    mprog.activatefeature('SN1_160250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('SN1_25250_K')
                elif stamping_press_type == 8:
                    mprog.activatefeature('SN1_160250_Body', 0)
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('SN1_25250_K')
                print('FRAME10 machining success')
            except:
                print('FRAME10 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME10 Update success')
                except:
                    print('FRAME10 Update error')
        elif name == 'FRAME11':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME11')
            try:
                parameter_name, parameter_value = excel.get_sheet_par(
                    'FRAME11', stamping_press_type)
                print('FRAME11 Parameter change success')
            except BaseException:
                print('FRAME11 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.activatefeature('挖孔', 0)
                    mprog.activatefeature('RC1/8', 0)
                elif stamping_press_type == 1:
                    mprog.activatefeature('挖孔', 0)
                    mprog.activatefeature('RC1/8', 0)
                elif stamping_press_type == 2:
                    mprog.activatefeature('挖孔', 0)
                    mprog.activatefeature('RC1/8', 0)
                elif stamping_press_type == 3:
                    mprog.activatefeature('挖孔', 0)
                    mprog.activatefeature('RC1/8', 0)
                elif stamping_press_type == 4:
                    mprog.activatefeature('挖孔', 0)
                    mprog.activatefeature('RC1/8', 0)
                elif stamping_press_type == 5:
                    mprog.activatefeature('挖孔', 0)
                    mprog.activatefeature('RC1/8', 0)
                elif stamping_press_type == 6:
                    mprog.activatefeature('挖孔', 0)
                elif stamping_press_type == 7:
                    mprog.activatefeature('挖孔', 0)
                    mprog.activatefeature('RC1/8', 0)
                elif stamping_press_type == 8:
                    mprog.activatefeature('挖孔', 0)
                    mprog.activatefeature('RC1/8', 0)
                print('FRAME11 machining success')
            except BaseException:
                print('FRAME11 machining activate error')
            try:
                mprog.Update()
                print('FRAME11 Update success')
            except BaseException:
                print('FRAME11 Update error')
        elif name == 'FRAME12':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME12')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME12', stamping_press_type)
                print('FRAME12 Parameter change success')
            except:
                print('FRAME12 Parameter change error')
            try:
                if stamping_press_type == 0:  # 25N
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.activatefeature('Pipe_clamp', 2)
                    mprog.activatefeature('drain_hole', 0)
                elif stamping_press_type == 1:  # 35N
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.activatefeature('Pipe_clamp', 2)
                    mprog.activatefeature('drain_hole', 0)
                elif stamping_press_type == 2:  # 45N
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.activatefeature('Pipe_clamp', 2)
                    mprog.activatefeature('drain_hole', 0)
                elif stamping_press_type == 3:  # 60N
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.activatefeature('Pipe_clamp', 2)
                    mprog.activatefeature('drain_hole', 0)
                elif stamping_press_type == 4:  # 80N
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.activatefeature('Pipe_clamp', 2)
                    mprog.activatefeature('drain_hole', 0)
                elif stamping_press_type == 5:  # 110N
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.activatefeature('Pipe_clamp', 2)
                    mprog.activatefeature('drain_hole', 0)
                elif stamping_press_type == 6:  # 160N
                    mprog.partbodyfeatureactivate('C8_160')
                    mprog.activatefeature('Pipe_clamp', 2)
                    mprog.activatefeature('drain_hole', 0)
                elif stamping_press_type == 7:  # 200N
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.activatefeature('Pipe_clamp', 0)
                    mprog.activatefeature('drain_hole', 0)
                elif stamping_press_type == 8:  # 250N
                    mprog.partbodyfeatureactivate('Chamfer.1')
                    mprog.activatefeature('Pipe_clamp', 0)
                    mprog.activatefeature('drain_hole', 0)
            except:
                print('FRAME12 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME12 Update success')
                except:
                    print('FRAME12 Update error')
        elif name == 'FRAME13':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME13')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME13', stamping_press_type)
                print('FRAME13 Parameter change success')
            except:
                print('FRAME13 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_25_45_CD')
                    mprog.partbodyfeatureactivate('SN1_25_45_G')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_25_45_CD')
                    mprog.partbodyfeatureactivate('SN1_25_45_G')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
            except:
                print('FRAME13 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME13 Update success')
                except:
                    print('FRAME13 Update error')
        elif name == 'FRAME13_1':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME13_1')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME13_1', stamping_press_type)
                print('FRAME13_1 Parameter change success')
            except:
                print('FRAME13_1 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_25_45_CD')
                    mprog.partbodyfeatureactivate('SN1_25_45_G')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_25_45_CD')
                    mprog.partbodyfeatureactivate('SN1_25_45_G')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('Hole_4', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
            except:
                print('FRAME13_1 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME13_1 Update success')
                except:
                    print('FRAME13_1 Update error')
        elif name == 'FRAME14':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME14')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME14', stamping_press_type)
                print('FRAME14 Parameter change success')
            except:
                print('FRAME14 Parameter change error')
            try:
                if stamping_press_type == 0:  # 25N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 1:  # 35N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 2:  # 45N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 3:  # 60N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 4:  # 80N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('outside_hole_E', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 5:  # 110N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 6:  # 160N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 7:  # 200N
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('outside_hole_E', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 8:  # 250N
                    mprog.partbodyfeatureactivate("250N")
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                print('FRAME14 machining success')
            except:
                print('FRAME14 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME14 Update success')
                except:
                    print('FRAME14 Update error')
        elif name == 'FRAME15':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME15')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME15', stamping_press_type)
                print('FRAME15 Parameter change success')
            except:
                print('FRAME15 Parameter change error')
            try:
                if stamping_press_type == 0:  # 25N
                    mprog.activatefeature('Hole', 0)
                    mprog.partbodyfeatureactivate('F')
                elif stamping_press_type == 1:  # 35N
                    mprog.activatefeature('Hole', 0)
                    mprog.partbodyfeatureactivate('F')
                elif stamping_press_type == 2:  # 45N
                    mprog.activatefeature('Hole', 0)
                    mprog.partbodyfeatureactivate('F')
                elif stamping_press_type == 3:  # 60N
                    mprog.activatefeature('Hole', 0)
                    # mprog.partbodyfeatureactivate('F')
                elif stamping_press_type == 4:  # 80N
                    mprog.activatefeature('Hole', 0)
                    # mprog.partbodyfeatureactivate('F')
                elif stamping_press_type == 5:  # 110N
                    mprog.activatefeature('Hole', 0)
                    mprog.partbodyfeatureactivate('F')
                elif stamping_press_type == 6:  # 160N
                    mprog.activatefeature('Hole', 0)
                    # mprog.partbodyfeatureactivate('F')
                elif stamping_press_type == 7:  # 200N
                    mprog.activatefeature('Hole', 0)
                    mprog.partbodyfeatureactivate('F')
                elif stamping_press_type == 8:  # 250N
                    # mprog.partbodyfeatureactivate('SN1_200250')
                    mprog.activatefeature('Hole', 0)
                    mprog.partbodyfeatureactivate('F')
            except:
                print('FRAME15 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME15 Update success')
                except:
                    print('FRAME15 Update error')
        elif name == 'FRAME16':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME16')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME16', stamping_press_type)
                print('FRAME16 Parameter change success')
            except:
                print('FRAME16 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_25_45_CD')
                    mprog.partbodyfeatureactivate('SN1_25_45_G')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_25_45_CD')
                    mprog.partbodyfeatureactivate('SN1_25_45_G')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 1)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('Hole_4', 0)
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                    mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                    mprog.activatefeature('GIB_OIL_HOLE', 0)
                    mprog.activatefeature('Hole_4', 0)
            except:
                print('FRAME16 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME16 Update success')
                except:
                    print('FRAME16 Update error')
        elif name == 'FRAME17':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME17')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME17', stamping_press_type)
                print('FRAME17 Parameter change success')
            except:
                print('FRAME17 Parameter change error')
            try:
                if stamping_press_type == 0:  # 25N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 1:  # 35N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 2:  # 45N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 3:  # 60N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 4:  # 80N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('outside_hole_E', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 5:  # 110N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 6:  # 160N
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 7:  # 200N
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('outside_hole_E', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                elif stamping_press_type == 8:  # 250N
                    mprog.partbodyfeatureactivate("250N")
                    mprog.activatefeature('inside_hole', 0)
                    mprog.activatefeature('outside_hole', 0)
                    mprog.activatefeature('processing_h', 0)
                    mprog.activatefeature('processing_i', 0)
                    mprog.activatefeature('processing_m', 0)
                    mprog.partbodyfeatureactivate("H")
                    mprog.partbodyfeatureactivate("I")
                print('FRAME17 machining success')
            except:
                print('FRAME17 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME17 Update success')
                except:
                    print('FRAME17 Update error')
        elif name == 'FRAME18':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME18')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME18', stamping_press_type)
                print('FRAME18 Parameter change success')
            except:
                print('FRAME18 Parameter change error')

            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('Pad.2')
                    mprog.partbodyfeatureactivate('HollowCylinder')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('Pad.2')
                    mprog.partbodyfeatureactivate('HollowCylinder')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('Pad.2')
                    mprog.partbodyfeatureactivate('HollowCylinder')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('Pad.2')
                    mprog.partbodyfeatureactivate('HollowCylinder')
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('Pad.2')
                    mprog.partbodyfeatureactivate('HollowCylinder')
                    mprog.partbodyfeatureactivate('EdgeFillet.1')
                elif stamping_press_type == 5:
                    mprog.activatefeature('Hole.1', 0)
                    mprog.partbodyfeatureactivate('EdgeFillet.1')
                elif stamping_press_type == 6:
                    mprog.activatefeature('Hole.1', 0)
                    mprog.partbodyfeatureactivate('EdgeFillet.1')
                elif stamping_press_type == 7:
                    mprog.activatefeature('Hole.1', 0)
                    mprog.partbodyfeatureactivate('EdgeFillet.1')
                elif stamping_press_type == 8:
                    mprog.activatefeature('Hole.1', 0)
                    mprog.partbodyfeatureactivate('EdgeFillet.1')
            except:
                print('FRAME18 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME18 Update success')
                except:
                    print('FRAME18 Update error')
        elif name == 'FRAME19':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME19')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME19', stamping_press_type)
                print('FRAME19 Parameter change success')
            except:
                print('FRAME19 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.activatefeature('SN1_2545_Body', 6)
                    mprog.partbodyfeatureactivate('SN1_25_X')
                    mprog.partbodyfeatureactivate('SN1_2535_Y')
                    mprog.partbodyfeatureactivate("machining_AI")
                elif stamping_press_type == 1:
                    mprog.activatefeature('SN1_2545_Body', 6)
                    mprog.partbodyfeatureactivate('SN1_2535_Y')
                    mprog.partbodyfeatureactivate("SN1_3545_AD")
                    mprog.partbodyfeatureactivate("machining_AI")
                elif stamping_press_type == 2:
                    mprog.activatefeature('SN1_2545_Body', 6)
                    mprog.partbodyfeatureactivate('SN1_45_Y')
                    mprog.partbodyfeatureactivate('SN1_45_X')
                    mprog.partbodyfeatureactivate("SN1_3545_AD")
                    mprog.partbodyfeatureactivate("machining_AI")
                elif stamping_press_type == 3:
                    mprog.activatefeature('SN1_60_Body', 8)
                    mprog.partbodyfeatureactivate("machining_AI")
                elif stamping_press_type == 4:
                    mprog.activatefeature('SN1_80110_Body', 5)
                    mprog.partbodyfeatureactivate("machining_AI")
                elif stamping_press_type == 5:
                    mprog.activatefeature('SN1_80110_Body', 5)
                    mprog.partbodyfeatureactivate("machining_AI")
                elif stamping_press_type == 6:
                    mprog.activatefeature('FRAME_SN1_160_Body', 0)
                    mprog.activatefeature('FRAME_34_Hole_1', 0)
                    mprog.partbodyfeatureactivate("machining_AI")
                elif stamping_press_type == 7:
                    mprog.activatefeature('SN1_200_Body', 4)
                    mprog.partbodyfeatureactivate("machining_AI")
                elif stamping_press_type == 8:
                    mprog.activatefeature('SN1_250_Body', 3)
                    mprog.partbodyfeatureactivate("machining_AI")
                print('FRAME19 machining success')
            except:
                print('FRAME19 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME19 Update success')
                except:
                    print('FRAME19 Update error')
        elif name == 'FRAME20':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME20')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME20', stamping_press_type)
                print('FRAME20 Parameter change success')
            except:
                print('FRAME20 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.partbodyfeatureactivate('SN1_80250_DEF')
                    mprog.partbodyfeatureactivate('SN1_80250_G')
                    mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.partbodyfeatureactivate('SN1_80250_DEF')
                    mprog.partbodyfeatureactivate('SN1_80250_G')
                    mprog.partbodyfeatureactivate('SN1_110160_ijk')
                    mprog.activatefeature('Hole_2', 0)
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.partbodyfeatureactivate('SN1_80250_DEF')
                    mprog.partbodyfeatureactivate('SN1_80250_G')
                    mprog.partbodyfeatureactivate('SN1_110160_ijk')
                    mprog.activatefeature('Hole_2', 0)
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60250_C')
                    mprog.partbodyfeatureactivate('SN1_80250_DEF')
                    mprog.partbodyfeatureactivate('SN1_80250_G')
                    mprog.partbodyfeatureactivate('SN1_200250_hij')
                    mprog.activatefeature('Hole_2', 1)
                elif stamping_press_type == 8:
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
                    print('FRAME20 Update success')
                except:
                    print('FRAME20 Update error')
        elif name == 'FRAME21':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME21')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME21', stamping_press_type)
                print('FRAME21 Parameter change success')
            except:
                print('FRAME21 Parameter change error')
            try:
                mprog.partdeactivate("H")
                if stamping_press_type == 0:  # 25
                    mprog.activatefeature('Hole', 0)
                elif stamping_press_type == 1:  # 35
                    mprog.activatefeature('Hole', 0)
                print('FRAME21 machining success')
            except:
                print('FRAME21 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME21 Update success')
                except:
                    print('FRAME21 Update error')
        elif name == 'FRAME22':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME22')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME22', stamping_press_type)
                print('FRAME22 Parameter change success')
            except:
                print('FRAME22 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('A(1)D')
                    mprog.partbodyfeatureactivate('SN1-2560_BE')
                    mprog.partbodyfeatureactivate('SN1_2545_I')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('A(1)D')
                    mprog.partbodyfeatureactivate('SN1-2560_BE')
                    mprog.partbodyfeatureactivate('SN1_2545_I')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('A(1)D')
                    mprog.partbodyfeatureactivate('SN1-2560_BE')
                    mprog.partbodyfeatureactivate('SN1_2545_I')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('A(1)D')
                    mprog.partbodyfeatureactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1-2560_BE')
                    mprog.partbodyfeatureactivate('SN1_60_I')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('Hole_3', 0)
                    mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                    mprog.partbodyfeatureactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                    mprog.partbodyfeatureactivate('SN1_80250_IN')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                    mprog.activatefeature('Hole_3', 0)
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                    mprog.partbodyfeatureactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                    mprog.partbodyfeatureactivate('SN1_80250_IN')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                    mprog.activatefeature('Hole_3', 0)
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                    mprog.partbodyfeatureactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                    mprog.partbodyfeatureactivate('SN1_80250_IN')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                    mprog.activatefeature('Hole_3', 0)
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                    mprog.partbodyfeatureactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                    mprog.partbodyfeatureactivate('SN1_80250_IN')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                    mprog.activatefeature('Hole_3', 0)
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('SN1_80250_A(1)D')
                    mprog.partbodyfeatureactivate('A(2)')
                    mprog.partbodyfeatureactivate('SN1_80250_FEBC')
                    mprog.partbodyfeatureactivate('SN1_80250_IN')
                    mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                    mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                    mprog.activatefeature('Hole_3', 0)
            except:
                print('FRAME22 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME22 Update success')
                except:
                    print('FRAME22 Update error')
        elif name == 'FRAME23':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME23')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME23', stamping_press_type)
                print('FRAME23 Parameter change success')
            except:
                print('FRAME23 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('80至250除料')
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('80至250除料')
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
            except:
                print('FRAME23 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME23 Update success')
                except:
                    print('FRAME23 Update error')
        elif name == 'FRAME23_1':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME23_1')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME23_1', stamping_press_type)
                print('FRAME23_1 Parameter change success')
            except:
                print('FRAME23_1 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('25至60除料')
                    mprog.partbodyfeatureactivate('25至60導圓角')
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('80至250除料')
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('80至250除料')
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('80至250除料')
                    mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
            except:
                print('FRAME23_1 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME23_1 Update success')
                except:
                    print('FRAME23_1 Update error')
        elif name == 'FRAME24':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME24')
            try:
                parameter_name, parameter_value = excel.get_sheet_par(
                    'FRAME24', stamping_press_type)
                print('FRAME24 Parameter change success')
            except BaseException:
                print('FRAME24 Parameter change error')
            if stamping_press_type == 0:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 1:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 2:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 3:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 4:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 5:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 6:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 7:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 8:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            try:
                mprog.Update()
                print('FRAME24 Update success')
            except BaseException:
                print('FRAME24 Update error')
        elif name == 'FRAME24_1':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME24_1')
            try:
                parameter_name, parameter_value = excel.get_sheet_par(
                    'FRAME24_1', stamping_press_type)
                print('FRAME24_1 Parameter change success')
            except BaseException:
                print('FRAME24_1 Parameter change error')
            if stamping_press_type == 0:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 1:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 2:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 3:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 4:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 5:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 6:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 7:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            elif stamping_press_type == 8:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.activatefeature('M5', 0)
                mprog.partbodyfeatureactivate('machining_bbbbb')
                mprog.partbodyfeatureactivate('machining_bbbbb_2')
                mprog.activatefeature('machining_Hole_1', 0)
            try:
                mprog.Update()
                print('FRAME24_1 Update success')
            except BaseException:
                print('FRAME24_1 Update error')
        elif name == 'FRAME25':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME25')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME25', stamping_press_type)
                print('FRAME25 Parameter change success')
            except:
                print('FRAME25 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('3_M8通', 0)
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('3_M8通', 0)
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('3_M8通', 0)
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.activatefeature('3_M8通', 0)
            except:
                print('FRAME25 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME25 Update success')
                except:
                    print('FRAME25 Update error')
        elif name == 'FRAME26':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME26')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME26', stamping_press_type)
                print('FRAME26 Parameter change success')
            except:
                print('FRAME26 Parameter change error')

            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('F')
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('F')
                    mprog.partbodyfeatureactivate('G')
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('250除料')
            except:
                print('FRAME26 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME26 Update success')
                except:
                    print('FRAME26 Update error')
        elif name == 'FRAME27':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME27')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME27', stamping_press_type)
                print('FRAME27 Parameter change success')
            except:
                print('FRAME27 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('35N前')
                    mprog.activatefeature('Hole', 0)
                    mprog.activatefeature('Piping', 0)
                    mprog.partbodyfeatureactivate('machining_aaaaa2')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('35N前')
                    mprog.activatefeature('Hole', 0)
                    mprog.activatefeature('Piping', 0)
                    mprog.partbodyfeatureactivate('machining_aaaaa2')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('Pocket.1')
                    mprog.activatefeature('Hole', 0)
                    mprog.activatefeature('Piping', 0)
                    mprog.activatefeature('machining_ddddd', 0)
                    mprog.partbodyfeatureactivate('machining_aaaaa2')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('Pocket.1')
                    mprog.activatefeature('Hole', 0)
                    mprog.activatefeature('Piping', 0)
                    mprog.activatefeature('machining_ddddd', 0)
                    mprog.partbodyfeatureactivate('machining_aaaaa2')
            except BaseException:
                print('FRAME27 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME27 Update success')
                except BaseException:
                    print('FRAME27 Update error')
        elif name == 'FRAME27_1':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME27_1')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME27_1', stamping_press_type)
                print('FRAME27_1 Parameter change success')
            except:
                print('FRAME27_1 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('35N前')
                    mprog.activatefeature('Hole', 0)
                    mprog.activatefeature('Piping', 0)
                    mprog.partbodyfeatureactivate('machining_aaaaa2')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('35N前')
                    mprog.activatefeature('Hole', 0)
                    mprog.activatefeature('Piping', 0)
                    mprog.partbodyfeatureactivate('machining_aaaaa2')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('Pocket.1')
                    mprog.activatefeature('Hole', 0)
                    mprog.activatefeature('Piping', 0)
                    mprog.activatefeature('machining_ddddd', 0)
                    mprog.partbodyfeatureactivate('machining_aaaaa2')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('Pocket.1')
                    mprog.activatefeature('Hole', 0)
                    mprog.activatefeature('Piping', 0)
                    mprog.activatefeature('machining_ddddd', 0)
                    mprog.partbodyfeatureactivate('machining_aaaaa2')
            except BaseException:
                print('FRAME27_1 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME27_1 Update success')
                except BaseException:
                    print('FRAME27_1 Update error')
        elif name == 'FRAME28':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME28')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME28', stamping_press_type)
                print('FRAME28 Parameter change success')
            except:
                print('FRAME28 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('SN1_25250_ABD')
                    mprog.partbodyfeatureactivate('SN1_25250_E')
                    mprog.partbodyfeatureactivate('SN1_25250_F')
            except:
                print('FRAME28 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME28 Update success')
                except:
                    print('FRAME28 Update error')
        elif name == 'FRAME29':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME29')
            try:
                parameter_name, parameter_value = excel.get_sheet_par(
                    'FRAME29', stamping_press_type)
                print('FRAME29 Parameter change success')
            except BaseException:
                print('FRAME29 Parameter change error')
            try:
                if stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('Hole.1')
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('Hole.1')
                else:
                    mprog.partbodyfeatureactivate('Hole.1')
                    mprog.partbodyfeatureactivate("F")
            finally:
                try:
                    mprog.Update()
                    print('FRAME29 Update success')
                except BaseException:
                    print('FRAME29 Update error')
        elif name == 'FRAME30':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME30')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME30', stamping_press_type)
                print('FRAME30 Parameter change success')
            except:
                print('FRAME30 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                    mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                    mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.activatefeature('4_M14通', 0)
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                    mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                    mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.activatefeature('4_M14通', 0)
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                    mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                    mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.activatefeature('4_M14通', 0)
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                    mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                    mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.activatefeature('4_M14通', 0)
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                    mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                    mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.activatefeature('4_M14通', 0)
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                    mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                    mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.activatefeature('4_M14通', 0)
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                    mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                    mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.activatefeature('4_M14通', 0)
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                    mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                    mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.activatefeature('4_M14通', 0)
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('SN1_25250_AE')
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                    mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                    mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                    mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                    mprog.activatefeature('4_M14通', 0)
            except:
                print('FRAME30 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME30 Update success')
                except:
                    print('FRAME30 Update error')
        elif name == 'FRAME31':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME31')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME31', stamping_press_type)
                print('FRAME31 Parameter change success')
            except:
                print('FRAME31 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.activatefeature('FRAME_Body', 2)
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.activatefeature('FRAME_Body', 0)
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.activatefeature('FRAME_Body', 0)
            except:
                print('FRAME31 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME31 Update success')
                except:
                    print('FRAME31 Update error')
        elif name == 'FRAME32':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME32')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME32', stamping_press_type)
                print('FRAME32 Parameter change success')
            except:
                print('FRAME32 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('Body')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('Body')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('Body')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('Body')
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('Body')
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('Body')
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('Body')
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('Body')
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('Body')
            except:
                print('FRAME32 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME32 Update success')
                except:
                    print('FRAME32 Update error')
        elif name == "FRAME33":
            excel = epc.ExcelOp('尺寸整理表', 'FRAME33')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME33', stamping_press_type)
                print('FRAME33 Parameter change success')
            except:
                print('FRAME33 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CEF')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('FRAME_2545_LMN')
                    mprog.partbodyfeatureactivate("O")
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CEF')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('FRAME_2545_LMN')
                    mprog.partbodyfeatureactivate("O")
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CEF')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate('FRAME_2545_LMN')
                    mprog.partbodyfeatureactivate("O")
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.activatefeature('Hole_1', 0)
                    mprog.partbodyfeatureactivate("O")
            except:
                print('FRAME33 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME33 Update success')
                except:
                    print('FRAME33 Update error')
        elif name == 'FRAME34':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME34')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME34', stamping_press_type)
                print('FRAME34 Parameter change success')
            except:
                print('FRAME34 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CD')
                    mprog.partbodyfeatureactivate('SN1_2535_J')
                    mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_1_2545_PQR')
                    mprog.partbodyfeatureactivate("S")
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CD')
                    mprog.partbodyfeatureactivate('SN1_2535_J')
                    mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_1_2545_PQR')
                    mprog.partbodyfeatureactivate("S")
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_2545_CD')
                    mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate('FRAME_1_2545_PQR')
                    mprog.partbodyfeatureactivate("S")
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('AB')
                    mprog.partbodyfeatureactivate('SN1_60_NM')
                    mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                    mprog.activatefeature('Hole_2', 0)
                    mprog.partbodyfeatureactivate("S")
            except:
                print('FRAME34 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME34 Update success')
                except:
                    print('FRAME34 Update error')
        elif name == 'FRAME35':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME35')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME35', stamping_press_type)
                print('FRAME35 Parameter change success')
            except:
                print('FRAME35 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                elif stamping_press_type == 1:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                elif stamping_press_type == 2:
                    mprog.activatefeature('SN1_2545_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                elif stamping_press_type == 3:
                    mprog.activatefeature('SN1_60_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                elif stamping_press_type == 4:
                    mprog.activatefeature('SN1_80200_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                elif stamping_press_type == 5:
                    mprog.activatefeature('SN1_80200_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                elif stamping_press_type == 6:
                    mprog.activatefeature('SN1_80200_Body', 0)
                    mprog.activatefeature('Hole', 2)
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                elif stamping_press_type == 7:
                    mprog.activatefeature('SN1_80200_Body', 6)
                    mprog.activatefeature('Hole', 2)
                    mprog.partbodyfeatureactivate('SN1_25250_L')
                elif stamping_press_type == 8:
                    mprog.activatefeature('SN1_250_Body', 0)
                    mprog.activatefeature('Hole', 3)
                    mprog.partbodyfeatureactivate('SN1_25250_L')
            except:
                print('FRAME35 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME35 Update success')
                except:
                    print('FRAME35 Update error')
        elif name == 'FRAME36':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME36')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME36', stamping_press_type)
                print('FRAME36 Parameter change success')
            except:
                print('FRAME36 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('FRAME_SN1_25_Body')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('FRAME_SN1_3560_Body')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('FRAME_SN1_3560_Body')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('FRAME_SN1_3560_Body')
            except:
                print('FRAME36 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME36 Update success')
                except:
                    print('FRAME36 Update error')
        elif name == 'FRAME37':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME37')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME37', stamping_press_type)
                print('FRAME37 Parameter change success')
            except:
                print('FRAME37 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('FRAME_Body')
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('FRAME_Body')
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('FRAME_Body')
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('FRAME_Body')
            except:
                print('FRAME37 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME37 Update success')
                except:
                    print('FRAME37 Update error')
        elif name == 'FRAME38':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME38')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME38', stamping_press_type)
                print('FRAME38 Parameter change success')
            except:
                print('FRAME38 Parameter change error')
            try:
                if stamping_press_type == 0:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif stamping_press_type == 1:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif stamping_press_type == 2:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('ABD')
                    mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            except:
                print('FRAME38 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME38 Update success')
                except:
                    print('FRAME38 Update error')
        elif name == 'FRAME41':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME41')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME41', stamping_press_type)
                print('FRAME41 Parameter change success')
            except:
                print('FRAME41 Parameter change error')
            try:
                if stamping_press_type == 1:  # 35
                    pass
            except:
                print('FRAME41 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME41 Update success')
                except:
                    print('FRAME41 Update error')
        elif name == 'FRAME42':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME42')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME42', stamping_press_type)
                print('FRAME42 Parameter change success')
            except:
                print('FRAME42 Parameter change error')

            try:
                if stamping_press_type == 2:  # 45
                    pass
                elif stamping_press_type == 3:  # 60
                    pass
                elif stamping_press_type == 4:  # 80
                    pass
            except:
                print('FRAME42 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME42 Update success')
                except:
                    print('FRAME42 Update error')
        elif name == 'FRAME43':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME43')
            try:
                parameter_name, parameter_value = excel.get_sheet_par(
                    'FRAME43', stamping_press_type)
                print('FRAME43 Parameter change success')
            except BaseException:
                print('FRAME43 Parameter change error')
            try:
                if stamping_press_type == 4:
                    mprog.partbodyfeatureactivate('machining_Pocket')
                    mprog.activatefeature("Hole_1", 0)
                elif stamping_press_type == 3:
                    mprog.partbodyfeatureactivate('machining_Pocket')
                    mprog.activatefeature("Hole_1", 0)
                elif stamping_press_type == 5:
                    mprog.partbodyfeatureactivate('machining_Pocket')
                    mprog.activatefeature("Hole_1", 0)
                elif stamping_press_type == 6:
                    mprog.activatefeature("Hole_1", 0)
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('machining_Pocket')
                    mprog.activatefeature("Hole_1", 0)
                elif stamping_press_type == 8:
                    mprog.activatefeature("Hole_1", 0)
                else:
                    pass
            except BaseException:
                print('FRAME43 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME43 Update success')
                except BaseException:
                    print('FRAME43 Update error')
        elif name == 'FRAME45':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME45')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME45', stamping_press_type)
                print('FRAME45 Parameter change success')
            except:
                print('FRAME45 Parameter change error')
            try:
                if stamping_press_type == 3:  # 60N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
                elif stamping_press_type == 4:  # 80N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
                elif stamping_press_type == 5:  # 110N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
                elif stamping_press_type == 6:  # 160N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
                elif stamping_press_type == 7:  # 200N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
                elif stamping_press_type == 8:  # 250N
                    mprog.partbodyfeatureactivate('Chamfer.G')
                    mprog.partbodyfeatureactivate('Chamfer.F')
            except:
                print('FRAME45 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME45 Update success')
                except:
                    print('FRAME45 Update error')
        elif name == 'FRAME47':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME47')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME47', stamping_press_type)
                print('FRAME47 Parameter change success')
            except:
                print('FRAME47 Parameter change error')

            try:
                if stamping_press_type == 4:  # 80
                    pass
                elif stamping_press_type == 5:  # 110
                    pass
                elif stamping_press_type == 6:  # 160
                    pass
                elif stamping_press_type == 7:  # 200
                    pass
                elif stamping_press_type == 8:  # 250
                    pass
            except:
                print('FRAME47 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME47 Update success')
                except:
                    print('FRAME47 Update error')
        elif name == 'FRAME48':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME48')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME48', stamping_press_type)
                print('FRAME48 Parameter change success')
            except:
                print('FRAME48 Parameter change error')

            try:
                if stamping_press_type == 4:  # 80
                    pass
                elif stamping_press_type == 5:  # 110
                    pass
                elif stamping_press_type == 6:  # 160
                    pass
                elif stamping_press_type == 7:  # 200
                    pass
                elif stamping_press_type == 8:  # 250
                    pass
            except:
                print('FRAME48 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME48 Update success')
                except:
                    print('FRAME48 Update error')
        elif name == 'FRAME49':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME49')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME49', stamping_press_type)
                print('FRAME49 Parameter change success')
            except:
                print('FRAME49 Parameter change error')
            try:
                if stamping_press_type == 4:
                    mprog.activatefeature('body', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                    mprog.activatefeature('M12通', 0)
                elif stamping_press_type == 5:
                    mprog.activatefeature('body', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                    mprog.activatefeature('M12通', 0)
                elif stamping_press_type == 6:
                    mprog.activatefeature('body', 0)
                    mprog.activatefeature('M12通', 0)
                elif stamping_press_type == 7:
                    mprog.activatefeature('body', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                    mprog.activatefeature('M12通', 0)
                elif stamping_press_type == 8:
                    mprog.activatefeature('body', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                    mprog.activatefeature('M12通', 0)
            except:
                print('FRAME49 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME49 Update success')
                except:
                    print('FRAME49 Update error')
        elif name == 'FRAME50':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME50')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME50', stamping_press_type)
                print('FRAME50 Parameter change success')
            except:
                print('FRAME50 Parameter change error')
            try:
                if stamping_press_type == 4:
                    mprog.activatefeature('Body', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_c')
                    mprog.activatefeature('Hole_1', 0)
                elif stamping_press_type == 5:
                    mprog.activatefeature('Body', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_c')
                    mprog.activatefeature('Hole_1', 0)
                elif stamping_press_type == 6:
                    mprog.activatefeature('Body', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_c')
                    mprog.activatefeature('Hole_1', 0)
                elif stamping_press_type == 7:
                    mprog.activatefeature('Body', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_c')
                    mprog.activatefeature('Hole_1', 0)
                elif stamping_press_type == 8:
                    mprog.activatefeature('Body', 0)
                    mprog.partbodyfeatureactivate('FRAME_SN1_c')
                    mprog.activatefeature('Hole_1', 0)
            except:
                print('FRAME50 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME50 Update success')
                except:
                    print('FRAME50 Update error')
        elif name == 'FRAME51':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME51')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME51', stamping_press_type)
                print('FRAME51 Parameter change success')
            except:
                print('FRAME51 Parameter change error')
            # 判斷是否需要模墊加工
            try:
                if stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('SN1_250_CD')
                    mprog.activatefeature('M10*35L', 0)
                else:
                    print('FRAME51 feature activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME51 Update success')
                except:
                    print('FRAME51 Update error')
        elif name == 'FRAME51_1':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME51_1')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME51_1', stamping_press_type)
                print('FRAME51_1 Parameter change success')
            except:
                print('FRAME51_1 Parameter change error')
            # 判斷是否需要模墊加工
            try:
                if stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('Body')
                    mprog.partbodyfeatureactivate('SN1_250_CD')
                    mprog.activatefeature('M10*35L', 0)
                else:
                    print('FRAME51_1 feature activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME51_1 Update success')
                except:
                    print('FRAME51_1 Update error')
        elif name == 'FRAME52':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME52')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME52', stamping_press_type)
                print('FRAME52 Parameter change success')
            except:
                print('FRAME52 Parameter change error')
            # 判斷是否需要模墊加工
            try:
                if machiningdiepad == 0:
                    mprog.partbodyfeatureactivate('machining_除料')
                    mprog.activatefeature('machining_通孔', 0)
                    print('FRAME52 feature activate success')
                else:
                    print('FRAME52 feature activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME52 Update success')
                except:
                    print('FRAME52 Update error')
        elif name == 'FRAME53':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME53')
            try:
                parameter_name, parameter_value = excel.get_sheet_par(
                    'FRAME53', stamping_press_type)
                print('FRAME53 Parameter change success')
            except BaseException:
                print('FRAME53 Parameter change error')
            try:
                mprog.Update()
                print('FRAME53 Update success')
            except BaseException:
                print('FRAME53 Update error')
        elif name == 'FRAME53_1':
            excel = epc.ExcelOp('尺寸整理表', 'FRAME53_1')
            try:
                parameter_name, parameter_value = excel.get_sheet_par(
                    'FRAME53_1', stamping_press_type)
                print('FRAME53_1 Parameter change success')
            except BaseException:
                print('FRAME53_1 Parameter change error')
            try:
                mprog.Update()
                print('FRAME53_1 Update success')
            except BaseException:
                print('FRAME53_1 Update error')
        elif name == 'FRAME54':  # 已更改
            excel = epc.ExcelOp('尺寸整理表', 'FRAME54')
            try:
                parameter_name, parameter_value = excel.get_sheet_par('FRAME54', stamping_press_type)
                print('FRAME54 Parameter change success')
            except:
                print('FRAME54 Parameter change error')
            try:
                if stamping_press_type == 6:
                    mprog.partbodyfeatureactivate('Body')
                elif stamping_press_type == 7:
                    mprog.partbodyfeatureactivate('Body')
                elif stamping_press_type == 8:
                    mprog.partbodyfeatureactivate('Body')
            except:
                print('FRAME54 Parameter activate error')
            finally:
                try:
                    mprog.Update()
                    print('FRAME54 Update success')
                except:
                    print('FRAME54 Update error')
        elif name == "crankshaft":
            if travel_type == 1:
                excel = epc.ExcelOp('尺寸整理表', 'crankshaft_S')
                try:
                    parameter_name, parameter_value = excel.get_sheet_par('crankshaft', stamping_press_type)
                    print('crankshaft_S Parameter change success')
                except:
                    print('crankshaft_S Parameter change error')
                try:
                    if stamping_press_type == 0:  # 302CC7S
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外1')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_6M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 1:  # 322CC7S
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('BALLr外1')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_6M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 2:  # 342CC7S
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外2')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_6M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 3:  # 372CC7S
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外2')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_6M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 4:  # 395CC7S
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.partbodyfeatureactivate('CALLr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 5:  # 412CC7S
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外2')
                        mprog.partbodyfeatureactivate('BALLr內2')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.partbodyfeatureactivate('Cex(SN1_110_S)')
                        mprog.activatefeature('Cex孔(SN1_110_S)', 0)
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                        mprog.activatefeature('Dex(SN1_110_S)', 0)
                    elif stamping_press_type == 6:  # 435CC7S
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 7:  # 455CC7S
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 8:  # 475CC7S
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外2')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                except:
                    print('crankshaft Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        print('crankshaft Update success')
                    except:
                        print('crankshaft Update error')
            elif travel_type == 2:
                excel = epc.ExcelOp('尺寸整理表', 'crankshaft_H')
                try:
                    parameter_name, parameter_value = excel.get_sheet_par('crankshaft', stamping_press_type)
                    print('crankshaft_H Parameter change success')
                except:
                    print('crankshaft_H Parameter change error')
                try:
                    if stamping_press_type == 0:  # 302CC7H
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外1')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_6M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 1:  # 322CC7H
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('BALLr外1')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_6M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 2:  # 342CC7H
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外2')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_6M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 3:  # 372CC7H
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外1')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 4:  # 395CC7H
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.partbodyfeatureactivate('CALLr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 5:  # 415CC7H
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外2')
                        mprog.partbodyfeatureactivate('BALLr內2')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 6:  # 435CC7H
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 7:  # 455CC7H
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 8:  # 475CC7H
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外2')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                except:
                    print('crankshaft Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        print('crankshaft Update success')
                    except:
                        print('crankshaft Update error')
            elif travel_type == 3:
                excel = epc.ExcelOp('尺寸整理表', 'crankshaft_P')
                try:
                    parameter_name, parameter_value = excel.get_sheet_par('crankshaft', stamping_press_type)
                    print('crankshaft_P Parameter change success')
                except:
                    print('crankshaft_P Parameter change error')
                try:
                    if stamping_press_type == 0:  # 302CC7P
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外1')
                        mprog.partbodyfeatureactivate('BALLr內2')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_6M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 1:  # 322CC7P
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('BALLr外1')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_6M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 2:  # 342CC7P
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外2')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_6M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 3:  # 372CC7P
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Aex(SN1_60_P)')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外1')
                        mprog.partbodyfeatureactivate('BALLr內1')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_1M', 0)
                    elif stamping_press_type == 4:  # 395CC7P
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.partbodyfeatureactivate('CALLr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 5:    # 415CC7P
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外2')
                        mprog.partbodyfeatureactivate('BALLr內2')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 6:  # 435CC7P
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 7:  # 455CC7P
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('CALLr1')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                    elif stamping_press_type == 8:  # 475CC7P
                        mprog.partbodyfeatureactivate('Bex')
                        mprog.partbodyfeatureactivate('C5')
                        mprog.partbodyfeatureactivate('Ar1')
                        mprog.partbodyfeatureactivate('AALLr')
                        mprog.partbodyfeatureactivate('BALLr外2')
                        mprog.partbodyfeatureactivate('Cr2')
                        mprog.activatefeature('D1吊掛孔', 0)
                        mprog.activatefeature('D2_4M', 0)
                        mprog.activatefeature('D3_4M', 0)
                        mprog.activatefeature('D4_4M', 0)
                except:
                    print('crankshaft Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        print('crankshaft Update success')
                    except:
                        print('crankshaft Update error')
        elif name == "BALL_SCREW":
            if travel_type == 1:
                excel = epc.ExcelOp('尺寸整理表', 'BALL_SCREW_S')
                try:
                    parameter_name, parameter_value = excel.get_sheet_par('BALL_SCREW', stamping_press_type)
                    print('BALL_SCREW_S Parameter change success')
                except:
                    print('BALL_SCREW_S Parameter change error')
                try:
                    if stamping_press_type == 0:  # 302B07S02
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_T')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 1:  # 322B07S02
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_T')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 2:  # 342B07
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_T')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 3:  # 372B07
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 4:  # 392B07S06
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 5:  # 412B07
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 6:  # 432B07
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 7:  # 452B07S03
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('Groove_AY')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 8:  # 472B07S03
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('EdgeFillet_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                        mprog.partbodyfeatureactivate('EdgeFillet_AZ')
                except:
                    print('crankshaft_S Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        print('crankshaft_S Update success')
                    except:
                        print('crankshaft_S Update error')
            elif travel_type == 2:
                excel = epc.ExcelOp('尺寸整理表', 'BALL_SCREW_H')
                try:
                    parameter_name, parameter_value = excel.get_sheet_par('BALL_SCREW', stamping_press_type)
                    print('BALL_SCREW_H Parameter change success')
                except:
                    print('BALL_SCREW_H Parameter change error')
                try:
                    if stamping_press_type == 0:  # 302B07S03
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_T')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 1:  # 322B07S03
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_T')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 2:  # 342B07S01
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_T')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 3:  # 372B07S04
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 4:  # 392B07S07
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 5:  # 412B07S01
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 6:  # 432B07S01
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 7:  # 452B07S04
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('Groove_AY')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 8:  # 472B07S04
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('EdgeFillet_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                except:
                    print('BALL_SCREW_H Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        print('BALL_SCREW_H Update success')
                    except:
                        print('BALL_SCREW_H Update error')
            elif travel_type == 3:
                excel = epc.ExcelOp('尺寸整理表', 'BALL_SCREW_P')
                try:
                    parameter_name, parameter_value = excel.get_sheet_par('BALL_SCREW', stamping_press_type)
                    print('BALL_SCREW_P Parameter change success')
                except:
                    print('BALL_SCREW_P Parameter change error')
                try:
                    if stamping_press_type == 0:  # 302B07S03
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_T')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 1:  # 322B07S03
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_T')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 2:  # 342B07S01
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_T')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 3:  # 372B07S02
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 4:  # 392B07S08
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 5:    # 412B07S02
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 6:  # 432B07S02
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 7:  # 452B07S05
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('Groove_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('Groove_AY')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                    elif stamping_press_type == 8:  # 472B07S05
                        mprog.partbodyfeatureactivate('BALL_SCREW_Body')
                        mprog.partbodyfeatureactivate('BALL_SCREW_Thread')
                        mprog.partbodyfeatureactivate('AE')
                        mprog.partbodyfeatureactivate('AB')
                        mprog.partbodyfeatureactivate('Hole_AAAC')
                        mprog.partbodyfeatureactivate('EdgeFillet_O')
                        mprog.partbodyfeatureactivate('Chamfer_P')
                        mprog.partbodyfeatureactivate('EdgeFillet_V')
                        mprog.partbodyfeatureactivate('EdgeFillet_X')
                        mprog.partbodyfeatureactivate('EdgeFillet_Y')
                        mprog.partbodyfeatureactivate('EdgeFillet_AD')
                        mprog.partbodyfeatureactivate('Chamfer_AUAV')
                        mprog.partbodyfeatureactivate('EdgeFillet_AW')
                        mprog.partbodyfeatureactivate('EdgeFillet_AX')
                        mprog.partbodyfeatureactivate('AFAH')
                        mprog.partbodyfeatureactivate('Hole_AJ')
                except:
                    print('BALL_SCREW_P Parameter activate error')
                finally:
                    try:
                        mprog.Update()
                        print('BALL_SCREW_P Update success')
                    except:
                        print('BALL_SCREW_P Update error')



    # print出跳錯欄位
    except Exception as e:
        s = sys.exc_info()
        print('報錯行數：{}\n報錯內容：{}'.format(__file__, s[2].tb_lineno, s[1]))

    return parameter_name, parameter_value