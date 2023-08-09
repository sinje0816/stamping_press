import main_program as mprog
import excel_parameter_change as epc
import time


# 母檔變數變換


def change_machining_parameter(name, i, machiningdiepad):
    if name == 'FRAME1':
        excel = epc.ExcelOp('FRAME1')
        try:
            FRAME1_parameter_name, FRAME1_parameter_value = excel.part_parameter(
                'FRAME1', i)
            print('FRAME1 Parameter change success')
        except BaseException:
            print('FRAME1 Parameter change error')
        if i == 0:
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
        elif i == 1:
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
        elif i == 2:
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
        elif i == 3:
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
        elif i == 4:
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
        elif i == 5:
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
        elif i == 6:
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
        elif i == 7:
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
        elif i == 8:
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
        try:
            mprog.update()
            print('FRAME1 update success')
        except BaseException:
            print('FRAME1 update error')
    elif name == 'FRAME2':
        excel = epc.ExcelOp('FRAME2')
        try:
            FRAME2_parameter_name, FRAME2_parameter_value = excel.part_parameter(
                'FRAME2', i)
            print('FRAME2 Parameter change success')
        except BaseException:
            print('FRAME2 Parameter change error')
        try:
            if i == 0:
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
            elif i == 1:
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
            elif i == 2:
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
            elif i == 3:
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
            elif i == 4:
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
            elif i == 5:
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
            elif i == 6:
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
            elif i == 7:
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
            elif i == 8:
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
        except BaseException:
            print('FRAME2 Part activate error')
        finally:
            mprog.update()
            print('FRAME2 update success')
    elif name == 'FRAME3':
        excel = epc.ExcelOp('FRAME3')
        try:
            excel.part_parameter('FRAME3', i)
            print('FRAME3 Parameter change success')
        except:
            print('FRAME3 Parameter change error')
        try:
            if i == 0:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
            elif i == 1:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')

            elif i == 2:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
            elif i == 3:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_S')
                mprog.partbodyfeatureactivate('FRAME_SN1_2560_U')
            elif i == 4:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
            elif i == 5:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
            elif i == 6:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
            elif i == 7:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
            elif i == 8:
                mprog.activatefeature('SN1_25250_Body', 2)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_25250_PQ')
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_R')
        except:
            print('FRAME3 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME3 Update success')
            except:
                print('FRAME3 Update error')
    elif name == 'FRAME4':  # 已更改
        excel = epc.ExcelOp('FRAME4')
        try:
            excel.part_parameter('FRAME4', i)
            print('FRAME4 Parameter change success')
        except:
            print('FRAME4 Parameter change error')
        try:
            if i == 0:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
            elif i == 1:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
            elif i == 2:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
            elif i == 3:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
            elif i == 4:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i == 5:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i == 6:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i == 7:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 2)
                mprog.activatefeature('Hole_4', 0)
            elif i == 8:
                mprog.activatefeature('SN1_25250_Body', 1)
                mprog.activatefeature('Hole_1', 1)
                mprog.activatefeature('CLUCTH_AIR_HOLE', 0)
                mprog.activatefeature('Hole_3', 4)
                mprog.activatefeature('Hole_4', 0)
        except:
            print('FRAME4 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME4 Update success')
            except:
                print('FRAME4 Update error')
    elif name == 'FRAME5':
        excel = epc.ExcelOp('FRAME5')
        try:
            FRAME5_parameter_name, FRAME5_parameter_value = excel.part_parameter(
                'FRAME5', i)
            print('FRAME5 Parameter change success')
        except BaseException:
            print('FRAME5 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('SN1-25_X')
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('Y')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
                mprog.activatefeature('PT1/2', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('Y')
                mprog.partbodyfeatureactivate('AA')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
                mprog.activatefeature('PT1/2', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('Y')
                mprog.partbodyfeatureactivate('AA')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
                mprog.activatefeature('PT1/2', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('before80_I')
                mprog.partbodyfeatureactivate('before80_R')
                mprog.partbodyfeatureactivate('S')
                mprog.partbodyfeatureactivate('Rx2')
                mprog.partbodyfeatureactivate('SN1-60_R')
                mprog.partbodyfeatureactivate('Chamfer.2')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.activatefeature('2-M8通孔', 0)
                mprog.activatefeature('TY牌電磁閥用', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 5:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('6-M6通C5')
                mprog.activatefeature('6-M6通', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('3-M8通_C5')
                mprog.activatefeature('3-M8通', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
                mprog.partbodyfeatureactivate('5-M8通_C5')
                mprog.activatefeature('5-M8通', 0)
                mprog.activatefeature('M6通孔LED', 0)
            elif i == 8:
                mprog.partbodyfeatureactivate('after80_S')
                mprog.partbodyfeatureactivate('after80_R')
                mprog.partbodyfeatureactivate('Chamfer.3')
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
        excel = epc.ExcelOp('FRAME6')
        try:
            excel.part_parameter('FRAME6', i)
            print('FRAME6 Parameter change success')
        except:
            print('FRAME6 Parameter change error')
        try:
            if i == 0:
                mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                mprog.activatefeature('M16_i', 0)
            elif i == 1:
                mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                mprog.activatefeature('M16_i', 0)
            elif i == 2:
                mprog.activatefeature('FRAME_SN1_2545_Body', 0)
                mprog.activatefeature('M16_i', 0)
            elif i == 3:
                mprog.activatefeature('FRAME_SN1_60_Body', 0)
                mprog.activatefeature('M16_h FRAME9_L', 0)
            elif i == 4:
                mprog.activatefeature('FRAME_SN1_80110_Body', 0)
                mprog.activatefeature('M16_h FRAME9_L', 0)
            elif i == 5:
                mprog.activatefeature('FRAME_SN1_80110_Body', 0)
                mprog.activatefeature('M16_h FRAME9_L', 0)
            elif i == 7:
                mprog.activatefeature('FRAME_SN1_200_Body', 0)
                mprog.activatefeature('M16_h FRAME9_L', 0)
        except:
            print('FRAME6 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME6 Update success')
            except:
                print('FRAME6 Update error')
    elif name == 'FRAME7':  # 已更改
        excel = epc.ExcelOp('FRAME7')
        try:
            excel.part_parameter('FRAME7', i)
            print('FRAME7 Parameter change success')
        except:
            print('FRAME7 Parameter change error')
        try:
            if i == 0:
                mprog.activatefeature('SN1_2560_Body', 1)
                mprog.partbodyfeatureactivate('SN1_25_VWX')
            elif i == 1:
                mprog.activatefeature('SN1_2560_Body', 1)
            elif i == 2:
                mprog.activatefeature('SN1_2560_Body', 1)
            elif i == 3:
                mprog.activatefeature('SN1_2560_Body', 1)
            elif i == 4:
                mprog.activatefeature('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i == 5:
                mprog.activatefeature('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i == 6:
                mprog.activatefeature('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i == 7:
                mprog.activatefeature('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
            elif i == 8:
                mprog.activatefeature('FRAME_SN1_80250_Body', 4)
                mprog.activatefeature('Hole_1', 0)
        except:
            print('FRAME7 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME7 Update success')
            except:
                print('FRAME7 Update error')
    elif name == 'FRAME8':
        excel = epc.ExcelOp('FRAME8')
        # try:
        FRAME8_parameter_name, FRAME8_parameter_value = excel.part_parameter(
            'FRAME8', i)
        #     print('FRAME8 Parameter change success')
        # except:
        #     print('FRAME8 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('上半倒角')
            elif i == 1:
                mprog.partbodyfeatureactivate('上半倒角')
                # mprog.activatefeature('38孔通', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('上半倒角')
            elif i == 3:
                mprog.partbodyfeatureactivate('上半倒角')
            elif i == 4:
                mprog.partbodyfeatureactivate('上半倒角80')
                mprog.partbodyfeatureactivate('R40特徵')
                mprog.partbodyfeatureactivate('下半開槽')
                mprog.activatefeature('38孔通', 0)
            elif i == 5:
                mprog.partbodyfeatureactivate('上半倒角')
                mprog.partbodyfeatureactivate('下半倒角')
                mprog.partbodyfeatureactivate('下半開槽')
                mprog.partbodyfeatureactivate('R40特徵')
                mprog.activatefeature('38孔通', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('上半倒角')
                mprog.partbodyfeatureactivate('下半倒角')
                mprog.partbodyfeatureactivate('下半開槽')
                mprog.partbodyfeatureactivate('R40特徵')
                mprog.activatefeature('38孔通', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('上半倒角')
                mprog.partbodyfeatureactivate('下半開槽')
                mprog.partbodyfeatureactivate('R40特徵')
                mprog.activatefeature('38孔通', 0)
            elif i == 8:
                mprog.partbodyfeatureactivate('上半倒角')
                mprog.partbodyfeatureactivate('R40特徵')
                mprog.activatefeature('38孔通', 0)
        except BaseException:
            print('FRAME8 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME8 Update success')
            except BaseException:
                print('FRAME8 Update error')
    elif name == 'FRAME9':
        excel = epc.ExcelOp('FRAME9')
        try:
            excel.part_parameter('FRAME9', i)
            print('FRAME9 Parameter change success')
        except:
            print('FRAME9 Parameter change error')
        try:
            if i == 0:
                mprog.activatefeature('SN1_2580_Body', 1)
            elif i == 1:
                mprog.activatefeature('SN1_2580_Body', 1)
            elif i == 2:
                mprog.activatefeature('SN1_2580_Body', 1)
            elif i == 3:
                mprog.activatefeature('SN1_2580_Body', 1)
            elif i == 4:
                mprog.activatefeature('channel steel', 0)
            elif i == 5:
                mprog.activatefeature('channel steel', 0)
            elif i == 6:
                mprog.activatefeature('channel steel', 0)
            elif i == 7:
                mprog.activatefeature('channel steel', 0)
            elif i == 8:
                mprog.activatefeature('channel steel', 0)
        except:
            print('FRAME9 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME9 Update success')
            except:
                print('FRAME9 Update error')
    elif name == 'FRAME10':  # 已更改
        excel = epc.ExcelOp('FRAME10')
        try:
            excel.part_parameter('FRAME10', i)
            print('FRAME10 Parameter change success')
        except:
            print('FRAME10 Parameter change error')
        try:
            if i == 0:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i == 1:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i == 2:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i == 3:
                mprog.activatefeature('SN1_60_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i == 4:
                mprog.activatefeature('SN1_80110_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i == 5:
                mprog.activatefeature('SN1_80110_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i == 6:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i == 7:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
            elif i == 8:
                mprog.activatefeature('SN1_160250_Body', 0)
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('SN1_25250_K')
        except:
            print('FRAME10 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME10 Update success')
            except:
                print('FRAME10 Update error')
    elif name == 'FRAME11':
        start = time.time()
        excel = epc.ExcelOp('FRAME11')
        try:
            FRAME11_parameter_name, FRAME11_parameter_value = excel.part_parameter(
                'FRAME11', i)
            print('FRAME11 Parameter change success')
        except BaseException:
            print('FRAME11 Parameter change error')
        if i == 0:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 1:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 2:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 3:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 4:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 5:
            mprog.activatefeature('挖孔', 0)
            mprog.activatefeature('RC1/8', 0)
        elif i == 6:
            mprog.activatefeature('挖孔', 0)
        elif i == 7:
            pass
        elif i == 8:
            pass
        try:
            mprog.Update()
            print('FRAME11 Update success')
        except BaseException:
            print('FRAME11 Update error')
    elif name == 'FRAME12':
        excel = epc.ExcelOp('FRAME12')
        try:
            excel.part_parameter('FRAME12', i)
            FRAME12_parameter_name, FRAME12_parameter_value = excel.part_parameter('FRAME12', i)
            print('FRAME12 Parameter change success')
        except:
            print('FRAME12 Parameter change error')
        try:
            if i == 0:  # 25N
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)
            elif i == 1:  # 35N
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)
            elif i == 2:  # 45N
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)
            elif i == 3:  # 60N
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)
            elif i == 4:  # 80N
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)
            elif i == 5:  # 110N
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)
            elif i == 6:  # 160N
                mprog.partbodyfeatureactivate('C8_160')
                mprog.activatefeature('Pipe_clamp', 2)
                mprog.activatefeature('drain_hole', 0)
            elif i == 7:  # 200N
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.activatefeature('Pipe_clamp', 0)
                mprog.activatefeature('drain_hole', 0)
            elif i == 8:  # 250N
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
        excel = epc.ExcelOp('FRAME13')
        try:
            FRAME13_parameter_name, FRAME13_parameter_value = excel.part_parameter('FRAME13', i)
            print('FRAME13 Parameter change success')
        except:
            print('FRAME13 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_25_45_CD')
                mprog.partbodyfeatureactivate('SN1_25_45_G')
                mprog.activatefeature('GIB_OIL_HOLE', 1)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 1)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_25_45_CD')
                mprog.partbodyfeatureactivate('SN1_25_45_G')
                mprog.activatefeature('GIB_OIL_HOLE', 1)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 1)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 5:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('GIB_OIL_HOLE', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('Hole_4', 0)
            elif i == 8:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_1', 0)
                mprog.activatefeature('SLIDE_LINER_FIX_HOLE_2', 0)
                mprog.activatefeature('GIB_OIL_HOLE', 0)
                mprog.activatefeature('Hole_4', 0)
        except:
            print('FRAME13 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME13 Update success')
            except:
                print('FRAME13 Update error')
    elif name == 'FRAME14':
        excel = epc.ExcelOp('FRAME14')
        try:
            excel.part_parameter('FRAME14', i)
            FRAME14_parameter_name, FRAME14_parameter_value = excel.part_parameter('FRAME14', i)
            print('FRAME14 Parameter change success')
        except:
            print('FRAME14 Parameter change error')
        try:
            if i == 0:  # 25N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
            elif i == 1:  # 35N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
            elif i == 2:  # 45N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
            elif i == 3:  # 60N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)
            elif i == 4:  # 80N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('outside_hole_E', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)
            elif i == 5:  # 110N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)
            elif i == 6:  # 160N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)
            elif i == 7:  # 200N
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('outside_hole_E', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)
            elif i == 8:  # 250N
                mprog.partbodyfeatureactivate('250N')
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)
        except:
            print('FRAME14 Parameter activate error')
        finally:
            try:
                # mprog.Update()
                print('FRAME14 Update success')
            except:
                print('FRAME14 Update error')

        #######################################################################################################
    elif name == 'FRAME15':
        excel = epc.ExcelOp('FRAME15')
        try:
            excel.part_parameter('FRAME15', i)
            FRAME15_parameter_name, FRAME15_parameter_value = excel.part_parameter('FRAME15', i)
            print('FRAME15 Parameter change success')
        except:
            print('FRAME15 Parameter change error')
        try:
            if i == 0:  # 25N
                mprog.activatefeature('Hole', 0)
            elif i == 1:  # 35N
                mprog.activatefeature('Hole', 0)
            elif i == 2:  # 45N
                mprog.activatefeature('Hole', 0)
            elif i == 3:  # 60N
                mprog.activatefeature('Hole', 0)
            elif i == 4:  # 80N
                mprog.activatefeature('Hole', 0)
            elif i == 5:  # 110N
                mprog.activatefeature('Hole', 0)
            elif i == 6:  # 160N
                mprog.activatefeature('Hole', 0)
            elif i == 7:  # 200N
                mprog.activatefeature('Hole', 0)
            elif i == 8:  # 250N
                mprog.partbodyfeatureactivate('250N')
                mprog.activatefeature('Hole', 0)
        except:
            print('FRAME15 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME15 Update success')
            except:
                print('FRAME15 Update error')
    elif name == 'FRAME17':
        excel = epc.ExcelOp('FRAME17')
        try:
            excel.part_parameter('FRAME17', i)
            FRAME17_parameter_name, FRAME17_parameter_value = excel.part_parameter('FRAME17', i)
            print('FRAME17 Parameter change success')
        except:
            print('FRAME17 Parameter change error')
        try:
            if i == 0:  # 25N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)

            elif i == 1:  # 35N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)

            elif i == 2:  # 45N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)

            elif i == 3:  # 60N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)

            elif i == 4:  # 80N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('outside_hole_E', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)

            elif i == 5:  # 110N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)

            elif i == 6:  # 160N
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)

            elif i == 7:  # 200N
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('outside_hole_E', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)

            elif i == 8:  # 250N
                mprog.partbodyfeatureactivate('250N')
                mprog.activatefeature('inside_hole', 0)
                mprog.activatefeature('outside_hole', 0)
                mprog.activatefeature('processing_h', 0)
                mprog.activatefeature('processing_i', 0)
                mprog.activatefeature('processing_m', 0)
        except:
            print('FRAME17 Parameter activate error')
        finally:
            try:
                # mprog.Update()
                print('FRAME17 Update success')
            except:
                print('FRAME17 Update error')
    elif name == 'FRAME18':  # Add el
        excel = epc.ExcelOp('FRAME18')
        try:
            excel.part_parameter('FRAME18', i)
            FRAME18_parameter_name, FRAME18_parameter_value = excel.part_parameter('FRAME18', i)
            print('FRAME18 Parameter change success')
        except:
            print('FRAME18 Parameter change error')

        try:
            if i == 0:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.partbodyfeatureactivate('HollowCylinder')
            elif i == 1:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.partbodyfeatureactivate('HollowCylinder')
            elif i == 2:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.partbodyfeatureactivate('HollowCylinder')
            elif i == 3:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.partbodyfeatureactivate('HollowCylinder')
            elif i == 4:
                mprog.partbodyfeatureactivate('Pad.2')
                mprog.partbodyfeatureactivate('HollowCylinder')
            elif i == 5:
                mprog.activatefeature('Hole.1', 0)
                mprog.partbodyfeatureactivate('EdgeFillet.1')
            elif i == 6:
                mprog.activatefeature('Hole.1', 0)
                mprog.partbodyfeatureactivate('EdgeFillet.1')
            elif i == 7:
                mprog.activatefeature('Hole.1', 0)
                mprog.partbodyfeatureactivate('EdgeFillet.1')
            elif i == 8:
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
        excel = epc.ExcelOp('FRAME19')
        try:
            excel.part_parameter('FRAME19', i)
            print('FRAME19 Parameter change success')
        except:
            print('FRAME19 Parameter change error')
        try:
            if i == 0:
                mprog.activatefeature('SN1_2545_Body', 6)
                mprog.partbodyfeatureactivate('SN1_25_X')
                mprog.partbodyfeatureactivate('SN1_2535_Y')
            elif i == 1:
                mprog.activatefeature('SN1_2545_Body', 6)
                mprog.partbodyfeatureactivate('SN1_2535_Y')
                mprog.partbodyfeatureactivate("SN1_3545_AD")
            elif i == 2:
                mprog.activatefeature('SN1_2545_Body', 6)
                mprog.partbodyfeatureactivate('SN1_45_Y')
                mprog.partbodyfeatureactivate('SN1_45_X')
                mprog.partbodyfeatureactivate("SN1_3545_AD")
            elif i == 3:
                mprog.activatefeature('SN1_60_Body', 8)
            elif i == 4:
                mprog.activatefeature('SN1_80110_Body', 5)
            elif i == 5:
                mprog.activatefeature('SN1_80110_Body', 5)
            elif i == 6:
                mprog.activatefeature('FRMAE_SN1_160_Body', 0)
                mprog.activatefeature('FRAME_34_Hole_1', 0)
            elif i == 7:
                mprog.activatefeature('SN1_200_Body', 4)
            elif i == 8:
                mprog.activatefeature('SN1_250_Body', 0)
        except:
            print('FRAME19 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME19 Update success')
            except:
                print('FRAME19 Update error')
    elif name == 'FRAME20':
        excel = epc.ExcelOp('FRAME20')
        try:
            excel.part_parameter('FRAME20', i)
            print('FRAME20 Parameter change success')
        except:
            print('FRAME20 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_45250_C')
                mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.partbodyfeatureactivate('SN1_80250_DEF')
                mprog.partbodyfeatureactivate('SN1_80250_G')
                mprog.activatefeature('MOTOR_FIX_ADJUSTMENT_HOLE', 0)
            elif i == 5:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.partbodyfeatureactivate('SN1_80250_DEF')
                mprog.partbodyfeatureactivate('SN1_80250_G')
                mprog.partbodyfeatureactivate('SN1_110160_ijk')
                mprog.activatefeature('Hole_2', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.partbodyfeatureactivate('SN1_80250_DEF')
                mprog.partbodyfeatureactivate('SN1_80250_G')
                mprog.partbodyfeatureactivate('SN1_110160_ijk')
                mprog.activatefeature('Hole_2', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60250_C')
                mprog.partbodyfeatureactivate('SN1_80250_DEF')
                mprog.partbodyfeatureactivate('SN1_80250_G')
                mprog.partbodyfeatureactivate('SN1_200250_hij')
                mprog.activatefeature('Hole_2', 1)
            elif i == 8:
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
        excel = epc.ExcelOp('FRAME21')
        try:
            excel.part_parameter('FRAME21', i)
            FRAME21_parameter_name, FRAME21_parameter_value = excel.part_parameter('FRAME21', i)
            print('FRAME21 Parameter change success')
        except:
            print('FRAME21 Parameter change error')

        try:
            if i == 0:  # 25
                mprog.activatefeature('Hole', 0)
            elif i == 1:  # 35
                mprog.activatefeature('Hole', 0)
        except:
            print('FRAME21 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME21 Update success')
            except:
                print('FRAME21 Update error')
    elif name == 'FRAME22':  # 已更改
        excel = epc.ExcelOp('FRAME22')
        try:
            excel.part_parameter('FRAME22', i)
            print('FRAME22 Parameter change success')
        except:
            print('FRAME22 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('A(1)D')
                mprog.partbodyfeatureactivate('SN1-2560_BE')
                mprog.partbodyfeatureactivate('SN1_2545_I')
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i == 3:
                mprog.activatefeature('FRAME_SN1_60_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('Hole_3', 0)
                mprog.activatefeature('SN1_2560_6_M8X16L_底孔27L', 0)
            elif i == 4:
                mprog.activatefeature('FRAME_SN1_80250_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i == 5:
                mprog.activatefeature('FRAME_SN1_80250_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i == 6:
                mprog.activatefeature('FRAME_SN1_80250_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i == 7:
                mprog.activatefeature('FRAME_SN1_80250_Body', 0)
                mprog.activatefeature('OPERATION_BOX_FIX_HOLE', 0)
                mprog.activatefeature('SN1_80250_6_M8X16L_底孔27L', 0)
                mprog.activatefeature('Hole_3', 0)
            elif i == 8:
                mprog.activatefeature('FRAME_SN1_80250_Body', 0)
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
        excel = epc.ExcelOp('FRAME23')
        try:
            excel.part_parameter('FRAME23', i)
            FRAME23_parameter_name, FRAME23_parameter_value = excel.part_parameter('FRAME23', i)
            print('FRAME23 Parameter change success')
        except:
            print('FRAME23 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('D1')
                mprog.partbodyfeatureactivate('D2')
                mprog.partbodyfeatureactivate('J1')
                mprog.partbodyfeatureactivate('J2')
                mprog.partbodyfeatureactivate('25至60除料')
                mprog.partbodyfeatureactivate('25至60導圓角')
            elif i == 1:
                mprog.partbodyfeatureactivate('D1')
                mprog.partbodyfeatureactivate('D2')
                mprog.partbodyfeatureactivate('J1')
                mprog.partbodyfeatureactivate('J2')
                mprog.partbodyfeatureactivate('25至60除料')
                mprog.partbodyfeatureactivate('25至60導圓角')
            elif i == 2:
                mprog.partbodyfeatureactivate('D1')
                mprog.partbodyfeatureactivate('D2')
                mprog.partbodyfeatureactivate('J1')
                mprog.partbodyfeatureactivate('J2')
                mprog.partbodyfeatureactivate('25至60除料')
                mprog.partbodyfeatureactivate('25至60導圓角')
            elif i == 3:
                mprog.partbodyfeatureactivate('D1')
                mprog.partbodyfeatureactivate('D2')
                mprog.partbodyfeatureactivate('J1')
                mprog.partbodyfeatureactivate('J2')
                mprog.partbodyfeatureactivate('25至60除料')
                mprog.partbodyfeatureactivate('25至60導圓角')
            elif i == 4:
                mprog.partbodyfeatureactivate('D1')
                mprog.partbodyfeatureactivate('D2')
                mprog.partbodyfeatureactivate('J1')
                mprog.partbodyfeatureactivate('J2')
                mprog.partbodyfeatureactivate('80至250除料')
                mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
            elif i == 5:
                mprog.partbodyfeatureactivate('D1')
                mprog.partbodyfeatureactivate('D2')
                mprog.partbodyfeatureactivate('J1')
                mprog.partbodyfeatureactivate('J2')
                mprog.partbodyfeatureactivate('80至250除料')
                mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
            elif i == 6:
                mprog.partbodyfeatureactivate('D1')
                mprog.partbodyfeatureactivate('D2')
                mprog.partbodyfeatureactivate('J1')
                mprog.partbodyfeatureactivate('J2')
                mprog.partbodyfeatureactivate('80至250除料')
            elif i == 7:
                mprog.partbodyfeatureactivate('D1')
                mprog.partbodyfeatureactivate('D2')
                mprog.partbodyfeatureactivate('J1')
                mprog.partbodyfeatureactivate('J2')
                mprog.partbodyfeatureactivate('80至250除料')
            elif i == 8:
                mprog.partbodyfeatureactivate('D1')
                mprog.partbodyfeatureactivate('D2')
                mprog.partbodyfeatureactivate('J1')
                mprog.partbodyfeatureactivate('J2')
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
    elif name == 'FRAME24':
        excel = epc.ExcelOp('FRAME24')
        try:
            FRAME24_parameter_name, FRAME24_parameter_value = excel.part_parameter(
                'FRAME24', i)
            print('FRAME24 Parameter change success')
        except BaseException:
            print('FRAME24 Parameter change error')
        if i == 0:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 1:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 2:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 3:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 4:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 5:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 6:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 7:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        elif i == 8:
            mprog.partbodyfeatureactivate('倒角C')
            mprog.activatefeature('M5', 0)
        try:
            mprog.Update()
            print('FRAME24 Update success')
        except BaseException:
            print('FRAME24 Update error')
    elif name == 'FRAME25':
        excel = epc.ExcelOp('FRAME25')
        try:
            excel.part_parameter('FRAME25', i)
            print('FRAME25 Parameter change success')
        except:
            print('FRAME25 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('3_M8通', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('3_M8通', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.activatefeature('3_M8通', 0)
            elif i == 3:
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
        excel = epc.ExcelOp('FRAME26')
        try:
            excel.part_parameter('FRAME26', i)
            FRAME26_parameter_name, FRAME26_parameter_value = excel.part_parameter('FRAME26', i)
            print('FRAME26 Parameter change success')
        except:
            print('FRAME26 Parameter change error')

        try:
            if i == 0:
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 1:
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 2:
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 3:
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 4:
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 5:
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.partbodyfeatureactivate('F')
            elif i == 6:
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 7:
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 8:
                mprog.partbodyfeatureactivate('Chamfer.1')
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
        excel = epc.ExcelOp('FRAME27')

        try:
            excel.part_parameter('FRAME27', i)
            FRAME27_parameter_name, FRAME27_parameter_value = excel.part_parameter('FRAME27', i)
            print('FRAME27 Parameter change success')
        except:
            print('FRAME27 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('35N前')
                mprog.activatefeature('Hole', 0)
                mprog.activatefeature('Piping', 0)
                mprog.partbodyfeatureactivate('machining_aaaaa2')
            elif i == 1:
                mprog.partbodyfeatureactivate('35N前')
                mprog.activatefeature('Hole', 0)
                mprog.activatefeature('Piping', 0)
                mprog.partbodyfeatureactivate('machining_aaaaa2')
            elif i == 2:
                mprog.partbodyfeatureactivate('Pocket.1')
                mprog.activatefeature('Hole', 0)
                mprog.activatefeature('Piping', 0)
                mprog.activatefeature('machining_ddddd', 0)
                mprog.partbodyfeatureactivate('machining_aaaaa2')
            elif i == 3:
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
        excel = epc.ExcelOp('FRAME27_1')
        try:
            excel.part_parameter('FRAME27_1', i)
            FRAME27_1_parameter_name, FRAME27_1_parameter_value = excel.part_parameter('FRAME27_1', i)
            print('FRAME27_1 Parameter change success')
        except:
            print('FRAME27_1 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('35N前')
                mprog.activatefeature('Hole', 0)
                mprog.activatefeature('Piping', 0)
                mprog.partbodyfeatureactivate('machining_aaaaa2')
            elif i == 1:
                mprog.partbodyfeatureactivate('35N前')
                mprog.activatefeature('Hole', 0)
                mprog.activatefeature('Piping', 0)
                mprog.partbodyfeatureactivate('machining_aaaaa2')
            elif i == 2:
                mprog.partbodyfeatureactivate('Pocket.1')
                mprog.activatefeature('Hole', 0)
                mprog.activatefeature('Piping', 0)
                mprog.activatefeature('machining_ddddd', 0)
                mprog.partbodyfeatureactivate('machining_aaaaa2')
            elif i == 3:
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
        excel = epc.ExcelOp('FRAME28')
        try:
            excel.part_parameter('FRAME28', i)
            print('FRAME28 Parameter change success')
        except:
            print('FRAME28 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i == 1:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i == 2:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i == 3:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i == 4:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i == 5:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i == 6:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i == 7:
                mprog.partbodyfeatureactivate('SN1_25250_ABD')
                mprog.partbodyfeatureactivate('SN1_25250_E')
                mprog.partbodyfeatureactivate('SN1_25250_F')
            elif i == 8:
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
        excel = epc.ExcelOp('FRAME29')
        try:
            FRAME29_parameter_name, FRAME29_parameter_value = excel.part_parameter(
                'FRAME29', i)
            print('FRAME29 Parameter change success')
        except BaseException:
            print('FRAME29 Parameter change error')
        try:
            mprog.partbodyfeatureactivate('Hole.1')
        finally:
            try:
                mprog.Update()
                print('FRAME29 Update success')
            except BaseException:
                print('FRAME29 Update error')
    elif name == 'FRAME30':  # 已更改
        excel = epc.ExcelOp('FRAME30')
        try:
            excel.part_parameter('FRAME30', i)
            print('FRAME30 Parameter change success')
        except:
            print('FRAME30 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i == 5:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('SN1_25250_AE')
                mprog.partbodyfeatureactivate('SN1_25250_L')
                mprog.partbodyfeatureactivate('SN1_25250_HIJ')
                mprog.activatefeature('FRAME_CRANK_SHAFT', 0)
                mprog.activatefeature('Second_degree_drop_wire_threading_hole', 0)
                mprog.activatefeature('4_M14通', 0)
            elif i == 8:
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
        excel = epc.ExcelOp('FRAME31')
        try:
            excel.part_parameter('FRAME31', i)
            print('FRAME31 Parameter change success')
        except:
            print('FRAME31 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('Body')
                mprog.activatefeature('FRAME_Body', 2)
            elif i == 1:
                mprog.partbodyfeatureactivate('Body')
                mprog.activatefeature('FRAME_Body', 0)
            elif i == 2:
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
        excel = epc.ExcelOp('FRAME32')
        try:
            excel.part_parameter('FRAME32', i)
            print('FRAME32 Parameter change success')
        except:
            print('FRAME32 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('Body')
            elif i == 1:
                mprog.partbodyfeatureactivate('Body')
            elif i == 2:
                mprog.partbodyfeatureactivate('Body')
            elif i == 3:
                mprog.partbodyfeatureactivate('Body')
            elif i == 4:
                mprog.partbodyfeatureactivate('Body')
            elif i == 5:
                mprog.partbodyfeatureactivate('Body')
            elif i == 6:
                mprog.partbodyfeatureactivate('Body')
            elif i == 7:
                mprog.partbodyfeatureactivate('Body')
            elif i == 8:
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
        excel = epc.ExcelOp('FRAME33')
        try:
            excel.part_parameter('FRAME33', i)
            print('FRAME33 Parameter change success')
        except:
            print('FRAME33 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('FRAME_2545_LMN')
            elif i == 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('FRAME_2545_LMN')
            elif i == 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CEF')
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('FRAME_2545_LMN')
            elif i == 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.activatefeature('Hole_1', 0)
                mprog.partbodyfeatureactivate('FRAME_60_LMN')
        except:
            print('FRAME33 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME33 Update success')
            except:
                print('FRAME33 Update error')
    elif name == 'FRAME34':
        excel = epc.ExcelOp('FRAME34')
        try:
            excel.part_parameter('FRAME34', i)
            print('FRAME34 Parameter change success')
        except:
            print('FRAME34 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.partbodyfeatureactivate('SN1_2535_J')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_1_2545_PQR')
            elif i == 1:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.partbodyfeatureactivate('SN1_2535_J')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_1_2545_PQR')
            elif i == 2:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_2545_CD')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_1_2545_PQR')
            elif i == 3:
                mprog.partbodyfeatureactivate('AB')
                mprog.partbodyfeatureactivate('SN1_60_NM')
                mprog.activatefeature('Air_pipe_wire_threading_hole', 0)
                mprog.activatefeature('Hole_2', 0)
                mprog.partbodyfeatureactivate('FRAME_1_60_PQR')
        except:
            print('FRAME34 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME34 Update success')
            except:
                print('FRAME34 Update error')
    elif name == 'FRAME35':  # 已更改
        excel = epc.ExcelOp('FRAME35')
        try:
            excel.part_parameter('FRAME35', i)
            print('FRAME35 Parameter change success')
        except:
            print('FRAME35 Parameter change error')
        try:
            if i == 0:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i == 1:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i == 2:
                mprog.activatefeature('SN1_2545_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i == 3:
                mprog.activatefeature('SN1_60_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i == 4:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i == 5:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i == 6:
                mprog.activatefeature('SN1_80200_Body', 0)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i == 7:
                mprog.activatefeature('SN1_80200_Body', 6)
                mprog.activatefeature('Hole', 2)
                mprog.partbodyfeatureactivate('SN1_25250_L')
            elif i == 8:
                mprog.activatefeature('SN1_250_Body', 0)
                mprog.activatefeature('Hole', 3)


        except:
            print('FRAME35 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME35 Update success')
            except:
                print('FRAME35 Update error')
    elif name == 'FRAME36':  # 已更改
        excel = epc.ExcelOp('FRAME36')
        try:
            excel.part_parameter('FRAME36', i)
            print('FRAME36 Parameter change success')
        except:
            print('FRAME36 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('FRAME_SN1_25_Body')
            elif i == 1:
                mprog.partbodyfeatureactivate('FRAME_SN1_3560_Body')
            elif i == 2:
                mprog.partbodyfeatureactivate('FRAME_SN1_3560_Body')
            elif i == 3:
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
        excel = epc.ExcelOp('FRAME37')
        try:
            excel.part_parameter('FRAME37', i)
            print('FRAME37 Parameter change success')
        except:
            print('FRAME37 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('Body')
                mprog.partbodyfeatureactivate('FRAME_Body')
            elif i == 1:
                mprog.partbodyfeatureactivate('Body')
                mprog.partbodyfeatureactivate('FRAME_Body')
            elif i == 2:
                mprog.partbodyfeatureactivate('Body')
                mprog.partbodyfeatureactivate('FRAME_Body')
            elif i == 3:
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
        excel = epc.ExcelOp('FRAME38')
        try:
            excel.part_parameter('FRAME38', i)
            print('FRAME38 Parameter change success')
        except:
            print('FRAME38 Parameter change error')
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i == 5:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('ABD')
                mprog.activatefeature('AIR_TANK_FIX_HOLE', 0)
            elif i == 8:
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
        excel = epc.ExcelOp('FRAME41')
        try:
            excel.part_parameter('FRAME41', i)
            FRAME41_parameter_name, FRAME41_parameter_value = excel.part_parameter('FRAME41', i)
            print('FRAME41 Parameter change success')
        except:
            print('FRAME41 Parameter change error')

        try:
            if i == 1:  # 35
                mprog.partbodyfeatureactivate('Chamfer.1')
                mprog.partbodyfeatureactivate('Chamfer.2')
                mprog.partbodyfeatureactivate('Chamfer.3')
        except:
            print('FRAME41 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME41 Update success')
            except:
                print('FRAME41 Update error')
    elif name == 'FRAME42':
        excel = epc.ExcelOp('FRAME42')
        try:
            FRAME42_parameter_name, FRAME42_parameter_value = excel.part_parameter(
                'FRAME42', i)
            print('FRAME42 Parameter change success')
        except BaseException:
            print('FRAME42 Parameter change error')
        try:
            mprog.Update()
            print('FRAME42 Update success')
        except BaseException:
            print('FRAME42 Update error')
    elif name == 'FRAME43':
        excel = epc.ExcelOp('FRAME43')
        try:
            FRAME43_parameter_name, FRAME43_parameter_value = excel.part_parameter(
                'FRAME43', i)
            print('FRAME43 Parameter change success')
        except BaseException:
            print('FRAME43 Parameter change error')
        try:
            if i == 4:
                mprog.partbodyfeatureactivate('machining_Pocket')
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
        excel = epc.ExcelOp('FRAME45')
        try:
            excel.part_parameter('FRAME45', i)
            FRAME45_parameter_name, FRAME45_parameter_value = excel.part_parameter('FRAME45', i)
            print('FRAME45 Parameter change success')
        except:
            print('FRAME45 Parameter change error')
        try:
            if i == 3:  # 60N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
            elif i == 4:  # 80N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
            elif i == 5:  # 110N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
            elif i == 6:  # 160N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
            elif i == 7:  # 200N
                mprog.partbodyfeatureactivate('Chamfer.G')
                mprog.partbodyfeatureactivate('Chamfer.F')
            elif i == 8:  # 250N
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
        excel = epc.ExcelOp('FRAME47')
        try:
            excel.part_parameter('FRAME47', i)
            FRAME47_parameter_name, FRAME47_parameter_value = excel.part_parameter('FRAME47', i)
            print('FRAME47 Parameter change success')
        except:
            print('FRAME47 Parameter change error')

        try:
            if i == 4:  # 80
                mprog.partbodyfeatureactivate('上導角')
            elif i == 5:  # 110
                mprog.partbodyfeatureactivate('上導角')
                mprog.partbodyfeatureactivate('下導角')
            elif i == 6:  # 160
                mprog.partbodyfeatureactivate('上導角')
                mprog.partbodyfeatureactivate('下導角')
            elif i == 7:  # 200
                mprog.partbodyfeatureactivate('上導角')
                mprog.partbodyfeatureactivate('下導角')
            elif i == 8:  # 250
                mprog.partbodyfeatureactivate('250上導角')
        except:
            print('FRAME47 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME47 Update success')
            except:
                print('FRAME47 Update error')
    elif name == 'FRAME48':
        excel = epc.ExcelOp('FRAME48')
        try:
            excel.part_parameter('FRAME48', i)
            FRAME48_parameter_name, FRAME48_parameter_value = excel.part_parameter('FRAME48', i)
            print('FRAME48 Parameter change success')
        except:
            print('FRAME48 Parameter change error')

        try:
            if i == 4:  # 80
                pass
            elif i == 5:  # 110
                pass
            elif i == 6:  # 160
                pass
            elif i == 7:  # 200
                pass
            elif i == 8:  # 250
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
        excel = epc.ExcelOp('FRAME49')
        try:
            excel.part_parameter('FRAME49', i)
            print('FRAME49 Parameter change success')
        except:
            print('FRAME49 Parameter change error')
        try:
            if i == 4:
                mprog.activatefeature('body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                mprog.activatefeature('M12通', 0)
            elif i == 5:
                mprog.activatefeature('body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                mprog.activatefeature('M12通', 0)
            elif i == 6:
                mprog.activatefeature('body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                mprog.activatefeature('M12通', 0)
            elif i == 7:
                mprog.activatefeature('body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_80250_c')
                mprog.activatefeature('M12通', 0)
            elif i == 8:
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
        excel = epc.ExcelOp('FRAME50')
        try:
            excel.part_parameter('FRAME50', i)
            print('FRAME50 Parameter change success')
        except:
            print('FRAME50 Parameter change error')
        try:
            if i == 4:
                mprog.activatefeature('Body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_c')
                mprog.activatefeature('Hole_1', 0)
            elif i == 5:
                mprog.activatefeature('Body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_c')
                mprog.activatefeature('Hole_1', 0)
            elif i == 6:
                mprog.activatefeature('Body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_c')
                mprog.activatefeature('Hole_1', 0)
            elif i == 7:
                mprog.activatefeature('Body', 0)
                mprog.partbodyfeatureactivate('FRAME_SN1_c')
                mprog.activatefeature('Hole_1', 0)
            elif i == 8:
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
    elif name == 'FRAME52':
        excel = epc.ExcelOp('FRAME52')
        try:
            excel.part_parameter('FRAME52', i)
            FRAME52_parameter_name, FRAME52_parameter_value = excel.part_parameter('FRAME52', i)
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
        excel = epc.ExcelOp('FRAME53')
        try:
            FRAME53_parameter_name, FRAME53_parameter_value = excel.part_parameter(
                'FRAME53', i)
            print('FRAME53 Parameter change success')
        except BaseException:
            print('FRAME53 Parameter change error')
        try:
            mprog.Update()
            print('FRAME53 Update success')
        except BaseException:
            print('FRAME53 Update error')
    elif name == 'FRAME54':  # 已更改
        excel = epc.ExcelOp('FRAME54')
        try:
            excel.part_parameter('FRAME54', i)
            print('FRAME54 Parameter change success')
        except:
            print('FRAME54 Parameter change error')
        try:
            if i == 6:
                mprog.partbodyfeatureactivate('Body')
            elif i == 7:
                mprog.partbodyfeatureactivate('Body')
            elif i == 8:
                mprog.partbodyfeatureactivate('Body')
        except:
            print('FRAME54 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME54 Update success')
            except:
                print('FRAME54 Update error')
