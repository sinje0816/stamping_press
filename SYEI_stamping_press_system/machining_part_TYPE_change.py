import main_program as mprog
import excel_parameter_change as epc
import time

# 母檔變數變換


def change_parameter(name, i, machiningdiepad):
    # 已改
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


change_parameter('FRAME52', 8, 0)