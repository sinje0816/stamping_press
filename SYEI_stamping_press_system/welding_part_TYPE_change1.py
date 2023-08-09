import main_program as mprog
import excel_parameter_change as epc


# 母檔變數變換
def change_welding_feature(name, i):
    # 零件特徵開啟
    if name == 'FRAME1':
        if i == 0:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('AA2')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('VV2')
            mprog.partbodyfeatureactivate('VV4')
            mprog.partbodyfeatureactivate('R倒角')
            mprog.partbodyfeatureactivate('R倒角填料_1')
            mprog.partbodyfeatureactivate('R倒角填料_2')
            mprog.partbodyfeatureactivate('Z倒角')
            mprog.partbodyfeatureactivate('小噸數E倒角_B')
            mprog.partbodyfeatureactivate('小噸數Y倒角_B1')
            mprog.partbodyfeatureactivate('喉口倒角_B')
            mprog.partbodyfeatureactivate('喉口倒角補料_B')
            mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
            mprog.activatefeature('110通孔', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('光電裝置電線固定孔', 5)
            mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 2)
            mprog.activatefeature('兩點組合', 0)
            mprog.partdeactivate('machining_DD')
            mprog.partdeactivate('machining_E上方研磨')
            mprog.bodydeactivate('machining_2x8_12通孔', 0)
            mprog.bodydeactivate('machining_aaaaa2', 0)
            mprog.bodydeactivate('machining_aaaaa4', 0)
        elif i == 1:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('VV2')
            mprog.partbodyfeatureactivate('VV4')
            mprog.partbodyfeatureactivate('R倒角')
            mprog.partbodyfeatureactivate('R倒角填料_1')
            mprog.partbodyfeatureactivate('R倒角填料_2')
            mprog.partbodyfeatureactivate('Z倒角')
            mprog.partbodyfeatureactivate('小噸數E倒角_B')
            mprog.partbodyfeatureactivate('小噸數Y倒角_B1')
            mprog.partbodyfeatureactivate('喉口倒角_B')
            mprog.partbodyfeatureactivate('喉口倒角補料_B')
            mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
            mprog.activatefeature('110通孔', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('光電裝置電線固定孔', 5)
            mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 3)
            mprog.activatefeature('兩點組合', 0)
            mprog.partdeactivate('machining_DD')
            mprog.partdeactivate('machining_E上方研磨')
            mprog.bodydeactivate('machining_2x8_12通孔', 0)
            mprog.bodydeactivate('machining_aaaaa2', 0)
            mprog.bodydeactivate('machining_aaaaa4', 0)
        elif i == 2:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('VV2')
            mprog.partbodyfeatureactivate('VV4')
            mprog.partbodyfeatureactivate('R倒角')
            mprog.partbodyfeatureactivate('R倒角填料_1')
            mprog.partbodyfeatureactivate('R倒角填料_2')
            mprog.partbodyfeatureactivate('Z倒角')
            mprog.partbodyfeatureactivate('小噸數E倒角_B')
            mprog.partbodyfeatureactivate('小噸數Y倒角_B1')
            mprog.partbodyfeatureactivate('喉口倒角_B')
            mprog.partbodyfeatureactivate('喉口倒角補料_B')
            mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
            mprog.activatefeature('110通孔', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('光電裝置電線固定孔', 5)
            mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 2)
            mprog.activatefeature('兩點組合', 0)
            mprog.partdeactivate('machining_DD')
            mprog.partdeactivate('machining_E上方研磨')
            mprog.bodydeactivate('machining_2x8_12通孔', 0)
            mprog.bodydeactivate('machining_aaaaa2', 0)
            mprog.bodydeactivate('machining_aaaaa4', 0)
        elif i == 3:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('V1')
            mprog.partbodyfeatureactivate('V3')
            mprog.partbodyfeatureactivate('V5')
            mprog.partbodyfeatureactivate('R倒角')
            mprog.partbodyfeatureactivate('R倒角填料_1')
            mprog.partbodyfeatureactivate('R倒角填料_2')
            mprog.partbodyfeatureactivate('Z倒角')
            mprog.partbodyfeatureactivate('喉口倒角_60')
            mprog.partbodyfeatureactivate('喉口倒圓角_60')
            mprog.partbodyfeatureactivate('喉口倒角補料_60')
            mprog.partbodyfeatureactivate('E_倒角')
            mprog.activatefeature('110通孔', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('光電裝置電線固定孔', 5)
            mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 2)
            mprog.activatefeature('兩點組合', 0)
            mprog.partdeactivate('machining_DD')
            mprog.partdeactivate('machining_E上方研磨')
            mprog.bodydeactivate('machining_2x8_12通孔', 0)
            mprog.bodydeactivate('machining_aaaaa2', 0)
            mprog.bodydeactivate('machining_aaaaa4', 0)
        elif i == 4:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('V1')
            mprog.partbodyfeatureactivate('V3')
            mprog.partbodyfeatureactivate('V5')
            mprog.partbodyfeatureactivate('Z倒角')
            mprog.partbodyfeatureactivate('喉口倒角_B')
            mprog.partbodyfeatureactivate('喉口倒角_B_圓角')
            mprog.partbodyfeatureactivate('B倒角_E段')
            mprog.partbodyfeatureactivate('喉口倒角補料_80')
            mprog.activatefeature('1-M5X10L(配管用)', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('光電裝置電線固定孔', 5)
            mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 3)
            mprog.activatefeature('兩點組合', 0)
            mprog.activatefeature('3-M5X10L(配管用)(r1)', 1)
            mprog.partdeactivate('machining_DD')
            mprog.partdeactivate('machining_E上方研磨')
            mprog.bodydeactivate('machining_2x8_12通孔', 0)
            mprog.bodydeactivate('machining_aaaaa2', 0)
            mprog.bodydeactivate('machining_aaaaa4', 0)

        elif i == 5:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('V1')
            mprog.partbodyfeatureactivate('V3')
            mprog.partbodyfeatureactivate('V5')
            mprog.partbodyfeatureactivate('Z倒角')
            mprog.partbodyfeatureactivate('喉口倒角_非B')
            mprog.partbodyfeatureactivate('喉口倒圓角_非B')
            mprog.partbodyfeatureactivate('喉口倒角補料_110')
            mprog.partbodyfeatureactivate('E_倒角')
            mprog.activatefeature('1-M5X10L(配管用)', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('光電裝置電線固定孔', 0)
            mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 3)
            mprog.activatefeature('兩點組合', 0)
            mprog.activatefeature('3-M5X10L(配管用)(r1)', 2)
            mprog.partdeactivate('machining_DD')
            mprog.partdeactivate('machining_E上方研磨')
            mprog.bodydeactivate('machining_2x8_12通孔', 0)
            mprog.bodydeactivate('machining_aaaaa2', 0)
            mprog.bodydeactivate('machining_aaaaa4', 0)

        elif i == 6:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('V1')
            mprog.partbodyfeatureactivate('V3')
            mprog.partbodyfeatureactivate('V5')
            mprog.partbodyfeatureactivate('Z倒角')
            mprog.partbodyfeatureactivate('喉口倒角_非B')
            mprog.partbodyfeatureactivate('喉口倒圓角_非B')
            mprog.partbodyfeatureactivate('喉口倒角補料_160')
            mprog.partbodyfeatureactivate('E_倒角')
            mprog.activatefeature('1-M5X10L(配管用)', 0)
            mprog.activatefeature('3-M5X10L(配管用)(r1)', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('光電裝置電線固定孔', 0)
            mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 5)
            mprog.activatefeature('兩點組合', 0)
            mprog.partdeactivate('machining_DD')
            mprog.partdeactivate('machining_E上方研磨')
            mprog.bodydeactivate('machining_2x8_12通孔', 0)
            mprog.bodydeactivate('machining_aaaaa2', 0)
            mprog.bodydeactivate('machining_aaaaa4', 0)
        elif i == 7:
            mprog.partbodyfeatureactivate('before250_喉口')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('V1')
            mprog.partbodyfeatureactivate('V3')
            mprog.partbodyfeatureactivate('V5')
            mprog.partbodyfeatureactivate('Z倒角')
            mprog.partbodyfeatureactivate('喉口倒角_非B')
            mprog.partbodyfeatureactivate('喉口倒圓角_非B')
            mprog.partbodyfeatureactivate('喉口倒角補料_200')
            mprog.partbodyfeatureactivate('E_倒角')
            mprog.activatefeature('1-M5X10L(配管用)', 0)
            mprog.activatefeature('3-M5X10L(配管用)(r1)', 2)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('光電裝置電線固定孔', 0)
            mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 5)
            mprog.activatefeature('兩點組合', 0)
            mprog.partdeactivate('machining_DD')
            mprog.partdeactivate('machining_E上方研磨')
            mprog.bodydeactivate('machining_2x8_12通孔', 0)
            mprog.bodydeactivate('machining_aaaaa2', 0)
            mprog.bodydeactivate('machining_aaaaa4', 0)
        elif i == 8:
            mprog.partbodyfeatureactivate('T250')
            mprog.partbodyfeatureactivate('Z1')
            mprog.partbodyfeatureactivate('V1_250')
            mprog.partbodyfeatureactivate('V3_250')
            mprog.partbodyfeatureactivate('V5_250')
            mprog.partbodyfeatureactivate('Z倒角')
            mprog.partbodyfeatureactivate('喉口倒角_250')
            mprog.partbodyfeatureactivate('喉口倒圓角_250')
            mprog.partbodyfeatureactivate('喉口倒角補料_250')
            mprog.partbodyfeatureactivate('E_倒角')
            mprog.activatefeature('1-M5X10L(配管用)', 0)
            mprog.activatefeature('3-M5X10L(配管用)(r1)', 0)
            mprog.activatefeature('油面劑用', 0)
            mprog.activatefeature('60通孔', 0)
            mprog.activatefeature('電動黃油泵用', 0)
            mprog.activatefeature('吊孔', 0)
            mprog.activatefeature('光電裝置電線固定孔', 0)
            mprog.activatefeature('配管用(電動黃油泵→潤滑分配塊用)', 0)
            mprog.activatefeature('兩點組合', 0)
            mprog.partdeactivate('machining_DD')
            mprog.partdeactivate('machining_E上方研磨')
            mprog.bodydeactivate('machining_2x8_12通孔', 0)
            mprog.bodydeactivate('machining_aaaaa2', 0)
            mprog.bodydeactivate('machining_aaaaa4', 0)
        try:
            mprog.update()
            print('FRAME1 update success')
        except BaseException:
            print('FRAME1 update error')
    elif name == 'FRAME2':
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒圓角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_45')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角_qqqq')
                mprog.partbodyfeatureactivate('R倒角_G1')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段_B1倒角')
                mprog.partbodyfeatureactivate('E段_B倒角_Y')
                mprog.partbodyfeatureactivate('pad1左&右下倒角')
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
            elif i == 1:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒圓角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_60')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角_qqqq')
                mprog.partbodyfeatureactivate('R倒角_G1')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段_B1倒角')
                mprog.partbodyfeatureactivate('E段_B倒角_Y')
                mprog.partbodyfeatureactivate('pad1左&右下倒角')
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
            elif i == 2:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('VV2')
                mprog.partbodyfeatureactivate('VV4')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒圓角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_45')
                mprog.partbodyfeatureactivate('R倒角')
                mprog.partbodyfeatureactivate('R倒角_qqqq')
                mprog.partbodyfeatureactivate('R倒角_G1')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段_B1倒角')
                mprog.partbodyfeatureactivate('E段_B倒角_Y')
                mprog.partbodyfeatureactivate('pad1左&右下倒角')
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
            elif i == 3:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('AA2')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_60')
                mprog.partbodyfeatureactivate('R倒角_無qqqq')
                mprog.partbodyfeatureactivate('R倒角_G1')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_非B')
                mprog.partbodyfeatureactivate('pad1左&右下倒角')
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
            elif i == 4:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('喉口倒角_B')
                mprog.partbodyfeatureactivate('喉口倒圓角_B')
                mprog.partbodyfeatureactivate('喉口倒角補料_80')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_B')
                mprog.partbodyfeatureactivate('jjj4')
                mprog.partbodyfeatureactivate('mmm4')
                mprog.partbodyfeatureactivate('pad1左&右下倒角')
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
            elif i == 5:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_110')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_非B')
                mprog.partbodyfeatureactivate('dd4')
                mprog.partbodyfeatureactivate('x4')
                mprog.partbodyfeatureactivate('pad1左&右下倒角')
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
            elif i == 6:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_160')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_非B')
                mprog.partbodyfeatureactivate('ggg4')
                mprog.partbodyfeatureactivate('xx3')
                mprog.partbodyfeatureactivate('iii4')
                mprog.partbodyfeatureactivate('fff4')
                mprog.partbodyfeatureactivate('pad1左&右下倒角')
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
            elif i == 7:
                mprog.partbodyfeatureactivate('before250_喉口')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1')
                mprog.partbodyfeatureactivate('V3')
                mprog.partbodyfeatureactivate('V5')
                mprog.partbodyfeatureactivate('喉口倒角_非B')
                mprog.partbodyfeatureactivate('喉口倒圓角_非B')
                mprog.partbodyfeatureactivate('喉口倒角補料_200')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_非B')
                mprog.partbodyfeatureactivate('iii4')
                mprog.partbodyfeatureactivate('ssss4')
                mprog.partbodyfeatureactivate('sss4')
                mprog.partbodyfeatureactivate('pad1左&右下倒角')
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
            elif i == 8:
                mprog.partbodyfeatureactivate('T250')
                mprog.partbodyfeatureactivate('Z1')
                mprog.partbodyfeatureactivate('V1_250')
                mprog.partbodyfeatureactivate('V3_250')
                mprog.partbodyfeatureactivate('V5_250')
                mprog.partbodyfeatureactivate('喉口倒角_250')
                mprog.partbodyfeatureactivate('喉口倒圓角_250')
                mprog.partbodyfeatureactivate('喉口倒角補料_250')
                mprog.partbodyfeatureactivate('Z倒角')
                mprog.partbodyfeatureactivate('E段倒角_非B')
                mprog.partbodyfeatureactivate('pad1左&右下倒角')
                mprog.partdeactivate('machining_DD')
                mprog.partdeactivate('machining_E上方研磨')
                mprog.bodydeactivate('machining_2x8_12通孔', 0)
                mprog.bodydeactivate('machining_aaaaa2', 0)
                mprog.bodydeactivate('machining_aaaaa4', 0)
        except BaseException:
            print('FRAME2 Part activate error')
        finally:
            mprog.update()
            print('FRAME2 update success')
    elif name == 'FRAME5':
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
    elif name == 'FRAME18':   # Add el
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
    elif name == 'FRAME21':
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
    elif name == 'FRAME23':
        try:
            if i == 0:
                mprog.partbodyfeatureactivate('25至60除料')
                mprog.partbodyfeatureactivate('25至60導圓角')
            elif i == 1:
                mprog.partbodyfeatureactivate('25至60除料')
                mprog.partbodyfeatureactivate('25至60導圓角')
            elif i == 2:
                mprog.partbodyfeatureactivate('25至60除料')
                mprog.partbodyfeatureactivate('25至60導圓角')
            elif i == 3:
                mprog.partbodyfeatureactivate('25至60除料')
                mprog.partbodyfeatureactivate('25至60導圓角')
            elif i == 4:
                mprog.partbodyfeatureactivate('80至250除料')
                mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
            elif i == 5:
                mprog.partbodyfeatureactivate('80至250除料')
                mprog.partbodyfeatureactivate('80至250導圓角(除160和200)')
            elif i == 6:
                mprog.partbodyfeatureactivate('80至250除料')
            elif i == 7:
                mprog.partbodyfeatureactivate('80至250除料')
            elif i == 8:
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
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 1:
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 2:
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 3:
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 4:
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 5:
                mprog.partbodyfeatureactivate('F')
            elif i == 6:
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 7:
                mprog.partbodyfeatureactivate('F')
                mprog.partbodyfeatureactivate('G')
            elif i == 8:
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
        try:
            if i == 0:
                mprog.partdeactivate('machining_aaaaa2')
            elif i == 1:
                mprog.partdeactivate('machining_aaaaa2')
            elif i == 2:
                mprog.partdeactivate('machining_aaaaa2')
                mprog.bodydeactivate('machining_ddddd', 0)
            elif i == 3:
                mprog.partdeactivate('machining_aaaaa2')
                mprog.bodydeactivate('machining_ddddd', 0)
        except BaseException:
            print('FRAME27 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME27 Update success')
            except BaseException:
                print('FRAME27 Update error')
    elif name == 'FRAME27_1':
        try:
            if i == 0:
                mprog.partdeactivate('machining_aaaaa2')
            elif i == 1:
                mprog.partdeactivate('machining_aaaaa2')
            elif i == 2:
                mprog.partdeactivate('machining_aaaaa2')
                mprog.bodydeactivate('machining_ddddd', 0)
            elif i == 3:
                mprog.partdeactivate('machining_aaaaa2')
                mprog.bodydeactivate('machining_ddddd', 0)
        except BaseException:
            print('FRAME27_1 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME27_1 Update success')
            except BaseException:
                print('FRAME27_1 Update error')
    elif name == 'FRAME29':
        try:
            mprog.partdeactivate('Hole.1')
        finally:
            try:
                mprog.Update()
                print('FRAME29 Update success')
            except BaseException:
                print('FRAME29 Update error')
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
        pass
    elif name == 'FRAME43':
        try:
            if i == 3:
                mprog.partbodyfeatureactivate('Chamfer.1')
            elif i == 4:
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
    elif name == 'FRAME47':
        excel = epc.ExcelOp('FRAME47')
        try:
            excel.part_parameter('FRAME47', i)
            FRAME47_parameter_name, FRAME47_parameter_value = excel.part_parameter('FRAME47', i)
            print('FRAME47 Parameter change success')
        except:
            print('FRAME47 Parameter change error')

        try:
            if i == 4:    # 80
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
            if i == 4:    # 80
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
    elif name == 'FRAME52':
        excel = epc.ExcelOp('FRAME52')
        try:
            excel.part_parameter('FRAME52', i)
            FRAME52_parameter_name, FRAME52_parameter_value = excel.part_parameter('FRAME52', i)
            print('FRAME52 Parameter change success')
        except:
            print('FRAME52 Parameter change error')
        try:
            if i == 5:
                mprog.partdeactivate('machining_除料')
                mprog.partdeactivate('machining_通孔')
                mprog.partbodyfeatureactivate('top')
                mprog.partbodyfeatureactivate('bottom')
            elif i == 6:
                mprog.partdeactivate('machining_除料')
                mprog.partdeactivate('machining_通孔')
                mprog.partbodyfeatureactivate('top')
                mprog.partbodyfeatureactivate('bottom')
            elif i == 7:
                mprog.partdeactivate('machining_除料')
                mprog.partdeactivate('machining_通孔')
                mprog.partbodyfeatureactivate('top')
                mprog.partbodyfeatureactivate('bottom')
            elif i == 8:
                mprog.partdeactivate('machining_除料')
                mprog.partdeactivate('machining_通孔')
                mprog.partbodyfeatureactivate('top')
                mprog.partbodyfeatureactivate('bottom')
        except:
            print('FRAME52 Parameter activate error')
        finally:
            try:
                mprog.Update()
                print('FRAME52 Update success')
            except:
                print('FRAME52 Update error')
    elif name == 'FRAME53':
        try:
            mprog.Update()
            print('FRAME53 Update success')
        except BaseException:
            print('FRAME53 Update error')


change_welding_feature('FRAME52', 2)