import excel_parameter_change as epc
import parameter as par


def edge_test(i):
        excel_cutout_limit = epc.ExcelOp('平板', '下料孔界線')
        excel_cutout_limit.get_cutout_limit(i)
        print(par.cutout_all_limit['X'], par.cutout_all_limit['Y'])
        if par.plate_hole_type[0] == '圓孔':
            if int(par.cutout_part_dimension[0]) / 2 + 10 >= par.cutout_all_limit[
                'X'] / 2 or int(par.cutout_part_dimension[0]) / 2 + 10 >= \
                    par.cutout_all_limit['Y'] / 2:
                print('error')
        elif par.plate_hole_type[0] == '方孔':
            if int(par.cutout_part_dimension[0]) / 2 + 10 >= par.cutout_all_limit[
                'X'] / 2 or int(par.cutout_part_dimension[0]) / 2 + 10 >= \
                    par.cutout_all_limit['Y'] / 2:
                print('error')
        elif par.plate_hole_type[0] == '漏斗型':
            if int(par.cutout_part_dimension[0]) / 2 + 10 >= par.cutout_all_limit[
                'X'] / 2 or int(par.cutout_part_dimension[0]) / 2 + 10 >= \
                    par.cutout_all_limit['Y'] / 2:
                print('error')
