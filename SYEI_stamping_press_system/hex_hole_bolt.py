import main_program as mprog
import excel_parameter_change as epc

def hex_hole_bolt_parameter_design_part(hex_hole_bolt_type):
    excel = epc.ExcelOp('螺栓尺寸表', '六角孔螺栓')
    if hex_hole_bolt_type == 'M3':
        hex_hole_bolt_number = 0
    if hex_hole_bolt_type == 'M4':
        hex_hole_bolt_number = 1
    if hex_hole_bolt_type == 'M5':
        hex_hole_bolt_number = 2
    if hex_hole_bolt_type == 'M6':
        hex_hole_bolt_number = 3
    if hex_hole_bolt_type == 'M8':
        hex_hole_bolt_number = 4
    if hex_hole_bolt_type == 'M10':
        hex_hole_bolt_number = 5
    if hex_hole_bolt_type == 'M12':
        hex_hole_bolt_number = 6
    if hex_hole_bolt_type == 'M14':
        hex_hole_bolt_number = 7
    if hex_hole_bolt_type == 'M16':
        hex_hole_bolt_number = 8
    if hex_hole_bolt_type == 'M20':
        hex_hole_bolt_number = 9
    if hex_hole_bolt_type == 'M22':
        hex_hole_bolt_number = 10
    if hex_hole_bolt_type == 'M24':
        hex_hole_bolt_number = 11
    if hex_hole_bolt_type == 'M30':
        hex_hole_bolt_number = 12
    if hex_hole_bolt_type == 'M33':
        hex_hole_bolt_number = 13
    if hex_hole_bolt_type == 'M36':
        hex_hole_bolt_number = 14

    try:
        parameter_name, paramter_value = excel.get_sheet_par('hex_hole_bolt', hex_hole_bolt_number)
        print('hex_hole_bolt parameter change success')
    except:
        print('hex_hole_bolt parameter change error')
    finally:
        try:
            mprog.update()
            print('hex_hole_bolt update success')
        except:
            print('hex_hole_bolt update error')

hex_hole_bolt_parameter_design_part('M36')