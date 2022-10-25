import math
import openpyxl


def addition(sheet, left_corner, right_corner, result_cell):
    summa = 0
    for cell_object in sheet[left_corner: right_corner]:
        for cell in cell_object:
            addend = cell.value
            if isinstance(addend, int) or isinstance(addend, float):
                summa += addend
    sheet[result_cell] = summa
    return summa


def subtraction(sheet, reduced_cell, subtrahend_cell, result_cell):
    reduced = sheet[reduced_cell].value
    subtrahend = sheet[subtrahend_cell].value
    if not (isinstance(reduced, int) or isinstance(reduced, float) or
            isinstance(subtrahend, int) or isinstance(subtrahend, float)):
        sheet[result_cell] = 'N/A'
        return 'N/A'
    difference = reduced - subtrahend
    sheet[result_cell] = difference
    return difference


def multiplication(sheet, left_corner, right_corner, result_cell):
    composition = 1
    for cell_object in sheet[left_corner: right_corner]:
        for cell in cell_object:
            factor = cell.value
            if isinstance(factor, int) or isinstance(factor, float):
                composition *= factor
    sheet[result_cell] = composition
    return composition


def division(sheet, divisible_cell, divisor_cell, result_cell):
    divisible = sheet[divisible_cell].value
    divisor = sheet[divisor_cell].value
    if not (isinstance(divisible, int) or isinstance(divisible, float) or
            isinstance(divisor, int) or isinstance(divisor, float)) or divisor == 0:
        sheet[result_cell] = 'N/A'
        return 'N/A'
    quotient = divisible / divisor
    sheet[result_cell] = quotient
    return quotient


def round_number(sheet, number_cell, decimal, result_cell):
    number = sheet[number_cell].value
    if (isinstance(number, int) or isinstance(number, float)) and isinstance(decimal, int):
        rounded = round(number, decimal)
        sheet[result_cell] = rounded
        return rounded
    return number


def exponentiation(sheet, number_cell, extent, result_cell):
    number = sheet[number_cell].value
    if not isinstance(number, int) or not isinstance(number, float) \
            or not isinstance(extent, int) or not isinstance(extent, float):
        sheet[result_cell] = 'N/A'
        return 'N/A'
    result = number ** extent
    sheet[result_cell] = result
    return result
    pass


def logarithm(sheet, number_cell, base, result_cell):
    number = sheet[number_cell].value
    if not isinstance(number, int) or not isinstance(number, float) or not isinstance(result_cell, int) or \
            not isinstance(result_cell, float) or number < 0 or base < 0 or base == 1:
        sheet[result_cell] = 'N/A'
        return 'N/A'
    result = math.log(number, base)
    sheet[result_cell] = result
    return result


def mean(sheet, left_corner, right_corner, result_cell):
    counter = 0
    summa = 0
    for cell_object in sheet[left_corner: right_corner]:
        for cell in cell_object:
            addend = cell.value
            if isinstance(addend, int) or isinstance(addend, float):
                summa += addend
                counter += 1
    if counter == 0:
        sheet[result_cell] = 'N/A'
        return 'N/A'
    result = summa/counter
    sheet[result_cell] = result
    return result


def use_math_module():
    pass


def move(sheet, cell_moved, cell_added):
    moved_value = sheet[cell_moved].value
    sheet[cell_added] = moved_value
    sheet[cell_moved] = ''


def copy(sheet, cell_copied, cell_pasted):
    copied_value = sheet[cell_copied].value
    sheet[cell_pasted] = copied_value


def delete(sheet, cell_deleted):
    sheet[cell_deleted] = ''


def compare(sheet, cell_1, cell_2, result_cell):
    value_1 = sheet[cell_1].value
    value_2 = sheet[cell_2].value
    if value_1 == value_2:
        sheet[result_cell] = True
        return True
    else:
        sheet[result_cell] = False
        return False


def find(sheet, element):
    similar_results = []
    for cell_object in sheet:
        for cell in cell_object:
            if element == cell.value or element in cell.value:
                similar_results.append(cell.coordinate)
    if len(similar_results) != 0:
        return similar_results


def sort(sheet, left_corner, right_corner):
    pass


def repeat_for_several_cells(sheet, function_name, first_start_cell, first_finish_cell, second_start_cell=None,
                             result_start_cell=None, additional_var=None):
    if sheet[first_start_cell].column != sheet[first_finish_cell].column:
        return 'N/A'
    result_row = None
    result_column = None
    second_row = 0
    second_column = 0
    if second_start_cell:
        second_row = int(sheet[second_start_cell].row)
        second_column = sheet[second_start_cell].column
    for cell_object in sheet[first_start_cell: first_finish_cell]:
        if function_name == 'delete':
            delete(sheet, cell_object.coordinate)
            continue
        second_cell_coordinate = str(second_column) + str(second_row)
        if result_start_cell:
            result_row = int(sheet[result_start_cell].row)
            result_column = sheet[result_start_cell].column
            result_cell_coordinate = str(result_column) + str(result_row)
            if function_name == 'subtraction':
                subtraction(sheet, cell_object.coordinate, second_cell_coordinate, result_cell_coordinate)
            if function_name == 'division':
                division(sheet, cell_object.coordinate, second_cell_coordinate, result_cell_coordinate)
            if function_name == 'compare':
                compare(sheet, cell_object.coordinate, second_cell_coordinate, result_cell_coordinate)
            result_row += 1
        else:
            if function_name == 'copy':
                copy(sheet, cell_object.coordinate, second_cell_coordinate)
            if function_name == 'move':
                move(sheet, cell_object.coordinate, second_cell_coordinate)
        second_row += 1
