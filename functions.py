import math


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
    if not isinstance(number, int) or not isinstance(number, float) or not isinstance(result_cell, int):
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


def mean():
    pass


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


def find():
    pass


def sort():
    pass


def repeat_for_several_cells(sheet, function_name, first_start_cell, first_finish_cell, second_start_cell,
                             result_start_cell=None, additional_var=None):
    pass
