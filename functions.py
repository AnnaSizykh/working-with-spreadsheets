import math
import openpyxl
import re


def addition(sheet, left_corner, right_corner, result_cell):
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, left_corner) is None or re.match(cell_pattern, right_corner) is \
            None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    summa = 0
    for cell_column in sheet[left_corner: right_corner]:
        for cell in cell_column:
            addend = cell.value
            if isinstance(addend, int) or isinstance(addend, float):
                summa += addend
    sheet[result_cell] = summa
    return summa


def subtraction(sheet, reduced_cell, subtrahend_cell, result_cell):
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, reduced_cell) is None or re.match(cell_pattern, subtrahend_cell) is \
            None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
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
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, left_corner) is None or re.match(cell_pattern, right_corner) is \
            None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    composition = 1
    for cell_column in sheet[left_corner: right_corner]:
        for cell in cell_column:
            factor = cell.value
            if isinstance(factor, int) or isinstance(factor, float):
                composition *= factor
    sheet[result_cell] = composition
    return composition


def division(sheet, divisible_cell, divisor_cell, result_cell):
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, divisor_cell) is None or re.match(cell_pattern, divisible_cell) is \
            None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
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
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, number_cell) is None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    number = sheet[number_cell].value
    if (isinstance(number, int) or isinstance(number, float)) and isinstance(decimal, int):
        rounded = round(number, decimal)
        sheet[result_cell] = rounded
        return rounded
    return number


def exponentiation(sheet, number_cell, extent, result_cell):
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, number_cell) is None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    number = sheet[number_cell].value
    if not (isinstance(number, int) and isinstance(number, float) and
            isinstance(extent, int) and isinstance(extent, float)):
        sheet[result_cell] = 'N/A'
        return 'N/A'
    result = number ** extent
    sheet[result_cell] = result
    return result
    pass


def logarithm(sheet, number_cell, base, result_cell):
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, number_cell) is None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    number = sheet[number_cell].value
    if (not (isinstance(number, int) and isinstance(number, float) and isinstance(result_cell, int)
             and isinstance(result_cell, float))) or number < 0 or base < 0 or base == 1:
        sheet[result_cell] = 'N/A'
        return 'N/A'
    result = math.log(number, base)
    sheet[result_cell] = result
    return result


def mean(sheet, left_corner, right_corner, result_cell):
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, left_corner) is None or re.match(cell_pattern, right_corner) is \
            None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    counter = 0
    summa = 0
    for cell_column in sheet[left_corner: right_corner]:
        for cell in cell_column:
            addend = cell.value
            if isinstance(addend, int) or isinstance(addend, float):
                summa += addend
                counter += 1
    if counter == 0:
        sheet[result_cell] = 'N/A'
        return 'N/A'
    result = summa / counter
    sheet[result_cell] = result
    return result


def move(sheet, moved_left_corner, moved_right_corner, added_left_corner, added_right_corner):
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, moved_left_corner) is None or re.match(cell_pattern, moved_right_corner) is \
            None or re.match(cell_pattern, added_left_corner) is None \
            or re.match(cell_pattern, added_right_corner) is None:
        return 'Coordinate error'
    paste_cut = sheet[added_left_corner : added_right_corner]
    for ind_x, cell_column in enumerate(sheet[moved_left_corner : moved_right_corner]):
        for ind_y, cell in enumerate(cell_column):
            moved_value = cell.value
            sheet.cell(row=ind_y, column=ind_x,  value=moved_value)
            sheet[cell_moved] = ''


def copy(sheet, copied_left_corner, copied_right_corner, added_left_corner, added_right_corner):
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, copied_left_corner) is None or re.match(cell_pattern, copied_right_corner) is \
            None or re.match(cell_pattern, added_left_corner) is None \
            or re.match(cell_pattern, added_right_corner) is None:
        return 'Coordinate error'
    paste_cut = sheet[added_left_corner: added_right_corner]
    for ind_x, cell_column in enumerate(sheet[copied_left_corner: copied_right_corner]):
        for ind_y, cell in enumerate(cell_column):
            copied_value = cell.value
            sheet.cell(row=ind_y, column=ind_x,  value=copied_value)


def delete(sheet, left_corner, right_corner):
    cell_pattern = re.compile('^[A-Z]+\d+$')
    if re.match(cell_pattern, left_corner) is None or re.match(cell_pattern, right_corner) is None:
        return 'Coordinate error'
    for cell_column in sheet[left_corner : right_corner]:
        for cell in cell_column:
            cell.value = ''


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
            if element == str(cell.value) or element in str(cell.value):
                similar_results.append(cell.coordinate)
    if len(similar_results) != 0:
        return similar_results
