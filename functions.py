import math
import re
import openpyxl


def addition(sheet: Worksheet, left_corner: str, right_corner: str, result_cell: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, left_corner) is None \
            or re.match(cell_pattern, right_corner) is \
            None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    summa = 0
    for cell_column in sheet[left_corner: right_corner]:
        for cell in cell_column:
            addend = cell.value
            if isinstance(addend, (float, int)):
                summa += addend
    sheet[result_cell] = summa
    return summa


def subtraction(sheet: Worksheet, reduced_cell: str, subtrahend_cell: str, result_cell: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, reduced_cell) is None \
            or re.match(cell_pattern, subtrahend_cell) is \
            None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    reduced = sheet[reduced_cell].value
    subtrahend = sheet[subtrahend_cell].value
    if not (isinstance(reduced, (float, int)) and isinstance(subtrahend, (float, int))):
        sheet[result_cell] = 'N/A'
        return 'N/A'
    difference = reduced - subtrahend
    sheet[result_cell] = difference
    return difference


def multiplication(sheet: Worksheet, left_corner: str, right_corner: str, result_cell: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, left_corner) is None \
            or re.match(cell_pattern, right_corner) is \
            None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    composition = 1
    for cell_column in sheet[left_corner: right_corner]:
        for cell in cell_column:
            factor = cell.value
            if isinstance(factor, (float, int)):
                composition *= factor
    sheet[result_cell] = composition
    return composition


def division(sheet: Worksheet, divisible_cell: str, divisor_cell: str, result_cell: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, divisor_cell) is None \
            or re.match(cell_pattern, divisible_cell) is \
            None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    divisible = sheet[divisible_cell].value
    divisor = sheet[divisor_cell].value
    if not (isinstance(divisible, (float, int))
            and isinstance(divisor, (float, integer))) or divisor == 0:
        sheet[result_cell] = 'N/A'
        return 'N/A'
    quotient = divisible / divisor
    sheet[result_cell] = quotient
    return quotient


def round_number(sheet: Worksheet, number_cell: str, decimal: int, result_cell: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, number_cell) is None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    number = sheet[number_cell].value
    if isinstance(number, (float, int)) and isinstance(decimal, int):
        rounded = round(number, decimal)
        sheet[result_cell] = rounded
        return rounded
    return number


def exponentiation(sheet: Worksheet, number_cell: str, extent: (int, float), result_cell: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, number_cell) is None \
            or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    number = sheet[number_cell].value
    if not (isinstance(number, (float, int)) and isinstance(extent, (float, int))):
        sheet[result_cell] = 'N/A'
        return 'N/A'
    result = number ** extent
    sheet[result_cell] = result
    return result


def logarithm(sheet: Worksheet, number_cell: str, base: (float, int), result_cell: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, number_cell) is None \
            or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    number = sheet[number_cell].value
    if (not (isinstance(number, (float, int)) and isinstance(result_cell, (float, int)))) \
            or number < 0 or base < 0 or base == 1:
        sheet[result_cell] = 'N/A'
        return 'N/A'
    result = math.log(number, base)
    sheet[result_cell] = result
    return result


def mean(sheet: Worksheet, left_corner: str, right_corner: str, result_cell: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, left_corner) is None \
            or re.match(cell_pattern, right_corner) is \
            None or re.match(cell_pattern, result_cell) is None:
        return 'Coordinate error'
    counter = 0
    summa = 0
    for cell_column in sheet[left_corner: right_corner]:
        for cell in cell_column:
            addend = cell.value
            if isinstance(addend, (float, int)):
                summa += addend
                counter += 1
    if counter == 0:
        sheet[result_cell] = 'N/A'
        return 'N/A'
    result = summa / counter
    sheet[result_cell] = result
    return result


def move(sheet: Worksheet, moved_left_corner: str, moved_right_corner: str, added_left_corner: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, moved_left_corner) is None or \
            re.match(cell_pattern, moved_right_corner) is None \
            or re.match(cell_pattern, added_left_corner) is None:
        return 'Coordinate error'
    moving_column = sheet[added_left_corner].column
    moving_row = sheet[added_left_corner].row
    for cell_column in sheet[moved_left_corner: moved_right_corner]:
        for cell in cell_column:
            moved_value = cell.value
            sheet.cell(row=moving_row, column=moving_column,  value=moved_value)
            cell.value = ''
            moving_row += 1
        moving_column += 1


def copy(sheet: Worksheet, copied_left_corner: str, copied_right_corner: str, added_left_corner: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, copied_left_corner) is None \
            or re.match(cell_pattern, copied_right_corner) is \
            None or re.match(cell_pattern, added_left_corner) is None:
        return 'Coordinate error'
    copy_column = sheet[added_left_corner].column
    copy_row = sheet[added_left_corner].row
    for cell_column in sheet[copied_left_corner: copied_right_corner]:
        for cell in cell_column:
            copied_value = cell.value
            sheet.cell(row=copy_row, column=copy_column,  value=copied_value)
            copy_row += 1
        copy_column += 1


def delete(sheet: Worksheet, left_corner: str, right_corner: str):
    cell_pattern = re.compile('^[A-Z]+[1-9]{1}[0-9]*$')
    if re.match(cell_pattern, left_corner) is None \
            or re.match(cell_pattern, right_corner) is None:
        return 'Coordinate error'
    for cell_column in sheet[left_corner: right_corner]:
        for cell in cell_column:
            cell.value = ''


def compare(sheet: Worksheet, cell_1: str, cell_2: str, result_cell: str):
    value_1 = sheet[cell_1].value
    value_2 = sheet[cell_2].value
    if value_1 == value_2:
        sheet[result_cell] = True
        return True
    else:
        sheet[result_cell] = False
        return False


def find(sheet: Worksheet, element: str):
    similar_results = []
    for cell_object in sheet:
        for cell in cell_object:
            if element == str(cell.value) or element in str(cell.value):
                similar_results.append(cell.coordinate)
    if len(similar_results) != 0:
        return similar_results
