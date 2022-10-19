import openpyxl


def addition(*addends):
    summa = 0
    for addend in addends:
        if isinstance(addend, int) or isinstance(addend, float):
            summa += addend
    return summa


def subtraction(redused, substracted):
    if not (isinstance(redused, int) or isinstance(redused, float) or
            isinstance(substracted, int) or isinstance(substracted, float)):
        return 'N/A'
    difference = redused - substracted
    return difference


def multiplication(*factors):
    composition = 1
    for factor in factors:
        if isinstance(factor, int) or isinstance(factor, float):
            composition *= factor
    return composition


def division(divisible, divisor):
    if not (isinstance(divisible, int) or isinstance(divisible, float) or
            isinstance(divisor, int) or isinstance(divisor, float)):
        return 'N/A'
    quotient = divisible / divisor
    return quotient

def round_number(number, decimal):
    if (isinstance(number, int) or isinstance(number, float)) and isinstance(decimal, int):
        return round(number, decimal)
    return number

def move():
    pass
def copy():
    pass
def delete():
    pass
def find():
    pass