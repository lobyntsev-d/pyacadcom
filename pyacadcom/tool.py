"""

    pyacadcom.tool
    ******************

    Utilities for AutoCAD automation

    :copyright: (c) 2022 by Dmitriy Lobyntsev
    :licence: BSD

"""

def double_from_string(str):
    """
    Функция, возвращающая число с плавающей точкой из строки.
    Предусмотрена обработка разделителя в виде запятой или точки.
    :param str: строка для преобразования
    :return: возвращаемый результат - число
    """
    str = str.replace(",", ".")
    try:
        return float(str)
    except ValueError:
        return "not a double in string"

def int_from_string(str):
    """
    Функция, возвращающая челое число из строки.
    :param str: строка для преобразования
    :return: возвращаемый результат - число
    """
    try:
        int(str)
    except ValueError:
        "not a int in string"

def sum_length(selset):
    """
    Функция подсчёта суммы длин линий, полилиний, мультилиний и дуг в выборке
    :param selset: выборка объектов типа Autocad SelectionSet
    :return: сумма длин
    """
    if selset.Count == 0:
        raise ValueError("No elements are chosen")
    line_types = ("AcDbPolyline", "AcDbLine")
    total_length = 0.0
    for item in selset:
        if item.ObjectName in line_types:
            total_length += item.Length
        elif item.ObjectName == "AcDbArc":
            total_length += item.ArcLength
        elif item.ObjectName == "AcDbMline":
            total_length += mlinelength(item)
    return total_length

def distance(coord1, coord2):
    """
    Функция вычисления расстояния между двумя точками
    :param coord1: координаты точки 1
    :param coord2: координаты точки 2
    :return:
    """
    if len(coord1) != len(coord2) or len(coord1) not in (2, 3):
        raise ValueError("distance() arguments must be same size (2 or 3 elements)")
    dist = 0
    for i in range(len(coord1)):
        dist += (coord1[i] - coord2[i]) ** 2
    return dist ** 0.5

def mlinelength(multiline):
    """
    Функция вычисления длины мультилинии
    :param multiline: объект мультилинии
    :return: длина мультилинии
    """
    coordinates = triplecoordinates(multiline.Coordinates)
    length = 0
    for i in range(len(coordinates)-1):
        length += distance(coordinates[i], coordinates[i+1])
    return length

def triplecoordinates(coordinates_set):
    """
    Функция получения координат в формате [[X1,Y1,Z1],[X2,Y2,Z2],[X3,Y3,Z3]] из формата [X1,Y1,Z1,X2,Y2,Z2,X3,Y3,Z3]
    :param coordinates_set: набор координат в формате [X1,Y1,Z1,X2,Y2,Z2,X3,Y3,Z3]
    :return: набор координат в формате [[X1,Y1,Z1],[X2,Y2,Z2],[X3,Y3,Z3]]
    """
    set_length = len(coordinates_set)
    if set_length % 3 == 0:
        result = []
        for i in range(set_length//3):
            point = coordinates_set[i*3:i*3+3:]
            result.append(point)
        return result
    else:
        raise ValueError("triplecoordinates() argument length must be a multiple of three")




