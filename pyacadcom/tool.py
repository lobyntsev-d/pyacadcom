
def double_from_string(str):
    """
    Функция, возвращающая число с плавающей точкой из строки.
    Предусмотрена обработка разделителя в виде запятой или точки.
    Предусмотрена обработка ошибок
    :param str: строка для преобразования
    :return: (Resultcode, Resultvalue)
                Resultcode: тип возвращаемого результата, 1 - конвертация прошла успешно, отрицательные числа - ошибки
                Resultvalue: возвращаемый результат - число или описание ошибки
    """
    str = str.replace(",", ".")
    try:
        return (1, float(str))
    except ValueError:
        return (-1, "not a double in string")

def int_from_string(str):
    """
    Функция, возвращающая челое число из строки.
    Предусмотрена обработка ошибок
    :param str: строка для преобразования
    (Resultcode, Resultvalue)
                Resultcode: тип возвращаемого результата, 1 - конвертация прошла успешно, отрицательные числа - ошибки
                Resultvalue: возвращаемый результат - число или описание ошибки
    """
    try:
        return (1, int(str))
    except ValueError:
        (-1, "not a int in string")

def sum_length(selset):
    """
    Функция подсчёта суммы длин линий, полилиний и дуг в выборке
    :param selset: выборка объектов типа Autocad SelectionSet
    :return: сумма длин
    """
    if selset.Count == 0:
        return (-2, "Ничего не выбрано")
    line_types = ("AcDbPolyline", "AcDbLine")
    total_length = 0.0
    for item in selset:
        if item.ObjectName in line_types:
            total_length += item.Length
        elif item.ObjectName == "AcDbArc":
            total_length += item.ArcLength
        elif item.ObjectName == "AcDbMline":
            total_length += mlinelength(item)
    return (1, total_length)

def distance(coord1, coord2):
    if len(coord1) != len(coord2) or len(coord1) not in (2,3):
        print("ValueError: distance() arguments must be same size (2 or 3 elements)")
        raise ValueError
    dist = 0
    for i in range(len(coord1)):
        dist += (coord1[i] - coord2[i]) ** 2
    return dist ** 0.5

def mlinelength(multiline):
    coordinates = triplecoordinates(multiline.Coordinates)
    length = 0
    for i in range(len(coordinates)-1):
        length += distance(coordinates[i], coordinates[i+1])
    return length

def triplecoordinates(coordinates_set):
    set_length = len(coordinates_set)
    if set_length % 3 == 0:
        result = []
        for i in range(set_length//3):
            point = coordinates_set[i*3:i*3+3:]
            result.append(point)
        return result
    else:
        raise ValueError("triplecoordinates() argument length must be a multiple of three")




