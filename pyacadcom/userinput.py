"""

    pyacadcom.userinput
    ******************

    Utilities for AutoCAD automation

    :copyright: (c) 2022 by Dmitriy Lobyntsev
    :licence: BSD

"""


from random import randint

from pythoncom import VT_R8, VT_ARRAY, VT_DISPATCH, VT_BSTR, VT_BYREF, com_error

from .tool import double_from_string, int_from_string


def text_input(app, request_type="str", prompt="", options=None, default=None):
    """
    Функция текстового ввода в активном документе
    :param app: экземпляр Autocad
    :param request_type: тип ожидаемого значения
                    "int" - целое
                    "double" - вещественное
                    "str" - строка без пробелов
                    "str_spaced" - строка с пробелами
    :param prompt: текст запроса
    :param options: перечень опций в виде словаря {имя опции: значение, ...}
    :return: Resultcode, Resultvalue
                Resultcode: тип возвращаемого результата:
                                                1 - введены данные
                                                2 - введена опция
                                                -1 - ввод отменен пользователем
                Resultvalue: возвращаемый результат - данные, имя выбранной опции или описание ошибки
    """
    doc = app.ActiveDocument
    if isinstance(options, dict):
        opt = {}
        init_string = " ".join([item.replace(" ", "") for item in options.values()])
        for key, value in options.items():
            opt[value] = key
        if default is not None:
            opt["@defaultvalue@"] = default
            options[default] = "<"+options[default]+">"
        options_string = " [" + "/".join([item for item in options.values()]) + "]"
    else:
        options_string = ""
        init_string = ""

    prompt = "\n" + prompt
    doc.Utility.InitializeUserInput(128, init_string)

    while True:
        try:
            text = doc.Utility.GetString(1 if request_type == "str_spaced" else 0, prompt + options_string)
            if text == "":
                text = "@defaultvalue@"
        except com_error as error:
            if error.hresult == -2147352567:
                doc.Utility.Prompt("Отменено пользователем/Canceled by user\n")
                return -1, "Esc is pressed"
            else:
                raise
        if isinstance(options, dict):
            for option in opt.keys():
                if option.replace(" ","").find(text) != -1:
                    return 2, opt[option]
        if request_type == "str" or request_type == "str_spaced":
            return 1, text
        elif request_type == "int":
            res = int_from_string(text)
            if res[0] >= 0:
                return 1, res[1]
        elif request_type == "double":
            res = double_from_string(text)
            if res[0] >= 0:
                return 1, res[1]

def get_obj(app, obj_type = "all", prompt = ""):
    """
    Функция выбора объектов автокада, с поддержкой фильтрации
    :param app: объект-приложение автокада
    :param obj_type: тип допустимых объектов (регистр не имеет значения), для выбора нескольких типов перечислить их через пробел
                        All - допустим выбор любого объекта
                        "Line", "AcDbLine" - линия
                        "Pline", "Polyline", "pl", "AcDbPolyline" - полилиния
                        "Block", "Blockreference", "bl", "blockref", "AcDbBlockReference" - блок
                        "Arc", "a", "AcDbArc" - дуга
                        ""Mline", "ml", "AcDbMline" - мультилиния
    :param prompt: текст запроса в командной строке
    :return: Resultcode, Resultvalue
                Resultcode: тип возвращаемого результата:
                                                1 - выбран объект/объекты
                                                -1 - ничего не выбрано
                                                -2 - нет объектов, соответствующих допустимому типу
                Resultvalue: возвращаемый результат - список объектов или описание ошибки
    """
    #подключаемся к активному документу
    doc = app.ActiveDocument
    #создаём временный набор и запрашиваем у пользователя добавление в него элементов
    selset = doc.SelectionSets.Add(str(randint(0, 100000)))
    doc.Utility.Prompt(prompt)
    selset.SelectOnScreen()
    #создаём список из элементов набора
    selection = [x for x in selset]
    #удаляем временный набор
    selset.Delete()
    #обрабатываем вариант, при котором ничего не выбрано
    if len(selection) == 0:
        doc.Utility.Prompt("Ничего не выбрано")
        return -1, "No choice"
    obj_type = obj_type.lower()
    if obj_type.lower() != "all":
        # приводим названия типов разрешенных лбъектов к каноническому виду
        allowed_types = [x.lower() for x in obj_type.split(" ")]
        for i, value in enumerate(allowed_types):
            if value in ("line", "acdbline", "l"):
                allowed_types[i] = "AcDbLine"
            elif value in ("pline", "polyline", "pl", "acdbpolyline"):
                allowed_types[i] = "AcDbPolyline"
            elif value in ("block", "blockreference", "bl", "blockref", "acdbblockreference"):
                allowed_types[i] = "AcDbBlockReference"
            elif value in ("arc", "a", "acdbarc"):
                allowed_types[i] = "AcDbArc"
            elif value in ("mleader", "multileader", "acdbmleader"):
                allowed_types[i] = "AcDbMLeader"
            elif value in ("arc", "a", "acdbarc"):
                allowed_types[i] = "AcDbArc"
            elif value in ("mline", "ml", "acdbmline"):
                allowed_types[i] = "AcDbMline"
            #TODO: Добавить все примитивы автокада
            else:
                allowed_types[i] = ""
        #определяем индексы элементов списка, которые проходят по фильтру
        items_to_keep = []
        for i, item in enumerate(selection):
            if item.ObjectName in allowed_types:
                items_to_keep.append(i)
        #определяем элементы, удовлетворяющие фильтру
        selection = [selection[i] for i in items_to_keep]
        #обрабатываем случай, при котором нет элементов удовлетворяющих списку
        if len(selection) == 0:
            return -2, "Не выбрано объектов, соответствующих фильтру"
    return 1, selection

def get_keyword(app, prompt = "", options = None, default = None):
    """
        Функция ввода опций в активном документе
        :param app: экземпляр Autocad
        :param prompt: текст запроса
        :param options: перечень опций в виде словаря {имя опции: значение, ...}
        :return: Resultcode, Resultvalue
                    Resultcode: тип возвращаемого результата:
                                                            -1 - ввод отменен пользователем
                                                            2 - введена опция, отрицательные числа - ошибки
                    Resultvalue: возвращаемый результат - имя выбранной опции или описание ошибки
        """
    doc = app.ActiveDocument
    if isinstance(options, dict):
        opt = {}
        init_string = " ".join([item.replace(" ", "") for item in options.values()])
        for key, value in options.items():
            opt[value] = key
        if default is not None:
            opt["@defaultvalue@"] = default
            options[default] = "<" + options[default] + ">"
        options_string = " [" + "/".join([item for item in options.values()]) + "]"
    else:
        options_string = ""
        init_string = ""
    prompt = "\n" + prompt
    doc.Utility.InitializeUserInput(128, init_string)

    while True:
        try:
            text = doc.Utility.GetKeyword(prompt + options_string)
            if text == "":
                text = "@defaultvalue@"
        except com_error as error:
            if error.hresult == -2147352567:
                doc.Utility.Prompt("Отмена\n")
                return -1, "Esc is pressed"
            else:
                raise
        for option in opt.keys():
            if option.replace(" ", "").find(text) != -1:
                return 2, opt[option]

def dist(app, options=None, default=None):
    """
    Функция для измерения расстояния
    :param app:
    :return:
    """
    doc = app.ActiveDocument
    if isinstance(options, dict):
        opt = {}
        init_string = " ".join([item.replace(" ", "") for item in options.values()])
        for key, value in options.items():
            opt[value] = key
        if default is not None:
            opt["@defaultvalue@"] = default
            options[default] = "<" + options[default] + ">"
        options_string = " [" + "/".join([item for item in options.values()]) + "]"

    else:
        options_string = ""
        init_string = ""

    doc.Utility.InitializeUserInput(128, init_string)

    finish = False
    while not finish:
        try:
            doc.Utility.Prompt("\nВведите первую точку для измерения расстояния" + options_string)
            prevpoint = doc.Utility.GetPoint()
            finish = True
        except com_error as error:
            if error.hresult == -2145320928:
                text = doc.Utility.GetInput()
                if text == "":
                    text = "@defaultvalue@"
            elif error.hresult == -2147352567:
                return -2, "Esc is pressed"
            else:
                raise
            if isinstance(options, dict):
                for option in opt.keys():
                    if option.replace(" ", "").find(text) != -1:
                        return 2, opt[option]
    dist = 0
    finish = False
    while not finish:
        try:
            curpoint = doc.Utility.GetPoint(win32com.client.VARIANT(VT_ARRAY | VT_R8, prevpoint), "Введите следующую точку для измерения расстояния [Ввод для завершения]")
            dist += ((curpoint[0] - prevpoint[0]) ** 2 + (curpoint[1] - prevpoint[1]) ** 2) ** 0.5
            prevpoint = curpoint
        except com_error as error:
            if error.hresult == -2145320928:
                key = doc.Utility.GetInput()
                if key == "" or key.upper() == "В":
                    finish = True
                else:
                    continue
            elif error.hresult == -2147352567:
                return -2, "Esc is pressed"
            else:
                raise
    return 1, dist