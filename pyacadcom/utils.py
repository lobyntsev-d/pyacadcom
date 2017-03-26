"""

    pyacadcom.utils
    ******************

    Utilities for AutoCAD automation

    :copyright: (c) 2017 by Dmitriy Lobyntsev
    :licence: BSD

"""
from math import sqrt

def correct_point(obj):
    """
    Function to check if object can be used as point coordinates in 3D Decart's system
    :param obj: object to check
    :return: True or False
    """
    check = tuple(obj)
    if len(check) == 3 and all([isinstance(item, (int, float)) for item in check]):
        return True
    else:
        return False

def distance(point1, point2):
    """
    Function to calculate distance between two points in 3D Decart's system
    :param point1: acadPoint object or coordinates as list [x,y,z] or tuple (x,y,z)
    :param point2: acadPoint object or coordinates as list [x,y,z] or tuple (x,y,z)
    :return: distance between point1 and point2
    """
    if not correct_point(point1) or not correct_point(point2):
        raise TypeError("Incorrect type of points")
    a = tuple(point1)
    b = tuple(point2)
    return sqrt((a[0]-b[0])**2+(a[1]-b[1])**2+(a[2]-b[2])**2)