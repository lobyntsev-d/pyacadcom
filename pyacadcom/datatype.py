"""
    pyacadcom.datatype
    ******************

    Datatypes for AutoCAD automation

    :copyright: (c) 2017 by Dmitriy Lobyntsev
    :licence: BSD
"""

import win32com.client
import operator
from pythoncom import VT_R8, VT_ARRAY

class acadPoint:
    """
    Class that represent AutoCAD point with x,y,z coordinates
    
    New point:
        >>p1 = acadPoint(50, 100.0, 10)
        (x=50, y=100.0, z=10)
        >>p2 = acadPoint((25, 50.0, 0))
        (x=25, y=50.0, z=0)
        >>p3 = acadPoint([10, 10, 10])
        (x=10, y=10, z=10)
        >>p4 = acadPoint(p1)
        (x=50, y=100.0, z=10)
    Supports math operations: `+`, `-`, `*`, `/`, `+=`, `-=`, `*=`, `/=`:
        >> p1 + p2
        (x=75, y=150.0, z=10)
        >> p1 - p2
        (x=25, y=50.0, z=10)
        >> p3 + 5
        (x=15, y=15, z=15)
        >> p3 * 2
        (x=20, y=20, z=20)
        >> p3/=5
         (x=2, y=2, z=2)
        >>p3*p4
        (x=500, y=1000.0, z=100)
    Can be converted to (tuple) or [list]:
        >> tuple(p1)
        (50, 100.0, 10)
        >> list(p1)
        [50, 100.0, 10]
    .coordinates attribute or call return coordinates  as variant array of doubles to be used in AutoCAD:
    >>p1.coordinates
    (50.0, 100.0, 10.0) - variant array of doubles
    >>p1()
    (50.0, 100.0, 10.0) - variant array of doubles
    """

    def __init__(self, *args):
        coords = []
        if len(args) == 1:
            if isinstance(args[0], (list, tuple)) and len(args[0]) == 3:
                coords = [item for item in args[0] if isinstance(item, (float, int))]
            elif isinstance(args[0], acadPoint):
                coords = [args[0].x, args[0].y, args[0].z]
        elif len(args) == 3:
            coords = [item for item in args if isinstance(item, (float, int))]
        if len(coords) != 3:
            raise TypeError("args in acadPoint(args) can be:"
                            "list, tuple of three float/int, or three float/int, or acadPoint object")
        else:
            self.x, self.y, self.z = coords

    def __call__(self):
        return self.coordinates

    @property
    def coordinates(self):
        return convertcoordinates(self.x, self.y, self.z)

    def __iter__(self):
        self.__i = 0
        return self

    def __next__(self):
        if self.__i == 0:
            self.__i += 1
            return self.x
        elif self.__i == 1:
            self.__i += 1
            return self.y
        elif self.__i == 2:
            self.__i += 1
            return self.z
        else:
            raise StopIteration

    def __str__(self):
        return "(x={}, y={}, z={})".format(self.x, self.y, self.z)

    def __repr__(self):
        return "acadPoint(x={}, y={}, z={})".format(self.x, self.y, self.z)

    def __eq__(self, other):
        if isinstance(other, (acadPoint, list, tuple)):
            return tuple(self) == tuple(other)
        else:
            raise TypeError("acadPoint can be compared with acadPoint, list or tuple only")

    def __add__(self, other):
        return self.__operand(self, other, operator.add)

    def __radd__(self, other):
        return self.__operand(other, self, operator.add)

    def __iadd__(self, other):
        return self.__ioperand(other, operator.add)

    def __sub__(self, other):
        return self.__operand(self, other, operator.sub)

    def __rsub__(self, other):
        return self.__operand(other, self, operator.sub)

    def __isub__(self, other):
        return self.__ioperand(other, operator.sub)

    def __mul__(self, other):
        return self.__operand(self, other, operator.mul)

    def __rmul__(self, other):
        return self.__operand(other, self, operator.mul)

    def __imul__(self, other):
        return self.__ioperand(other, operator.mul)

    def __truediv__(self, other):
        return self.__operand(self, other, operator.truediv)

    def __rtruediv__(self, other):
        return self.__operand(other, self, operator.truediv)

    def __itruediv__(self, other):
        return self.__ioperand(other, operator.truediv)

    def __operand(self, point1, point2, operation):
        points = [point1, point2]
        coordinates = []
        for i in range(2):
            if isinstance(points[i], acadPoint):
                coordinates.append([points[i].x, points[i].y, points[i].z])
            elif isinstance(points[i], (list, tuple)) and len(points[i]) == 3:
                coordinates.append([points[i][0], points[i][1], points[i][2]])
            elif isinstance(points[i], (int, float)):
                coordinates.append([points[i], points[i], points[i]])
            else:
                raise TypeError("Incorrect type. Only int, float, [x,y,z], (x,y,z) or acadPoint can be operated acadPoint")
        return acadPoint(operation(coordinates[0][0], coordinates[1][0]), operation(coordinates[0][1], coordinates[1][1]),
                         operation(coordinates[0][2], coordinates[1][2]))

    def __ioperand(self, point, operation):
        if isinstance(point, acadPoint):
            coordinates = [point.x, point.y, point.z]
        elif isinstance(point, (list, tuple)):
            if len(point) == 3:
                coordinates = [point[0], point[1], point[2]]
            elif len(point) == 2:
                coordinates = [point[0], point[1], 0]
            else:
                raise TypeError("Incorrect length. List or tuple size should be from 2 to 3")
        elif isinstance(point, (int, float)):
            coordinates = [point, point, point]
        else:
            raise TypeError("Incorrect type. Only int, float, [x,y,z], (x,y,z) or acadPoint can be added to acadPoint")
        self.x = operation(self.x, coordinates[0])
        self.y = operation(self.y, coordinates[1])
        self.z = operation(self.z, coordinates[2])
        return self


def convertcoordinates(*args):
    """
    Convert set of numbers to array of doubles to use as coordinates in AutoCAD
    :param args: coordinates
    :return: VARIANT array of doubles (x,y,z)
    """
    return win32com.client.VARIANT(VT_ARRAY | VT_R8, args)