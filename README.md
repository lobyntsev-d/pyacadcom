pyacadcom - Pyton automation of AutoCAD
------
Pyton scripts coding for AutoCAD via ActiveX Automation  
v.0.0.10

Requires:
------
- pywin32

Features:
------
- Solves connection to Autocad COM
- Simplifies work with coordinates
- Simplifies user input

Simple usage example:
------
```Python
    from pyacadcom import AutoCAD, AcadPoint
    acad = AutoCAD()
    point1 = AcadPoint(25, 50, 0)
    nextpoint = (100,100,10)
    point2 = AcadPoint(nextpoint)
    acad.ActiveDocument.ModelSpace.AddLine(point1(), point2())
    point3 = point1 + point2
    acad.ActiveDocument.ModelSpace.AddLine(point2(), point3())

```
AcadPoint() is equal to AcadPoint.coordinates and returns variant array of doubles of x, y, z coordinates
AcadPoint.coordinates2D returns variant array of doubles of x, y coordinates

Links
------
- **Source code and issue tracking** at https://github.com/lobyntsev-d/pyacadcom
- **pywin32:** http://pypi.python.org/pypi/pywin32


Version history
------
0.0.10:
- class acadPoint renamed to AcadPoint
- property coordinates2D is added to AcadPoint class

0.0.9: new types added for wrapping

0.0.8: fix attribute errors cases

0.0.7: new COM wrapper

0.0.6: bugfix

0.0.5: bugfix

0.0.4: user input methods

0.0.3: service methods

0.0.2: AcadPoint class

0.0.1: first release
