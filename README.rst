pyacadcom - Pyton automation of AutoCAD
----------------------------------------

Library for simplifying writing ActiveX_ Automation_ Pyton scripts for AutoCAD_

Requires:
----------

- pywin32_

Features:
-----------

- Simplifies work with coordinates (3D points)

Simple usage example:
---------------------

.. code-block:: python

    from pyacadcom import AutoCAD, acadPoint

    acad=AutoCAD()
    point1 = acadPoint(25, 50, 0)
    nextpoint = (100,100,10)
    point2 = acadPoint(nextpoint)
    acad.ActiveDocument.ModelSpace.AddLine(point1(), point2())
    point3 = point1 + point2
    acad.ActiveDocument.ModelSpace.AddLine(point2(), point3())

Links
-----

- **Source code and issue tracking** at `GitHub <https://github.com/lobyntsev-d/pyacadcom>`_.

.. _ActiveX: http://wikipedia.org/wiki/ActiveX
.. _Automation: http://en.wikipedia.org/wiki/OLE_Automation
.. _AutoCAD: http://wikipedia.org/wiki/AutoCAD
.. _pywin32: http://pypi.python.org/pypi/pywin32
