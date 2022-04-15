"""
    pyacadcom.api
    ******************

    Main AutoCAD object through COM automation using pywin32
    :copyright: (c) 2022 by Dmitriy Lobyntsev
    :licence: BSD
    big thanks to xlwings team for wrapper realisation idea
"""

from win32com.client.dynamic import Dispatch
from win32com.client import CDispatch, CoClassBaseClass, DispatchBaseClass
from pywintypes import com_error
from types import MethodType
from time import sleep

_DELAY = 0.05  # seconds: default delay step incremental
_TIMEOUT = 10.0  # seconds: default maximum delay

class COMRetryMethodWrapper:

    def __init__(self, method):
        self.__method = method

    def __call__(self, *args, **kwargs):
        delay_time = 0
        while delay_time <= _TIMEOUT:
            try:
                v = self.__method(*args, **kwargs)
                if isinstance(v, (CDispatch, CoClassBaseClass, DispatchBaseClass)):
                    return COMRetryObjectWrapper(v)
                elif type(v) is MethodType:
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except com_error as error:
                if error.hresult != -2147418111:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)


class COMRetryObjectWrapper:
    def __init__(self, inner):
        object.__setattr__(self, "_inner", inner)

    def __repr__(self):
        return repr(self._inner)

    def __setattr__(self, key, value):
        delay_time = 0
        while delay_time <= _TIMEOUT:
            try:
                return setattr(self._inner, key, value)
            except com_error as error:
                if error.hresult != -2147418111:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)


    def __getattr__(self, item):
        delay_time = 0
        while delay_time <= _TIMEOUT:
            try:
                v = getattr(self._inner, item)
                if isinstance(v, (CDispatch, CoClassBaseClass, DispatchBaseClass)):
                    return COMRetryObjectWrapper(v)
                elif type(v) is MethodType:
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except com_error as error:
                if error.hresult != -2147418111:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)


    def __call__(self, *args, **kwargs):
        delay_time = 0
        while delay_time <= _TIMEOUT:
            try:
                v = self._inner(*args, **kwargs)
                if isinstance(v, (CDispatch, CoClassBaseClass, DispatchBaseClass)):
                    return COMRetryObjectWrapper(v)
                elif type(v) is MethodType:
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except com_error as error:
                if error.hresult != -2147418111:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)


    def __iter__(self):
        for v in self._inner:
            if isinstance(v, (CDispatch, CoClassBaseClass, DispatchBaseClass)):
                yield COMRetryObjectWrapper(v)
            else:
                yield v


class AutoCAD:
    """
    Class that represent AutoCAD program via COM
    """
    def __init__(self, Visible=True):
        """
        Initiation of AutoCAD application
        :param Visible: States if AutoCAD application to be visible, visible by default
        """
        object.__setattr__(self, "_inner", COMRetryObjectWrapper(Dispatch("AutoCAD.Application")))
        if Visible:
            self._inner.Visible = True

    def __repr__(self):
        return repr(self.self._inner)

    def __setattr__(self, key, value):
        return setattr(self._inner, key, value)

    def __getattr__(self, item):
        return getattr(self._inner, item)

    def __call__(self, *args, **kwargs):
        return self._inner(*args, **kwargs)

    def __iter__(self):
        for result in self._inner:
            yield result