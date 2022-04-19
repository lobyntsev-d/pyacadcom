"""
    pyacadcom.api
    ******************

    Main AutoCAD object through COM automation using pywin32
    :copyright: (c) 2022 by Dmitriy Lobyntsev
    :licence: BSD
    big thanks to xlwings team for wrapper realisation idea
"""

from win32com.client.dynamic import CDispatch as dynCDispatch
from win32com.client import Dispatch, CDispatch, CoClassBaseClass, DispatchBaseClass, Constants, EventsProxy
from pywintypes import com_error
from types import MethodType
from time import sleep

_DELAY = 0.05  # seconds: default delay step incremental
_TIMEOUT = 15.0  # seconds: default maximum delay
_ERRORCODES = [-2147418111, -2147417847, -2147417846]
_TYPES_TO_WRAP = (CDispatch, CoClassBaseClass, DispatchBaseClass, dynCDispatch, Constants, EventsProxy)

class COMRetryMethodWrapper:

    def __init__(self, method):
        self.__method = method

    def __call__(self, *args, **kwargs):
        delay_time = 0
        while delay_time <= _TIMEOUT:
            try:
                i = self.__method(*args, **kwargs)
                if isinstance(i, _TYPES_TO_WRAP):
                    return COMRetryObjectWrapper(i)
                elif type(i) is MethodType:
                    return COMRetryMethodWrapper(i)
                else:
                    return i
            except com_error as error:
                if error.hresult not in _ERRORCODES:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)
            except AttributeError:
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
                if error.hresult not in _ERRORCODES:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)
            except AttributeError:
                delay_time += _DELAY
                sleep(delay_time)

    def __getattr__(self, item):
        delay_time = 0
        while delay_time <= _TIMEOUT:
            try:
                i = getattr(self._inner, item)
                if isinstance(i, _TYPES_TO_WRAP):
                    return COMRetryObjectWrapper(i)
                elif type(i) is MethodType:
                    return COMRetryMethodWrapper(i)
                else:
                    return i
            except com_error as error:
                if error.hresult not in _ERRORCODES:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)
            except AttributeError:
                try:
                    self._oleobj_.GetIDsOfNames(0, item)
                except com_error as error:
                    if error.hresult == -2147352570:
                        raise
                delay_time += _DELAY
                sleep(delay_time)

    def __call__(self, *args, **kwargs):
        delay_time = 0
        while delay_time <= _TIMEOUT:
            try:
                i = self._inner(*args, **kwargs)
                if isinstance(i, _TYPES_TO_WRAP):
                    return COMRetryObjectWrapper(i)
                elif type(i) is MethodType:
                    return COMRetryMethodWrapper(i)
                else:
                    return i
            except com_error as error:
                if error.hresult not in _ERRORCODES:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)
            except AttributeError:
                delay_time += _DELAY
                sleep(delay_time)

    def __iter__(self):
        for i in self._inner:
            if isinstance(i, _TYPES_TO_WRAP):
                yield COMRetryObjectWrapper(i)
            else:
                yield i


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
        for i in self._inner:
            yield i

