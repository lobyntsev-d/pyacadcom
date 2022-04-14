"""
    pyacadcom.api
    ******************

    Main AutoCAD object through COM automation using pywin32
    :copyright: (c) 2022 by Dmitriy Lobyntsev
    :licence: BSD
    big thanks to xlwings team for wrapper realisation idea
"""

import win32com.client
from pywintypes import com_error, Dispatch, CoClassBaseClass, CDispatch, DispatchEx, DispatchBaseClass
import types
from time import sleep

_DELAY = 0.1  # seconds: delay step incremental
_TIMEOUT = 30.0  # seconds: maximum delay
_CATCH_ATTRIBUTE_ERRORS = False  # Just for test purposes if pywin32 raises AttributeError instead of COM error


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
                elif type(v) is types.MethodType:
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except com_error as error:
                if error.hresult != -2147418111:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)
            except AttributeError:
                if _CATCH_ATTRIBUTE_ERRORS:
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
            except AttributeError:
                if _CATCH_ATTRIBUTE_ERRORS:
                    delay_time += _DELAY
                    sleep(delay_time)

    def __getattr__(self, item):
        delay_time = 0
        while delay_time <= _TIMEOUT:
            try:
                v = getattr(self._inner, item)
                if isinstance(v, (CDispatch, CoClassBaseClass, DispatchBaseClass)):
                    return COMRetryObjectWrapper(v)
                elif type(v) is types.MethodType:
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except com_error as error:
                if error.hresult != -2147418111:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)
            except AttributeError:
                if _CATCH_ATTRIBUTE_ERRORS:
                    delay_time += _DELAY
                    sleep(delay_time)

    def __call__(self, *args, **kwargs):
        delay_time = 0
        while delay_time <= _TIMEOUT:
            try:
                v = self._inner(*args, **kwargs)
                if isinstance(v, (CDispatch, CoClassBaseClass, DispatchBaseClass)):
                    return COMRetryObjectWrapper(v)
                elif type(v) is types.MethodType:
                    return COMRetryMethodWrapper(v)
                else:
                    return v
            except com_error as error:
                if error.hresult != -2147418111:
                    raise
                else:
                    delay_time += _DELAY
                    sleep(delay_time)
            except AttributeError:
                if _CATCH_ATTRIBUTE_ERRORS:
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
        self.__app = COMRetryObjectWrapper(win32com.client.dynamic.Dispatch("AutoCAD.Application"))
        if Visible:
            self.__app.Visible = True

    def __repr__(self):
        return repr(self.__app)

    def __setattr__(self, key, value):
        return setattr(self.__app, key, value)

    def __getattr__(self, item):
        return getattr(self.__app, item)

    def __call__(self, *args, **kwargs):
        return self.__app(*args, **kwargs)

    def __iter__(self):
        for result in self.__app:
            yield result




























def _com_call_wrapper(f, *args, **kwargs):
    """
    COMWrapper support function.
    Repeats calls when 'Call was rejected by callee.' exception occurs.
    """
    # Unwrap inputs
    args = [arg._wrapped_object if isinstance(arg, ComWrapper) else arg for arg in args]
    kwargs = dict([(key, value._wrapped_object)
                   if isinstance(value, ComWrapper)
                   else (key, value)
                   for key, value in dict(kwargs).items()])
    start_time = None
    while True:
        try:
            result = f(*args, **kwargs)
        except com_error as e:
            if e.hresult == -2147418111:
                if start_time is None:
                    start_time = time.time()

                elif time.time() - start_time >= _TIMEOUT:
                    raise

                sleep(_DELAY)
                continue

            raise

        break

    if isinstance(result, win32com.client.CDispatch) or callable(result):
        return ComWrapper(result)
    return result


class ComWrapper(object):
    """
    Class to wrap COM objects to repeat calls when 'Call was rejected by callee.' exception occurs.
    """

    def __init__(self, wrapped_object):
        assert isinstance(wrapped_object, win32com.client.CDispatch) or callable(wrapped_object)
        self.__dict__['_wrapped_object'] = wrapped_object

    def __getattr__(self, item):
        return _com_call_wrapper(self._wrapped_object.__getattr__, item)

    def __getitem__(self, item):
        return _com_call_wrapper(self._wrapped_object.__getitem__, item)

    def __setattr__(self, key, value):
        _com_call_wrapper(self._wrapped_object.__setattr__, key, value)

    def __setitem__(self, key, value):
        _com_call_wrapper(self._wrapped_object.__setitem__, key, value)

    def __call__(self, *args, **kwargs):
        return _com_call_wrapper(self._wrapped_object.__call__, *args, **kwargs)

    def __repr__(self):
        return 'ComWrapper<{}>'.format(repr(self._wrapped_object))


class AutoCAD:
    """
    Class that represent AutoCAD program via COM
    """

    def __init__(self, Visible=True):
        """
        Initiation of AutoCAD application
        :param Visible: States if AutoCAD application to be visible, visible by default
        """
        self._app_unwrapped = win32com.client.Dispatch("AutoCAD.Application")
        self.__app = ComWrapper(self._app_unwrapped)
        if Visible == True:
            self.__app.Visible = True
    @property
    def app(self):
        """
        :return: COM "AutoCAD.Application" active instance
        """
        return self.__app

    @property
    def Documents(self):
        """
        :return: Documents collection of active AutoCAD application
        """
        return self.__app.Documents

    @property
    def ActiveDocument(self):
        """
        :return: Active document object of AutoCAD application
        """
        return self.__app.ActiveDocument

    @property
    def Preferences(self):
        """
        :return: Settings object of AutoCAD application
        """