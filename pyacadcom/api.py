"""
    pyacadcom.api
    ******************

    Main AutoCAD object through COM automation using pywin32

    :copyright: (c) 2017 by Dmitriy Lobyntsev
    :licence: BSD
"""
import win32com.client
from pywintypes import com_error
import time

_DELAY = 0.05  # seconds
_TIMEOUT = 20.0  # seconds


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

                time.sleep(_DELAY)
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