"""
    pyacadcom.api
    ******************

    Main AutoCAD object through COM automation using pywin32

    :copyright: (c) 2017 by Dmitriy Lobyntsev
    :licence: BSD
"""
import win32com.client

class AutoCAD:
    """
    Class that represent AutoCAD program via COM
    """

    def __init__(self, Visible=True):
        """
        Initiation of AutoCAD application
        :param Visible: States if AutoCAD application to be visible, visible by default
        """
        self.__app = win32com.client.Dispatch("AutoCAD.Application")
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