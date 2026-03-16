import win32com.client


class SAPClient:

    def __init__(self):
        self.session = self._connect()

    def _connect(self):

        SapGuiAuto = win32com.client.GetObject("SAPGUI")

        if not SapGuiAuto:
            raise Exception("SAP GUI is not running")

        application = SapGuiAuto.GetScriptingEngine

        # Loop through connections and sessions safely
        for connection in application.Children:
            for session in connection.Children:
                return session

        raise Exception("No active SAP session found. Log into SAP first.")