import win32com.client
import subprocess
import time

class Connect:
    def __init__(self):
        self.process = None
        self.sapguiauto = None
        self.application = None
        self.connection = None
        self.session = []
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        self.user = None
        self.password = None
        self.language = None
        self.server_name = None

    def open_sap(self, user, password, language, server_name):
        self.user = user
        self.password = password
        self.language = language
        self.server_name = server_name

        self.process = subprocess.Popen(self.path)
        time.sleep(4)

        self.sapguiauto = win32com.client.GetObject('SAPGUI')
        if not type(self.sapguiauto) == win32com.client.CDispatch:
            return

        self.application = self.sapguiauto.GetScriptingEngine
        if not type(self.application) == win32com.client.CDispatch:
            self.sapguiauto = None
            return

        # Checks if user is already logged in SAP in the computer used. If so, it uses the open connection
        if len(self.application.Children) > 0:
            for con in range(0, len(self.application.Children)):
                self.connection = self.application.Children(con)
                if not type(self.connection) == win32com.client.CDispatch:
                    self.application = None
                    self.sapguiauto = None
                    return
                if self.connection.Children(0).Info.User == self.user:
                    for i in range(0, len(self.connection.Children)):
                        self.session.append(self.connection.Children(i))
                    return
            self.login()
            return

        # If not, creates a new login session
        else:
            self.login()
            return

    def login(self):
        self.connection = self.application.Openconnection(self.server_name, True)

        if not type(self.connection) == win32com.client.CDispatch:
            self.application = None
            self.sapguiauto = None
            return

        self.session.append(self.connection.Children(0))
        self.session[0].findById("wnd[0]/usr/txtRSYST-BNAME").text = self.user
        self.session[0].findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password
        self.session[0].findById("wnd[0]/usr/txtRSYST-LANGU").text = self.language
        self.session[0].findById("wnd[0]").sendVKey(0)
        while len(self.session[0].Children) == 2:
            try:
                self.session[0].findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select()
            except:
                pass
            self.session[0].findById("wnd[1]").sendVKey(0)

    def new_session(self):
        """Verifies the amount of sessions open, opens a new one and appends it to the pre existing sessions"""

        open_sessions = len(self.connection.Children)
        self.session[0].createsession()
        time.sleep(1)
        self.session.append(self.connection.Children(open_sessions))
        return

    def disconnect(self):
        self.connection.CloseConnection()

        
