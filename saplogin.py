import win32com.client
import subprocess
import time


class Connect:
    '''The intention of this code is to create or connect
    to a SAP connection and return a list of the present sessions.'''

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
        self.connection = None

    def open_sap(self, user, password, language, connection='PRD [R3 PRODUCCION]'):
        self.user = user
        self.password = password
        self.language = language
        self.connection = connection  # Connects to PRD by default
        self.process = subprocess.Popen(self.path)
        time.sleep(4)  # time sleep so the computer has time to open SAP.

        # Connecting to the SAP API
        self.sapguiauto = win32com.client.GetObject('SAPGUI')
        if not type(self.sapguiauto) == win32com.client.CDispatch:
            return
        self.application = self.sapguiauto.GetScriptingEngine
        if not type(self.application) == win32com.client.CDispatch:
            self.sapguiauto = None
            return

        # Checks if user is already logged in SAP in the computer used. If so, it uses the current connection
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
        self.connection = self.application.Openconnection(self.connection, True)

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

    # Verifies the amount of open sessions, open a new one and appends it to the list os sessions.
    def new_session(self):

        open_sessions = len(self.connection.Children)
        self.session[0].createsession()
        time.sleep(1)
        self.session.append(self.connection.Children(open_sessions))
        return

    def disconnect(self):
        self.connection.CloseConnection()

    # Forces entry when we are facing warning messages while trying to go to the next step.
    def force_entry(self, session_num):
        while self.session[session_num].findById("wnd[0]/sbar/").text != '':
            self.session[session_num].findById("wnd[0]").sendVKey(0)

    # Forces entry on possible warning Popup screens.
    def force_popup(self, session_num):
        while True:
            try:
                self.session[session_num].findById("wnd[1]").sendVKey(0)
            except Exception:
                break


        
