# sapguilogin
My personal Code for logging in and running Scripts in SAP Netweaver GUI with Python.

If you want to use this code, you should change the following path to the path of your SAP logon file.
//self.path = r"***\SAP\FrontEnd\SAPgui\saplogon.exe"

The Connect class should be first instantiated and then you should pass the necessary parameters "user", "password", "language", "connection='PRD [R3 PRODUCCION]'". If no parameter is passed to the connection, it will try to connect to the PRD (not QA).

The code will open SAP and then check whether the account is already logged in or not. If the account isn't logged, it will initiate a new connection, otherwise, it will connect to the preexisting connection.

It will then generate a list of sessions (this list property is usually what you are going to work on from this point onward). If there was no preexisting connection, the code will create a single session that should be accessed using object.session[0]

Usually from this point on the desirable code should come from recording manually made actions on SAP, then adapt to Python the VB code that SAP generates.
If you need information that cannot be accessed by recording SAP scriptings, please download the API DOC from the official website.

https://help.sap.com/doc/9215986e54174174854b0af6bb14305a/760.01/en-US/sap_gui_scripting_api_761.pdf

Methods/Functions:

new_session(): creates a new session in the connection and append it to the list of sessions.

force_entry(): From my experience, the code gets stuck very often because of warning messages from SAP (usually date related warnings). This method will try to force SAP into going to the next window/step.
Be careful! You should only use this if you're sure about what type of warning you might get.

force_popup(): The same from force_entry, with the difference that it will force popped up warning screens. 

disconnect(): ends the connection of the created object.
