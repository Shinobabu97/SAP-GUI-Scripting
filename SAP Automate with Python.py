pip install pywin32
import win32com.client

# Connect to SAP GUI
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

#SAP steps
