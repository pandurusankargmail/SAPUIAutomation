import win32com.client
import subprocess
import time
from dotenv import load_dotenv
import os


def close_sap_logon_window():
    # /nex closes active SAP sessions but can leave the SAP Logon launcher open.
    for process_name in ["saplogon.exe", "saplgpad.exe"]:
        subprocess.run(
            ["taskkill", "/IM", process_name, "/F", "/T"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )



def login_sap(username, password, sap_gui_path):
    # Start the SAP GUI application
    subprocess.Popen(sap_gui_path)
    
    # Wait for the SAP GUI to start
    time.sleep(8)  # Adjust this delay as needed
    
    # Connect to the SAP GUI application
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    connection = application.OpenConnection("26: LNX Dry Run [T00]", True)  # Replace with your SAP system name
    
    # # Get the first session (assuming only one session is open)
    session = application.Children(0).Children(0)
    session.findById("wnd[0]").maximize()  # Maximize the SAP GUI window
    
    # Enter the username and password
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "020"  # Set client (adjust if needed)
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EN"  # Set language to English (adjust if needed)
    session.findById("wnd[0]/tbar[0]/btn[0]").press()  # Click the login button
    return session


def logoff_and_exit_sap(session):
    # /nex logs off from SAP and closes all sessions for the current connection.
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)
    close_sap_logon_window()


if __name__ == "__main__":
    load_dotenv(dotenv_path="credentials.env")  # Load variables from credentials.env
    # Replace these with your SAP credentials and the path to your SAP GUI executable
    username = os.getenv("username")  # Load username from environment variable
    password = os.getenv("password")  # Load password from environment variable
    
    sap_gui_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"  # Adjust this path as needed
    
    session = login_sap(username, password, sap_gui_path)
    time.sleep(2)  # Keep session open briefly after login
    logoff_and_exit_sap(session)
