# SAPUIAutomation
# SAP UI Automation (Python + pywin32)

This project automates SAP GUI login and logoff by using the SAP GUI Scripting COM interface from Python.

## Required Python modules

Install these packages before running the script:

```bash
pip install pywin32 python-dotenv
```

Module summary:

1. pywin32
	- Provides access to Windows COM automation from Python.
	- Used to connect to SAP GUI scripting objects (SAPGUI, Application, Connection, Session).
	- In this project, it is imported as win32com.client.

2. python-dotenv
	- Loads environment variables from a .env style file into the process.
	- In this project, load_dotenv reads credentials.env so username/password are not hardcoded.

## SAP GUI prerequisites

Before running automation, ensure the following:

1. SAP GUI for Windows is installed.
2. SAP Logon entry exists for your target system.
3. SAP GUI scripting is enabled:
	- On client: SAP Logon Options -> Accessibility & Scripting -> Enable scripting.
	- On server: scripting must also be allowed by BASIS/admin policy.

If scripting is disabled, COM calls can fail even if credentials are correct.

## Credentials setup with python-dotenv

Create or update credentials.env in this folder.

Example:

```env
username=YOUR_SAP_USERNAME
password=YOUR_SAP_PASSWORD
```

How it is used in the script:

1. load_dotenv(dotenv_path="credentials.env") loads key-value pairs.
2. os.getenv("username") and os.getenv("password") read the values.
3. Values are then entered into SAP login fields.

## Run the automation

From this project folder:

```bash
python LoginSAP.py
```

## What the script does

1. Starts SAP Logon (saplogon.exe).
2. Opens the configured SAP system connection.
3. Enters client, username, password, and language.
4. Logs in.
5. Sends /nex to log off and close active session(s).

## Notes

1. Update the SAP system name in the script if your environment is different.
2. Update the SAP GUI executable path if SAP is installed in another location.
3. Never commit real credentials to source control.





