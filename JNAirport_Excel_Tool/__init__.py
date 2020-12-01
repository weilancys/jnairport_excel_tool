from .ui import JNAirport_Excel_Tool
import os


def makesure_app_dir():
    home_dir = os.path.expanduser("~")
    config_dir = os.path.join(home_dir, ".jnairport_excel_tool")
    os.makedirs(config_dir, exist_ok=True)
    return config_dir


def get_smtp_credentials(config_dir):
    filename = "smtp_credentials.txt"
    file_path = os.path.join(config_dir, filename)

    try:
        for line in open(file_path, encoding="utf-8"):
            line = line.strip()
            if line.startswith("#") or line == "":
                continue
            if line.startswith("SENDER_EMAIL"):
                SENDER_EMAIL = line.split("=")[1].strip()
            elif line.startswith("PASSWORD"):
                PASSWORD = line.split("=")[1].strip()
            elif line.startswith("SMTP_SERVER_ADDR"):
                SMTP_SERVER_ADDR = line.split("=")[1].strip()
            elif line.startswith("SMTP_SERVER_PORT"):
                SMTP_SERVER_PORT = int(line.split("=")[1].strip())
        smtp_auth = {
            "SENDER_EMAIL": SENDER_EMAIL,
            "PASSWORD": PASSWORD,
            "SMTP_SERVER_ADDR": SMTP_SERVER_ADDR,
            "SMTP_SERVER_PORT": SMTP_SERVER_PORT
        }
        return smtp_auth
    except:
        smtp_auth = {
            "SENDER_EMAIL": None,
            "PASSWORD": None,
            "SMTP_SERVER_ADDR": None,
            "SMTP_SERVER_PORT": None
        }
        return smtp_auth


def get_recipients(config_dir):
    filename = "recipients.txt"
    file_path = os.path.join(config_dir, filename)
    if not os.path.exists(file_path):
        return []
    
    recipients = []
    for line in open(file_path, encoding="utf-8"):
        line = line.strip()
        if line.startswith("#") or line == "":
            continue
        recipients.append(line)
    return recipients


def main():
    config_dir = makesure_app_dir()
    smtp_auth = get_smtp_credentials(config_dir)
    recipients = get_recipients(config_dir)
    tool = JNAirport_Excel_Tool(smtp_auth=smtp_auth, recipients=recipients)
    tool.mainloop()