import sys
import os
import pickle

from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QMessageBox
from PyQt5.QtGui import QIcon

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from calendar_style import Ui_Form as CalendarUI

from add_event import AddEvent

CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.pickle"
SCOPES = ["https://www.googleapis.com/auth/calendar"]

def get_google_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "wb") as token:
            pickle.dump(creds, token)
    return build("calendar", "v3", credentials=creds)

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath(".."), relative_path)

class CalendarApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_visuals()
        self.init_connections()
        self.google_service = get_google_service()

    def init_visuals(self):
        self.c = CalendarUI()
        self.c.setupUi(self)
        self.setWindowTitle("Calendar App")
        self.setFixedSize(1000, 800)

    def init_connections(self):
        self.c.addEventButton.clicked.connect(self.open_add_event_dialog)

    def open_add_event_dialog(self):
        dialog = AddEvent()
        dialog.set_google_service(self.google_service)
        dialog.exec()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = CalendarApp()
    window.show()
    app.exec()
