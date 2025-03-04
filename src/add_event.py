import sys
import os
import pickle
from PyQt5.QtWidgets import QMessageBox, QDialog
from PyQt5.QtCore import QDateTime, QTime, Qt

from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from add_event_style import Ui_Dialog as AddEventUI

SCOPES = ['https://www.googleapis.com/auth/calendar']

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath(".."), relative_path)

def authenticate_google_api():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('calendar', 'v3', credentials=creds)

class AddEvent(QDialog):
    def __init__(self):
        super().__init__()
        self.google_service = None
        self.init_visuals()
        self.ui.confirmButton.clicked.connect(self.add_event)
        self.google_service = authenticate_google_api()

    def init_visuals(self):
        self.ui = AddEventUI()
        self.ui.setupUi(self)
        self.setWindowTitle("Add Event")
        self.setFixedSize(400, 600)

    def check_values(self):
        currDateTime = QDateTime.currentDateTime()
        userDateTime = self.ui.dateTime.dateTime()
        userDur = self.ui.duration.time()

        if not self.ui.name.toPlainText().strip():
            self.error = "No name provided"
        elif userDateTime < currDateTime:
            self.error = "Cannot add event in the past"
        elif userDur < QTime(0, 1):
            self.error = "Duration too short"
        else:
            return True

        return False

    def set_google_service(self, service):
        self.google_service = service

    def sync_with_google(self, event):
        if not self.google_service:
            self.error = "Google service not initialized"
            return False
        try:
            self.google_service.events().insert(calendarId='primary', body=event).execute()
            return True
        except Exception as e:
            self.error = f"Google API Error: {e}"
            print(f"Google API Error: {e}")
            return False

    def add_event(self):
        if self.check_values():
            startDateTime = self.ui.dateTime.dateTime()
            durationTime = self.ui.duration.time()
            totalSeconds = durationTime.hour() * 3600 + durationTime.minute() * 60
            endDateTime = startDateTime.addSecs(totalSeconds)

            event = {
                'summary': self.ui.name.toPlainText(),
                'start': {
                    'dateTime': startDateTime.toString(Qt.ISODate),
                    'timeZone': 'UTC'
                },
                'end': {
                    'dateTime': endDateTime.toString(Qt.ISODate),
                    'timeZone': 'UTC'
                }
            }
            if self.google_service and self.sync_with_google(event):
                self.accept()
                QMessageBox.information(self, "Info", "Event added!")
            else:
                QMessageBox.warning(self, "Error", f"{self.error}")
        else:
            QMessageBox.warning(self, "Error", f"{self.error}")
            print(f"{self.error}")
            return
