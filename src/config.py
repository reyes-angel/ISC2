from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from yaml.loader import SafeLoader
import os
import yaml
import datetime as dt
import pandas as pd
import io
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

class Config:
    def __init__(self):
        with open("../../config.yaml") as f:
            config = yaml.load(f, Loader=SafeLoader)
        
        self.MetricsSheetID = config["metrics_sheet_id"]
        self.CPEFolderID = config["cpe_folder_id"]
        self.AttendeeReportFolderID = config["attendee_report_folder_id"]
        self.OutputPath = config["output_path"]
        self.InputPath = config["input_path"]
        self.MeetingStartHour = config["meeting_start_hour"]
        self.MeetingEndHour = config["meeting_end_hour"]
        self.CPETemplatePath = config["cpe_template_path"]
        self.ChapterName = config["chapter_name"]
        self.MeetingInfoSheetID = config["meeting_info_sheet_id"]
        self.Officers = []

        for officer in config["chapter-officers"].keys():
            o = Officer(
                config["chapter-officers"][officer]["email"],
                config["chapter-officers"][officer]["user_name"],
                config["chapter-officers"][officer]["member_first_name"],
                config["chapter-officers"][officer]["member_last_name"],
                config["chapter-officers"][officer]["member_certification_id"],
                config["chapter-officers"][officer]["personal_email"],
                config["chapter-officers"][officer]["phone_number"],
                officer
            )

            self.Officers.append(o)


class Officer:
    def __init__(self, chapter_email, user_name, first_name, last_name, certification_id, personal_email, phone_number, role):

        self.ChapterEmail = chapter_email
        self.UserName = user_name
        self.FirstName = first_name
        self.LastName = last_name
        self.CertificationID = certification_id
        self.PersonalEmail = personal_email
        self.PhoneNumber = phone_number
        self.Role = role

class Google:
    def __init__(self):
        #TODO: The paths to these files should likely be stored in the config file
        token_file = "token_doc.json"
        credential_file = "credentials.json"
        self.creds = None

        token_scope = [
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/documents',
            'https://www.googleapis.com/auth/spreadsheets',
            'https://mail.google.com/'
        ]

        # The file token file stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(token_file):
            self.creds = Credentials.from_authorized_user_file(token_file, token_scope)

        # If there is no (valid) token file available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(credential_file, token_scope)
                self.creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open(token_file, 'w') as token:
                token.write(self.creds.to_json())

    def drive_service(self):
        return build('drive', 'v3', credentials=self.creds)
    
    def docs_service(self):
        return build('docs', 'v1', credentials=self.creds)
            
    def sheets_service(self):
        return build('sheets', 'v4', credentials=self.creds)
    
    def mail_service(self):
        return build('gmail', 'v1', credentials=self.creds)
    
    def download_file_csv(self, file_id, file_path_name, mime_type):
        try:
            request = self.drive_service().files().export_media(fileId=file_id, mimeType=mime_type)
            file = io.BytesIO()
            downloader = MediaIoBaseDownload(file, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
        except HttpError as error:
            print(F'An error occurred: {error}')
            file = None

        with open(file_path_name, "wb") as f:
            f.write(file.getbuffer())

class Meeting:
    def __init__(self):
        self.config = Config()
        g = Google()

        self.sheets_service = g.sheets_service()

    def GetInformation(self, meeting_date):
        meeting_topic = ""
        cpe_group = ""
        cpe_domain = ""
        meeting_month_date = ""
        meet_date = ""
        meeting_found = False

        data =  self.sheets_service.spreadsheets().values().get(spreadsheetId=self.config.MeetingInfoSheetID, range='Sheet1').execute().get('values')[1:]

        if isinstance(meeting_date, pd.Timestamp):
            meet_date = meeting_date.strftime("%m/%d/%Y")
            meeting_month_date = meeting_date.strftime("%b %Y")

        for i, row in enumerate(data):
            if row[0] == meet_date:
                meeting_topic = "{0} (ISC)² Silicon Valley Chapter Meeting - {1} - {2}".format(meeting_month_date, row[1], row[4])
                cpe_group = row[2]
                cpe_domain = row[3]
                meeting_found = True
                break

        if not meeting_found:
            raise Exception("Meeting Info not found in the spreadsheet")

        return meeting_topic, cpe_group, cpe_domain