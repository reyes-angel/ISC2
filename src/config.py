from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from yaml.loader import SafeLoader
import os
import yaml
from datetime import datetime as dt
import pandas as pd
import io
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
import numpy

class Config:
    def __init__(self):
        with open("/Users/Angel/Documents/python_code/ISC2_config/config.yaml") as f:
            config = yaml.load(f, Loader=SafeLoader)
        
        self.CPECertificateFolderID = config["cpe_certificate_folder_id"]
        self.CPECertificateTemplateID = config["cpe_certifictate_template_id"]
        self.CPECertificateSheetID = config["cpe_certificate_sheet_id"]
        self.MetricsSheetID = config["metrics_sheet_id"]
        self.CPEFolderID = config["cpe_folder_id"]
        self.AttendeeReportFolderID = config["attendee_report_folder_id"]
        self.ConfigPath = config["config_path"]
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
    def __init__(self, config):
        token_file = config.ConfigPath + "token.json"
        credential_file = config.ConfigPath + "credentials.json"
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
    
    def download_file(self, file_id, file_path_name, mime_type):
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

    # Checks if a sheet name exists in a google spreadsheet
    def sheet_name_exists(self, spreadsheet_id, sheet_name):
        sheet_exists = False
        sheet_id = -1

        spreadsheet = self.sheets_service().spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        for sheet in spreadsheet["sheets"]:
            if sheet["properties"]["title"] == sheet_name:
                sheet_id = sheet["properties"]["sheetId"]
                sheet_exists = True
    
        output = {}
        output["Spreadsheet_ID"] = spreadsheet_id
        output["Sheet_Exists"] = sheet_exists
        output["Sheet_Name"] = sheet_name
        output["Sheet_ID"] = sheet_id
        return output
    
    # Adds a new sheet with the specified sheet name to a google spreadsheet
    def add_sheet(self, spreadsheet_id, sheet_name):
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name,
                        'tabColor': {
                            'red': 0.44,
                            'green': 0.99,
                            'blue': 0.50
                        }
                    }
                }
            }]
        }

        response = self.sheets_service().spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=request_body).execute()

        return response

class Meeting:
    def __init__(self):
        self.config = Config()
        g = Google(self.config)

        self.sheets_service = g.sheets_service()

    def GetInformation(self, meeting_date):
        output = {}
        meeting_found = False

        data =  self.sheets_service.spreadsheets().values().get(spreadsheetId=self.config.MeetingInfoSheetID, range='Sheet1').execute().get('values')[1:]

        if isinstance(meeting_date, numpy.datetime64):
            meeting_date = pd.to_datetime(meeting_date)

        if isinstance(meeting_date, pd.Timestamp):
            meet_date = meeting_date.strftime("%m/%d/%Y")
            meeting_month_date = meeting_date.strftime("%b %Y")
        else:
            raise Exception("Meeting date is not valid")

        for i, row in enumerate(data):
            if row[0] == meet_date:
                output["Meeting_Date"] = row[0]
                output["Meeting_Topic"] = row[1]
                output["CPE_Group"] = row[2]
                output["CPE_Domain"] = row[3]
                output["Speaker_Name"] = row[4]
                output["Meeting_Topic_Long"] = "{0} (ISC)² Silicon Valley Chapter Meeting - {1} - {2}".format(meeting_month_date, row[1], row[4])  
                meeting_found = True
                break

        if not meeting_found:
            raise Exception("Meeting Info not found in the spreadsheet")

        return output