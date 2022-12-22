from __future__ import print_function
from config import Config
from config import Google
from googleapiclient.http import MediaFileUpload
from googleapiclient.http import MediaIoBaseDownload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from os import listdir
from os.path import isfile, join
import csv
import io
import os
import os.path
import shutil
import pandas as pd
    
class ReportDataManager:

    def __init__(self, configuration):

        self.config = configuration
        self.g = Google(configuration)
        self.sheets_service = self.g.sheets_service()
        self.drive_service = self.g.drive_service()

    # Writes the contents of a CSV file on the local machine to a Google sheet
    def _writeCSVToSpreadsheet(self, file_name, spreadsheetID, sheet_name, include_header):
       with open(file_name, 'r') as csvfile:
            rows = []
            csvreader = csv.reader(csvfile)

            if include_header:
                start_line = 0
            else:
                start_line = 1
            
            for row in csvreader:
                if csvreader.line_num > start_line:
                    rows.append(row)

            body = {'values':  rows}
            range_name = "%s!A1" % (sheet_name)
            self.sheets_service.spreadsheets().values().append(spreadsheetId=spreadsheetID, range=range_name, valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS", body=body).execute()

    # After the attendee report(s) have been processed the results will be written to the google sheet where the metrics are stored.
    # The google sheet is the source of the data studio dashboard
    def WriteMetricsToSheet(self, file_name):
        self._writeCSVToSpreadsheet(file_name, self.config.MetricsSheetID, "Sheet1", False)

    # After the attendee report(s) have been processed results will be written to the google sheet where infor related to issuing
    # CPE certificates is stored
    def WriteCertificateInfoToSheet(self, file_name, meeting_date):
        sheet_name = meeting_date.strftime("%Y-%m-%d")
        sheet_exists_info = self.g.sheet_name_exists(self.config.CPECertificateSheetID, sheet_name)

        if not sheet_exists_info["Sheet_Exists"]:
            self.g.add_sheet(self.config.CPECertificateSheetID, sheet_name)
        else:
            raise Exception("A sheet with the specified name already exists in the Google Sheet Doc")

        self._writeCSVToSpreadsheet(file_name, self.config.CPECertificateSheetID, sheet_name, True)


    def WriteCPEFileToDrive(self, file_name):
        f = os.path.basename(file_name)
        file_metadata = {
            "name" : f,
            "parents" : [self.config.CPEFolderID]
        }
        
        media_content = MediaFileUpload(file_name, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.drive_service.files().create(body=file_metadata, media_body=media_content).execute()

    # Retrieve the list of files from a specific folder
    def GetDriveFiles(self, folder_id, file_name=None):
        
        if file_name is not None:
            query = f"parents = '{folder_id}' and name = '{file_name}'"
        else:
            query = f"parents = '{folder_id}'"

        results = self.drive_service.files().list(q=query).execute()
        items = results.get('files', [])

        # It is possible that there are multiple files with the same name
        return items

    # Download files from a specific folder.   
    # If no file name is provided all files in the folder will be downloaded
    def DownloadDriveFile(self, folder_id, file_name=None):

        for item in self.GetDriveFiles(folder_id, file_name):
            file_id = item['id']
            file_name = item['name']

            file_name = self.config.InputPath + file_name

            if os.path.exists(file_name):
                os.remove(file_name)

            request = self.drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False

            while done is False:
                status, done = downloader.next_chunk()
            
            fh.seek(0)

            with open(file_name, 'wb') as f:
                shutil.copyfileobj(fh, f, length=131072)