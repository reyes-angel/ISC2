from __future__ import print_function
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
    
class ReportDataManager:

    def __init__(self, configuration):

        self.config = configuration
        g = Google()
        self.sheets_service = g.sheets_service()
        self.drive_service = g.drive_service()

    # After the attendee report(s) have been processed the results will be written to the google sheet where the metrics are stored.
    # The google sheet is the source of the data studio dashboard
    def WriteMetricsToSheet(self, fileName):

        with open(fileName, 'r') as csvfile:
            rows = []
            csvreader = csv.reader(csvfile)
            
            for row in csvreader:
                if csvreader.line_num > 1:
                    rows.append(row)

            body = {'values':  rows}
            self.sheets_service.spreadsheets().values().append(spreadsheetId=self.config.MetricsSheetID, range="Sheet1!A1", valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS", body=body).execute()

    def WriteCPEFileToDrive(self, fileName):
        f = os.path.basename(fileName)
        file_metadata = {
            "name" : f,
            "parents" : [self.config.CPEFolderID]
        }
        
        media_content = MediaFileUpload(fileName, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
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

    # Print the file list from a specific folder
    def PrintDriveFileList(self, folder_id):
        for item in self.GetDriveFiles(folder_id):
            print(u'{0} ({1})({2})'.format(item['name'], item['id'], item['mimeType']))

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