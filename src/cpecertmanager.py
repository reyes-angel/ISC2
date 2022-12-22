from config import Config
from config import Google
from config import Meeting
from datetime import datetime as dt
from googleapiclient.errors import HttpError
import pandas as pd
from emailmessage import messaging

class CPECertificateManager:

    def __init__(self, configuration):
        self.config = configuration
        self.g = Google(configuration)
        self.drive_service = self.g.drive_service()
        self.docs_service = self.g.docs_service()
        self.sheets_service = self.g.sheets_service()

        self.m = messaging()

    def _get_data(self, sheet_tab_name):
        try:
            return self.sheets_service.spreadsheets().values().get(spreadsheetId=self.config.CPECertificateSheetID, range=sheet_tab_name).execute().get('values')[1:]
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    # Copy the template file into the CPE Certificates folder with a file name specific to the member.  One file will be copied per member who earned CPE's
    def _copy_template(self, user_cpe_file_name):
        try:
            body = {'name': user_cpe_file_name, 'parents': [self.config.CPECertificateFolderID]}
            return self.drive_service.files().copy(body=body, fileId=self.config.CPECertificateTemplateID, fields='id').execute().get('id')
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    def _merge_template(self, user_cpe_file_name, merge_data):
        try:
            user_cpe_file_id = self._copy_template(user_cpe_file_name)
            context = merge_data.iteritems() if hasattr({}, 'iteritems') else merge_data.items()

            reqs = [{'replaceAllText': {
                'containsText': {
                    'text': '{{%s}}' % key.upper(),
                    'matchCase': True,
                },
                'replaceText': value,
            }} for key, value in context]

            self.docs_service.documents().batchUpdate(body={'requests': reqs}, documentId=user_cpe_file_id, fields='').execute()
            return user_cpe_file_id
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    def SendCPECertificates(self, meeting_date):

        merge_data = Meeting().GetInformation(meeting_date)
        merge_data["date_signed"] = dt.today().strftime("%m/%d/%Y")

        columns = ['date_of_activity', 'member_name', 'email_address', 'cpe_count', 'member_number']
        sheet_tab_name = meeting_date.strftime("%Y-%m-%d")
        data = self._get_data(sheet_tab_name)

        for i, row in enumerate(data):
            merge_data.update(dict(zip(columns, row)))
            user_cpe_file_name = '%s_CPE_%s' % (dt.strptime(merge_data["Meeting_Date"], "%m/%d/%Y").strftime("%Y %b"), merge_data["member_name"])
            user_cpe_file_name = user_cpe_file_name.upper().replace(" ", "_")
            user_cpe_file_id = self._merge_template(user_cpe_file_name, merge_data)

            user_cpe_pdf_file_name = "%s%s.pdf" % (self.config.OutputPath, user_cpe_file_name)
            
            self.g.download_file(user_cpe_file_id, user_cpe_pdf_file_name, "application/pdf")
            self.m.send_message_and_attachments(merge_data["member_name"], merge_data["email_address"], user_cpe_pdf_file_name, meeting_date)

c = Config()

m = CPECertificateManager(c)
meeting_date = pd.to_datetime("12/13/2022")
m.SendCPECertificates(meeting_date)

