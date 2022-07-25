import yaml
from yaml.loader import SafeLoader

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
        self.Officers = []

        for officer in config["chapter-officers"].keys():
            o = Officer(
                config["chapter-officers"][officer]["email"],
                config["chapter-officers"][officer]["user_name"],
                config["chapter-officers"][officer]["member_first_name"],
                config["chapter-officers"][officer]["member_last_name"],
                config["chapter-officers"][officer]["member_certification_id"],
                config["chapter-officers"][officer]["personal_email"]
            )

            self.Officers.append(o)


class Officer:
    def __init__(self, chapter_email, user_name, first_name, last_name, certification_id, personal_email):

        self.ChapterEmail = chapter_email
        self.UserName = user_name
        self.FirstName = first_name
        self.LastName = last_name
        self.CertificationID = certification_id
        self.PersonalEmail = personal_email