from reportprocessor import ReportProcessor
from reportdatamanager import ReportDataManager
from config import Config
import sys
import getopt

def main(argv):
    c = Config()
    webinarFile = ""
    registrationFile = ""
    attendeeFile = ""
    processedFiles = ""

    # try:
    #     opts, args = getopt.getopt(argv, "i:lh", ["input=", "list", "help"])
    # except getopt.GetoptError:
    #     print("main.py -i <inputfile name>")
    #     sys.exit(2)

    # for opt, arg in opts:
    #     if opt in ("-h", "--help"):
    #         print("main.py -i <inputfile>")
    #         sys.exit()
    #     elif opt in ("-l", "--list"):
    #         driveFiles = rdm.GetDriveFiles(c.AttendeeReportFolderID)
    #         for file in driveFiles:
    #             print(file["name"])
    #         sys.exit()
    #     elif opt in ("-i", "--input"):
    #         inputFile = arg

    rdm = ReportDataManager(c)
    rp = ReportProcessor(c)

    #If the files being processed are meeting files, as opposed to a webinar file, the webinarFile variable value should be a zero length string.
    webinarFile = "" #"20220913_attendee_report.csv"

    registrationFile = "Dec13-2022_RegistrationReport.csv"
    attendeeFile = "Dec13-2022_ParticipantsReport.csv"

    # NOTE: If the file exists in the input folder, on your local system, it will be removed so that a fresh copy will be downloaded.
    if len(webinarFile) >= 1:
        rdm.DownloadDriveFile(c.AttendeeReportFolderID, webinarFile)
        processedFiles = rp.ProcessWebinarFile(webinarFile)
    elif len(registrationFile) >= 1 & len(attendeeFile) >= 1:
        rdm.DownloadDriveFile(c.AttendeeReportFolderID, registrationFile)
        rdm.DownloadDriveFile(c.AttendeeReportFolderID, attendeeFile)
        processedFiles = rp.ProcessMeetingFile(registrationFile, attendeeFile)
    else:
        print("The input parameters appear to not be set correctly. The process will not continue.")
        sys.exit()
    
    rdm.WriteCPEFileToDrive(processedFiles["cpe"])
    rdm.WriteMetricsToSheet(processedFiles["metrics"])

if __name__ == "__main__":
    main(sys.argv[1:])