from reportprocessor import ReportProcessor
from reportdatamanager import ReportDataManager
from config import Config
import sys
import getopt


def main(argv):

    c = Config()
    rdm = ReportDataManager(c)
    rp = ReportProcessor(c)

    # try:
    #     opts, args = getopt.getopt(argv, "i:lh", ["input=", "list", "help"])
    # except getopt.GetoptError:
    #     print("main.py -i <inputfile>")
    #     sys.exit(2)

    # for opt, arg in opts:
    #     if opt in ("-h", "--help"):
    #         print("main.py -i <inputfile>")
    #         sys.exit()
    #     elif opt in ("-l", "--list"):
    #         driveFiles = rdm.GetDriveFiles(c.AttendeeReportFolderID)
    #         for file in driveFiles:
    #             print(file["name"])
    #     elif opt in ("-i", "--input"):
    #         inputfile = arg
    
    # TODO: Remove this hard coded value
    inputFile = "20220712_attendee_report.csv"
    

    # TODO: Add a check to be sure it doesn't already exist
    # Should the check be in the report data manager? Or should it be here?
    #fileList = [f for f in listdir(input_path) if isfile(join(input_path, f)) and f.endswith(".csv")]
    # If the file exists, ask if they want to delete it, continue with the existing file, or exit the program

    #rdm.DownloadDriveFile(c.AttendeeReportFolderID, inputFile)
    
    # TODO: Check that the file exists locally after download, it should also be a CSV. If the name is wrong it won't download and won't throw an error.

    processedFiles = rp.ProcessFiles(inputFile)
    rdm.WriteCPEFileToDrive(processedFiles["cpe"])
    rdm.WriteMetricsToSheet(processedFiles["metrics"])
    # TODO: Rename the input file on Google Drive, thier names are not indicative of the meeting date

if __name__ == "__main__":
    main(sys.argv[1:])
