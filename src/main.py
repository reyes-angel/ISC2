import io
import os
from os import listdir
from os.path import isfile, join
import re
from numpy import NaN, datetime64, float64
import numpy as np
import pandas as pd
import datetime
from datetime import datetime as dt
from datetime import timedelta as td
import yaml
from yaml.loader import SafeLoader

# Attendee Use Cases:
# 1) Attendee joined and left before the chapter meeting begun - Discard
# 2) Attendee joined and left after the chapter meeting ended - Discard
# 3) Attendee joined the meeting multiple times
#   a] The attendance interval of one record is completely within the attendance interval of another record - Discard
#   b] The Join Time of one record starts within 60 seconds of another ending. This might happen 
#      when an attendee gets updated to a panelist - Extend/bridge the two records
#   c] The Join Time of one record starts more than 60 seconds after another ended.  This might happen
#      when an attendee disconnects from the call for some reason.  Nothing to do here, we'll sum the attendance records 
#      but there's no need to extend or bridge records. 
# 4) Attendee joined before the meeting started or within five minutes (grace period) after the start of the meeting.
#      Extend join time to beginning of the meeting
# 5) Attendee left after the meeting started or within five minutes (grace period) before the end of the meeting.
#      Extend leave time to the end of the meeting

# TODO: Potential future feature - Update records with no certification number from previous meeting files based on email address.
# TODO: Reporting the average minutes attended is going to be skewed by the fact that we are allowing 
# for grace periods and bridging, etc... I need to keep that separate from the actual join and leave times

input_path = ""
output_path = ""
grace_period_minutes = 0
meeting_start_hour = 0
meeting_end_hour = 0
chapter_officers = {}

def ReadConfig():
    global grace_period_minutes
    global meeting_start_hour
    global meeting_end_hour
    global chapter_officers
    global input_path
    global output_path

    with open("./src/config.yaml") as f:
        config = yaml.load(f, Loader=SafeLoader)
    
    input_path = config["input_path"]
    output_path = config["output_path"]
    meeting_start_hour = config["meeting_start_hour"]
    meeting_end_hour = config["meeting_end_hour"]
    grace_period_minutes = config["grace_period_minutes"]
    chapter_officers = config["chapter-officers"]
    

def ProcessFiles():
    ReadConfig()
    
    fileList = [f for f in listdir(input_path) if isfile(join(input_path, f)) and f.endswith(".csv")]
    
    for f in fileList:
        in_file_name = input_path + f

        file_name, file_extension  = os.path.splitext(f)
        metrics_file_name = output_path + file_name + "_metrics.csv"
        cpe_file_name = output_path + file_name + "_cpe.csv"

        io_dict = ReadReportFile(in_file_name)
        df, meeting_date, meeting_name = CreateDataFrame(io_dict)
        ProcessDataFrame(df, meeting_date.date(), meeting_name)
        
        df.drop(["Join Time", "Leave Time", "Minutes Attended", "Registration Time", "Approval Status"], axis=1, inplace=True)
        df.drop_duplicates(keep="first", subset=["Email"], inplace=True)
        df.sort_values(by=["Email"], inplace=True, ignore_index=True)

        # Export the Metrics File
        df[["Attended", "Region", "Type", "CPE Qualifying Minutes", "Date of Activity"]].to_csv(metrics_file_name, index=False)
        
        # Export the CPE File
        df.loc[(df["# CPEs"] > 0) & (df["(ISC)2 Member #"]).notna(), ["(ISC)2 Member #", "Member First Name", "Member Last Name", "Title of Meeting", "# CPEs", "Date of Activity", "CPE Qualifying Minutes"]].to_csv(cpe_file_name, index=False)

def ReadReportFile(fileName):
    table_names = ["Host Details", "Panelist Details", "Attendee Details", "Other Attended", "Report Generated"]

    io_host = io.StringIO()
    io_panelist = io.StringIO()
    io_attendee = io.StringIO()
    io_other = io.StringIO()
    io_report = io.StringIO()
    attendee_type = "None"

    with open(fileName, "r") as f_in:
        
        for line in f_in.readlines():
            if line.startswith(table_names[0]):
                attendee_type = "Host"
                continue
            elif  line.startswith(table_names[1]):
                attendee_type = "Panelist"
                continue
            elif line.startswith(table_names[2]):
                attendee_type = "Attendee"
                continue
            elif line.startswith(table_names[3]):
                attendee_type = "Other"
                continue
            elif line.startswith(table_names[4]):
                attendee_type = "Report"
                continue

            line = line.replace("--", "").strip().rstrip(",") + "\n"
            if attendee_type == "Host":
                io_host.write(line)
            elif attendee_type == "Attendee":
                io_attendee.write(line)
            elif attendee_type == "Panelist":
                io_panelist.write(line)
            elif attendee_type == "Other":
                io_other.write(line)
            elif attendee_type == "Report":
                io_report.write(line)

    io_report.seek(0)
    io_host.seek(0)
    io_attendee.seek(0)
    io_panelist.seek(0)
    io_other.seek(0)

    io_dict ={
        "Report": io_report,
        "Host": io_host,
        "Attendee": io_attendee,
        "Panelist": io_panelist,
        "Other": io_other
    }

    return io_dict

# Returns a tuple (data frame, meeting date, meeting name)
def CreateDataFrame(ioDictionary):
    df_report = pd.read_csv(ioDictionary["Report"], header=0)
    meeting_name = df_report.iloc[0]["Topic"]

    df_host = pd.read_csv(ioDictionary["Host"], header=0)
    df_host["Type"] = "Host"
    # We weren't able to reliably get the meeting date from the Report table because the "Actual Start Time" value is sometimes the day before the meeting!
    # We'll get it from the Host table.
    meeting_date = pd.to_datetime(df_host.iloc[0]["Join Time"])

    df_attendee = pd.read_csv(ioDictionary["Attendee"], header=0)
    df_attendee["Type"] = "Attendee"

    df_panelist = pd.read_csv(ioDictionary["Panelist"], header=0)
    df_panelist["Type"] = "Panelist"
    
    # It is common for there to not be any attendee's listed in the "Other" table
    if len(ioDictionary["Other"].getvalue()) > 0:
        df_other = pd.read_csv(ioDictionary["Other"], header=0)
        df_other.rename(columns={"User Name": "User Name (Original Name)"},  inplace=True)
        df_other["Type"] = "Other"
        df_other["Email"] = df_other["User Name (Original Name)"].map(str)
        df_other["First Name"] = "Unknown"
        df_other["Last Name"] = "Unknown"
        df_other["Attended"] = "Yes"
    else:
        df_other = pd.DataFrame()

    df = pd.concat([df_host, df_panelist, df_attendee, df_other], ignore_index=True)
    return df, meeting_date, meeting_name

def ProcessDataFrame(df, meeting_date, meeting_name):
    meeting_start = np.datetime64(dt.combine(meeting_date, datetime.time(meeting_start_hour, 00)))
    meeting_end = np.datetime64(dt.combine(meeting_date, datetime.time(meeting_end_hour, 00)))

    df.rename(columns={"First Name": "Member First Name", "Last Name": "Member Last Name", "User Name (Original Name)": "User Name", "Time in Session (minutes)": "Minutes Attended", "Country/Region Name": "Region", "(ISC)2 Certification:": "(ISC)2 Member #"},  inplace=True)

    # Chapter officers will likely register to attend the meeting under their personal email address and provide thier Member #.  However, when they join the meeting
    # the panelist/host email invite may have them joining under thier chapter email addres and not associate a Member #.  In this case, they would not get CPE credit for having
    # joined the call.  To avoid this, we update the chapter officers record with thier first, last, and member #.
    for officer in chapter_officers.keys():
        df.loc[df["Email"]== chapter_officers[officer]["email"], ["User Name"]] = chapter_officers[officer]["email"]
        df.loc[df["Email"]== chapter_officers[officer]["email"], ["Member First Name"]] = chapter_officers[officer]["member_first_name"]
        df.loc[df["Email"]== chapter_officers[officer]["email"], ["Member Last Name"]] = chapter_officers[officer]["member_last_name"]
        df.loc[df["Email"]== chapter_officers[officer]["email"], ["(ISC)2 Member #"]] = chapter_officers[officer]["member_num"]

    # Some attendees fail to add only their certification ID when registering for the meeting.
    # Here we extract the numeric digits from the "(ISC)2 Member #" filed and replace the existing field value.
    # Most certification nunbers are 6 digit but this is looking for a number between 4 and 9 digits long.
    df["(ISC)2 Member #"] = df["(ISC)2 Member #"].astype(str).str.extract(r"([0-9]{4,9})")
    df["Join Time"] = pd.to_datetime(df["Join Time"])
    df["Leave Time"] = pd.to_datetime(df["Leave Time"])

    # Add new fields to the data frame
    df["# CPEs"] = 0.0
    df["CPE Qualifying Minutes"] = 0.0
    df["Date of Activity"] = dt.strftime(meeting_date, "%m/%d/%Y")
    df["Title of Meeting"] = meeting_name
    
    # When someone has connected to the meeting more than once there are multiple entries in the report.  
    # Only the first will list an "(ISC)2 Member #" number, assuming the attendee provided one when they registered.
    # Here we backfill all of those empty values.  
    df["(ISC)2 Member #"] = df.groupby(["Email"])["(ISC)2 Member #"].transform("first")

    # Covers Attendee Use Case 1 and 2
    # Discard any record where the attendee either left before the meeting started or joined after it ended
    df.drop(df.loc[(df["Join Time"] >= meeting_end) | (df["Leave Time"] <= meeting_start)].index, axis=0, inplace=True)

    # Covers Attendee Use Case 3A
    DiscardInvtervals(df)

    # Covers Attendee Use Case 3B
    BridgeIntevals(df)

    # Covers Attendee Use Case 4
    # If the attendee joined at any point before the meeting or within 5 minutes of the meeting start we will
    # consider the join time to be at the start of the meeting (typically 6:00 PM)
    df["Join Time"] = df.apply(lambda x: meeting_start if  x["Join Time"] < meeting_start or abs((meeting_start  - x["Join Time"]).total_seconds()/60.0) <= grace_period_minutes else x["Join Time"], axis=1)

    # Covers Attendee Use Case 5
    # if the attendee left the meeting after the meeting ended or within 5 minutes of the meeting end we will
    # consider the leave time to the at the end of the meeting (typically 8:00 PM)
    df["Leave Time"] = df.apply(lambda x: meeting_end if  x["Leave Time"] >= meeting_end or abs((x["Leave Time"] - meeting_end).total_seconds()/60.0) <= grace_period_minutes else x["Leave Time"], axis=1)

    # Calculate how many minutes of the meeting were attended, per record (not yet aggregated)
    df["Minutes Attended"] = (df["Leave Time"] - df["Join Time"]).astype("timedelta64[m]")

    #Calculate CPE Qualifying Minutes and # CPEs
    df["CPE Qualifying Minutes"] = df.groupby(["Email"])["Minutes Attended"].transform("sum")

    # Round CPE Qualifying Minutes to the nearest 15 minutes and then divide by 60 to find the number of CPE's to assign for attendance
    # Each 15 minutes of attendance gives the attendee .25 CPE's.  An attendee that accumulated less than 15 minutes of attendance will receive 0 CPE's 
    df["# CPEs"] = ((df["CPE Qualifying Minutes"]/15).round().astype(int) * 15)/60


def DiscardInvtervals(df):
    df.sort_values(by=["Email", "Join Time", "Leave Time"], ascending=[True, True, False], inplace=True, ignore_index=True)
    index_list_to_delete = []

    for email_address in df[df.duplicated(subset=["Email"], keep="first")]["Email"].unique():
        filter_by_email = df["Email"].eq(email_address)
        attendee_intervals = df[filter_by_email].filter(items=["Join Time", "Leave Time"], axis=1).to_dict()
        
        max_index = max(attendee_intervals["Join Time"].keys())
        min_index = min(attendee_intervals["Join Time"].keys())

        index = max_index

        while index > min_index:
            current_start = attendee_intervals["Join Time"][index]
            current_end = attendee_intervals["Leave Time"][index]
            
            compare_index = max_index if index == min_index else index - 1
            previous_start = attendee_intervals["Join Time"][compare_index]
            previous_end = attendee_intervals["Leave Time"][compare_index]

            if(current_start >= previous_start and current_end <= previous_end):
                index_list_to_delete.append(index)
            
            index -= 1

    df.drop(index_list_to_delete, axis=0, inplace=True)


def BridgeIntevals(df):
    df.sort_values(by=["Email", "Join Time", "Leave Time"], ascending=[True, True, False], inplace=True, ignore_index=True)
    index_list_to_delete = []

    for email_address in df[df.duplicated(subset=["Email"], keep="first")]["Email"].unique():
        filter_by_email = df["Email"].eq(email_address)
        attendee_intervals = df[filter_by_email].filter(items=["Join Time", "Leave Time"], axis=1).to_dict()
        
        max_index = max(attendee_intervals["Join Time"].keys())
        min_index = min(attendee_intervals["Join Time"].keys())

        index = max_index

        while index > min_index:
            current_start = attendee_intervals["Join Time"][index]
            current_end = attendee_intervals["Leave Time"][index]
            
            compare_index = max_index if index == min_index else index - 1
            previous_end = attendee_intervals["Leave Time"][compare_index]
            
            if((current_start - td(minutes=1)) <= previous_end):
                # Here we're updating both the dataframe and the attendee intervals dictionary
                # The source of our comparisons is the dictionary, if we only update the dataframe 
                # we won't have the most up to date information
                df.loc[compare_index, ["Leave Time"]] = current_end
                attendee_intervals["Leave Time"][compare_index] = current_end
                index_list_to_delete.append(index)
            
            index -= 1

    df.drop(index_list_to_delete, axis=0, inplace=True)

ProcessFiles()
