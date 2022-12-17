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
from config import Google
from config import Config

# This script creates a pivor table from the data in the membership metrics file. 
# The pivot table shows the date of each meeting and a value indicating if the member attended.
# Some potential uses of this are to gather information to decide who to target for a board role or potentially
# to reward members who have attended the previous X meetings.

# NOTE: The pivot is based on name, email, and cert #.  The user input that we receive isn't great and so you may 
# find that there are multiple rows for the same user; each row representing the distinct combination of values we have in the metrics file for each user.
# We could potentially enhance the process to clean up the metrics file, but that hasn't been yet.

class MeetingsAttended:

    def __init__(self, configuration):
        self.config = configuration
        self.g = Google()
    
    def CreatePivotTable(self, inputFile, outputFile):    
        self.g.download_file_csv(self.config.MetricsSheetID, inputFile, "text/csv")

        df = pd.read_csv(inputFile, header=0, dtype={"(ISC)2 Member #": str})
        df["Date of Activity"] = pd.to_datetime(df["Date of Activity"])

        df["Member First Name"] = df["Member First Name"].str.strip().str.lower()
        df["Member Last Name"] = df["Member Last Name"].str.strip().str.lower()
        df["Name"] = df["Member First Name"] + " " + df["Member Last Name"]
        df["Email"] = df["Email"].str.strip().str.lower()
        df["cert #"] = df["(ISC)2 Member #"].astype(str)

        df.sort_values(by=["Date of Activity", "Member First Name", "Member Last Name"], ascending=[False, True, False], inplace=True, ignore_index=True)

        df2 = df[(df["Member First Name"] != "unknown") & (df["Date of Activity"] > "2021-09-01")]

        pivot = pd.pivot_table(
            data=df2,
            index=['Name', "Email" ,'cert #'],
            columns="Date of Activity",
            values="CPE Qualifying Minutes",
            fill_value="-1" 
        )

        # -  -1 values indicate that the member did not register or attend the meeting.
        # - 0.0 indicates that the member registered but did not attend the meeting.
        # - Any value greater than 0.0 indicates that the member attended the meeting.
        pivot.to_csv(outputFile, index=True)
        print(pivot)

c = Config()

input_file = c.InputPath + "MembershipMetrics.csv"
output_file = c.OutputPath + "MembershipMetrics_MeetingsAttended.csv"

ma = MeetingsAttended(c)
ma.CreatePivotTable(input_file, output_file)