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
from config import Config
from config import Google

# Shows the distinct number of attendees from the date of the meeting to present.  
# For example, If we needed to know the count of distinct attendees across all meetings since a particular month we would run this script
# to find that count.  The numbers will change monthly.

class DistinctAttendees:

    def __init__(self, configuration):
        self.config = configuration
        self.g = Google(configuration)
    
    def CreateReport(self, inputFile):
        self.g.download_file(self.config.MetricsSheetID, inputFile, "text/csv")

        df = pd.read_csv(inputFile, header=0, dtype={"(ISC)2 Member #": str})
        df["Date of Activity"] = pd.to_datetime(df["Date of Activity"])

        df["Member First Name"] = df["Member First Name"].str.strip().str.lower()
        df["Member Last Name"] = df["Member Last Name"].str.strip().str.lower()
        df["Email"] = df["Email"].str.strip().str.lower()
        df["cert #"] = df["(ISC)2 Member #"].astype(str)

        df.sort_values
        df.sort_values(by=["Date of Activity", "Member First Name", "Member Last Name"], ascending=[False, True, False], inplace=True, ignore_index=True)

        meetings = pd.unique(df["Date of Activity"])

        for m in meetings:
            df2 = df[(df["Attended"]=="Yes") & (df["Member First Name"] != "unknown") & (df["Date of Activity"] >= m)].drop_duplicates(["Member First Name", "Member Last Name"])
            print("%s\t-->\t%s" % (m.astype('datetime64[s]').item().strftime('%Y-%m-%d'), len(df2)))




c = Config()

input_file = c.InputPath + "MembershipMetrics.csv"

da = DistinctAttendees(c)
da.CreateReport(input_file)

#Test
#Test2