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
import requests

input_file = "/Users/Angel/Downloads/MembershipMetrics.csv"
df = pd.read_csv(input_file, header=0, dtype={"(ISC)2 Member #": str})
df["Member First Name"] = df["Member First Name"].str.strip().str.lower()
df["Member Last Name"] = df["Member Last Name"].str.strip().str.lower()
df["Email"] = df["Email"].str.strip().str.lower()

df["group1"] = df["Member First Name"] + " " + df["Member Last Name"]
df["group2"] = df["Email"]
df["group3"] = df["group1"] + "; " + df["Email"]
df["cert #"] = df["(ISC)2 Member #"].astype(str)

# # Check for duplicate certification numbers based on First and Last Name
# s = df.groupby("group1")["(ISC)2 Member #"].transform("nunique").rename("Unique Counts")
# print(df[s > 1].groupby(["group1", s])["group1", "cert #"].agg(lambda x: "; ".join(set(x))))

# # Check for duplicate certification numbers based on Email Address
# s = df.groupby("group2")["(ISC)2 Member #"].transform("nunique").rename("Unique Counts")
# print(df[s > 1].groupby(["group2", s])["group2", "cert #"].agg(lambda x: "; ".join(set(x))))

# # Check for duplicate certification numbers based on First Name, Last Name, and Email Address
# s = df.groupby("group3")["(ISC)2 Member #"].transform("nunique").rename("Unique Counts")
# print(df[s > 1].groupby(["group2", s])["group3", "cert #"].agg(lambda x: "; ".join(set(x))))

# # Check for duplicate First and Last Name based on ceftificatio number
# s = df.groupby("(ISC)2 Member #")["group1"].transform("nunique").rename("Unique Counts")
# print(df[s > 1].groupby(["(ISC)2 Member #", s])["group1"].agg(lambda x: "; ".join(set(x))))

# Check for duplicate Email based on ceftificatio number
#s = df.groupby("(ISC)2 Member #")["group2"].transform("nunique").rename("Unique Counts")
#print(df[s > 1].groupby(["(ISC)2 Member #", s])["group2"].agg(lambda x: "; ".join(set(x))))