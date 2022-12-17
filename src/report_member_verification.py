from config import Config
from config import Google
from reportdatamanager import ReportDataManager
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd

# The script will iterate through the downloaded file, for each record it will call out, using selenium, to the ISC2 verification page 
# with the last name and membership number.  If there was no matching ISC2 record we will write that to the output file, otherwise we'll write 
# the members certification information to the file.  Some members have multiple certifications, each under the same membership number.
# Information for each certification will be written to the output file.
# Note: Records should be filtered to only include the previous 12 months, be distinct, and include an ISC2 Member Number.

# NOTE: I've made some changes to the code here, but they aren't able to be tested because the chrome drive doesn't appear to be available 
# for my current version of chrome.  I could potentially downgrade chrome, but that's not worth the effort.
# The code did work before, so I don't have any reason to believe that it won't now...but it's best to be aware and to test.

class MemberVerification:

    def __init__(self, configuration):
        self.config = configuration
        self.g = Google()
    
    def VerifyMembers(self, inputFile, outputFile):
        self.g.download_file_csv(self.config.MetricsSheetID, inputFile, "text/csv")

        options = Options()
        options.headless = False
        options.add_argument("--window-size=500,500")

        df = pd.read_csv(inputFile, header=0, dtype={"(ISC)2 Member #": str})

        # Consider filtering this to include only records within the past X months.
        df2 = df.drop_duplicates(subset=["Member Last Name", "(ISC)2 Member #"], keep="last")

        with open(outputFile, "w") as f:
            for i, row in df2[(~pd.isnull(df["(ISC)2 Member #"]))].iterrows():
                first_name = row["Member First Name"]
                last_name = row["Member Last Name"]
                member_number = row["(ISC)2 Member #"]

                driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
                url = "https://www.isc2.org/MemberVerification?LastName=" + last_name + "&MemberNumber=" + member_number
                driver.get(url)
                time.sleep(5) # Ensures that the data binding javascript is executed.  It does sometimes fail.
                
                blnFound = len(driver.find_elements(By.CSS_SELECTOR, value="div[class='card dashboard mvnoresults'][style='display:none']")) == 1

                if blnFound:
                    certTypes = driver.find_elements(By.CSS_SELECTOR, value="p[data-bind='text: longName'")
                    certDates = driver.find_elements(By.CSS_SELECTOR, value="p[data-bind='text: certifiedDate'")
                    certExpirations = driver.find_elements(By.CSS_SELECTOR, value="p[data-bind='text: expirationDate'")

                    for idx, val in enumerate(certTypes):
                        output = "'%s','%s','%s','%s','%s','%s'\n" % (first_name, last_name, member_number, certTypes[idx].text, certDates[idx].text, certExpirations[idx].text)
                        f.write(output)
                        f.flush()
                        print(output)
                else:
                    output = "'%s','%s','%s','Not Found'\n" % (first_name, last_name, member_number)
                    f.write(output)
                    f.flush()
                    print(output)

                driver.quit()
       
    def FindLastAttended(self, inputFile, outputFile):
        inputFile = "/Users/Angel/Downloads/MembershipMetrics.csv"
        df1 = pd.read_csv(inputFile, header=0, dtype={"(ISC)2 Member #": str})
        df1["Date of Activity"] = pd.to_datetime(df1["Date of Activity"])

        df2 = pd.read_csv(outputFile, header=0, dtype={"(ISC)2 Member #": str})
        df2["Last Attended"] = np.nan

        for i, row in df2.iterrows():
            first_name = row["Member First Name"]
            last_name = row["Member Last Name"]
            member_number = row["(ISC)2 Member #"]

            df3 = df1[(df1["Member First Name"] == first_name) & (df1["Member Last Name"] == last_name) & (df1["(ISC)2 Member #"] == member_number)]
            
            df2.loc[i, "Last Attended"] = pd.to_datetime(df3["Date of Activity"].max()).strftime("%m/%d/%Y")

        df2.to_csv(outputFile, index=False)


c = Config()

input_file = c.InputPath + "MembershipMetrics.csv"
output_file = c.OutputPath + "MembershipMetrics_MemberVerification.csv"

mv = MemberVerification(c)
mv.VerifyMembers(input_file, output_file)
mv.FindLastAttended(input_file, output_file)



