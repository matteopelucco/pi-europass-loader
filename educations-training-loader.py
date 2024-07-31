import requests, xlrd
import pandas as pd
import openpyxl
import time
from datetime import datetime


# Read Excel file
file_path = 'Corsi di formazione 2024.xlsx'
df = pd.read_excel(file_path, header=None, skiprows=2)

# Setup Europass API params
profile_id = "66a6493fa89ba072bef3badb"
url = "https://europa.eu/europass/eportfolio/api/eprofile/profiles/" + profile_id + "/educations-training"
dtpc = "33$376069889_815h33vRBUKGPEARPUFRRPKUABBDQGKRAKQWSIG-0e0"
xsrf_token = "3e73a999-87c8-43c1-86f7-f6d276bf0973"
cookie = "dtCookie=v_4_srv_33_sn_6C52DBC52E34875E3A87496B9B405FD9_perc_100000_ol_0_mul_1_app-3A5808f124b494edbd_1_rcs-3Acss_0; XSRF-TOKEN=3e73a999-87c8-43c1-86f7-f6d276bf0973; EUROPASS_API_SESSION_ID=c80f7a50-8dad-4396-b1ba-60bcc66cb5af; EUROPASS_AUTH_SESSION_ID=OTY0M2E5ZTAtZGQxOC00ZDNmLWE2MTUtYWEzZDE4YTc1YWZk; dtCookie=v_4_srv_33_sn_0678FC912AF5A36C048DF28FA42B00DA; _pk_ses.1bd3d421-8073-4872-99de-c7c0e2242372.9e93=*; _pk_id.1bd3d421-8073-4872-99de-c7c0e2242372.9e93=11a7dd421db762a6.1722173409.2.1722376075.1722376067.; cck1=^%^7B^%^22cm^%^22^%^3Atrue^%^2C^%^22all1st^%^22^%^3Atrue^%^2C^%^22closed^%^22^%^3Afalse^%^7D; rxVisitor=1722376061958D72L17L9GQKVA0B4E5SCE2JC0IDL53GE; dtCookie=v_4_srv_33_sn_6C52DBC52E34875E3A87496B9B405FD9_perc_100000_ol_0_mul_1_app-3A5808f124b494edbd_1_rcs-3Acss_0; dtSa=-; rxvt=1722377994535^|1722376061959; dtPC=33^$376069889_815h33vRBUKGPEARPUFRRPKUABBDQGKRAKQWSIG-0e0^"

headers = {
  "Content-Type": "application/json; charset=utf-8", 
  "Accept": "application/json, text/plain",
  "Accept-Language" : "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7",
  "Cache-Control": "No-Cache",
  "Connection" : "keep-alive",
  "Content-Type" : "application/json",
  "Cookie" : cookie,
  "Origin" : "https://europa.eu",
  "Referer" : "https://europa.eu/europass/eportfolio/screen/profile?lang=it",
  "Sec-Fetch-Dest" : "empty",
  "Sec-Fetch-Mode" : "cors",
  "Sec-Fetch-Site" : "same-origin",
  "User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
  "X-Requested-With" : "XMLHttpRequest",
  "X-XSRF-TOKEN" : xsrf_token,
  "x-dtpc" : dtpc
  }


# Date conversion function
def convert_date_1(date):
    if isinstance(date, str):
        date = datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
    elif isinstance(date, pd.Timestamp):
        date = date.to_pydatetime()
    return date.strftime('%Y-%d-%m')

def convert_date_2(date):
    if isinstance(date, str):
        date = datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
    elif isinstance(date, pd.Timestamp):
        date = date.to_pydatetime()
    return date.strftime('%Y-%m-%dT%H:%M:%S.000Z')


# Now
current_timestamp = datetime.now().isoformat()

# Data set
data = []

# Iteration
for index, row in df.iterrows():
    
    print("date: " + str(row[0]))

    course_data = {
        "qualification": row[2],  # Colonna C
        "organisationName": row[5],  # Colonna F
        "creationDate": current_timestamp,
        "startDate": {
            "date": convert_date_2(row[0]),  # Colonna A
            "dateType": "DAY",
            "empty": False
        },
        "endDate": {
            "date": convert_date_2(row[1]),  # Colonna B
            "dateType": "DAY",
            "empty": False
        },
        "ongoing": False
    }

    print(course_data)
    data.append(course_data)

    response = requests.post(url, headers=headers, json=course_data)

    print(response.text)
    print(response.status_code)
    print(response.headers)

    # just for test
    # break

    # just to avoid massive spam
    time.sleep(2)

# Verify read data
# import json
# print(json.dumps(data, indent=4))







