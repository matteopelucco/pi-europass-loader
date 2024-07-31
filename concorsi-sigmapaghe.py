import requests, xlrd
import pandas as pd
import openpyxl
import time
from datetime import datetime


# Read Excel file
file_path = 'Corsi di formazione 2024.xlsx'
df = pd.read_excel(file_path, header=None, skiprows=2)

# Setup SIGMAPaghe API params

url = "https://concorsi.sigmapaghe.com/WCONCR01F.pgm"
cookie = "isLDAP=0; Regio1=3; Azien1=58; NSC_MCWT_GBSN_OHJOY_IUUQT=ffffffffaf1c0a8945525d5f4f58455e445a4a42378b"


headers = {
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Cookie": cookie,
    "Origin": "https://concorsi.sigmapaghe.com",
    "Referer": "https://concorsi.sigmapaghe.com/WCONCR01F.pgm?task=compila&rrn_000=000001052&rrn_007=000008594&rrn_005=000081402&smurfid=00207518fd79adf5ec67ed59ef9fe9e46df89dc0c133da535bdce1f47ec6dc64",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "sec-ch-ua": '"Not)A;Brand";v="99", "Google Chrome";v="127", "Chromium";v="127"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"'
}

# Date conversion function
def convert_date(date):
    if isinstance(date, str):
        date = datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
    elif isinstance(date, pd.Timestamp):
        date = date.to_pydatetime()
    return date.strftime('%d/%m/%Y')

# constants
COURSE_TYPE_CORSO = "01"
COURSE_TYPE_FAD = "03"

# Iteration
for index, row in df.iterrows():
    
    date_row = str(row[1])
    title = str(row[2])

    course_type = COURSE_TYPE_CORSO
    if (title.startswith("FAD")): 
        course_type = COURSE_TYPE_FAD

    date_str = convert_date(date_row)
    date = datetime.strptime(date_str, "%d/%m/%Y")

    print("date: " + date_str)
    print("title: " + title)
    print("course_type: " + course_type)

    if datetime(2020, 1, 1) <= date <= datetime(2024, 12, 31):

        data = {
            "smurfid": "00207518fd79adf5ec67ed59ef9fe9e46df89dc0c133da535bdce1f47ec6dc64",
            "rrn_005": "81402",
            "rrn_000": "1052",
            "rrn_007": "8594",
            "rrncr1": "0",
            "task": "new_Tit2",
            "SLC001": course_type, 
            "SLC002": "002",
            "SLC003": "001",
            "DAT001": date_str
        }
        

        response = requests.post(url, headers=headers, data=data)
        print(response.status_code)
        # print(response.text)

        # just to avoid massive spam
        time.sleep(2)

    else:
        print("La data DAT001 non è compresa tra il 2019 e il 2024.")
        continue
