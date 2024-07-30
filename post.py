import requests, xlrd
import pandas as pd
import openpyxl



# read file
dataframe1 = pd.read_excel('Corsi di formazione 2024.xlsx')
print(dataframe1)

# iterate 


# congrats






url = "https://europa.eu/europass/eportfolio/api/eprofile/profiles/66a6493fa89ba072bef3badb/educations-training"

data = {
    "qualification": "Anna was here",
    "organisationName": "asd",
    "mediaAttachments": [],
    "links": [],
    "order": 3
}

headers = {
  "Content-Type": "application/json; charset=utf-8", 
  "Accept": "application/json, text/plain",
  "Accept-Language" : "it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7",
  "Cache-Control": "No-Cache",
  "Connection" : "keep-alive",
  "Content-Type" : "application/json",
  "Cookie" : "dtCookie=v_4_srv_38_sn_D4D75FFABD22061B1DCD74C4C44E8934_perc_100000_ol_0_mul_1_app-3A5808f124b494edbd_1_rcs-3Acss_0; XSRF-TOKEN=9f0504ee-3f04-496f-91b7-30c5da9e03fc; EUROPASS_API_SESSION_ID=f7032816-2c9b-4fcc-a1aa-0de4220f8552; EUROPASS_AUTH_SESSION_ID=OWY4ZjNiOTAtNjYxYy00ZjNkLWI3MzctNTBmZDcwMzY3MDk0; dtCookie=v_4_srv_38_sn_64352D6BB0752CC68F6CC3BDDE6C7997; _pk_id.1bd3d421-8073-4872-99de-c7c0e2242372.9e93=11a7dd421db762a6.1722173409.1.1722200682.1722173409.; rxVisitor=1722173407618K2LCPOGO5T163R6J4CNE7HCPL5C3TE3B; dtCookie=v_4_srv_38_sn_D4D75FFABD22061B1DCD74C4C44E8934_perc_100000_ol_0_mul_1_app-3A5808f124b494edbd_1_rcs-3Acss_0; cck1=^%^7B^%^22cm^%^22^%^3Atrue^%^2C^%^22all1st^%^22^%^3Atrue^%^2C^%^22closed^%^22^%^3Afalse^%^7D; dtSa=-; rxvt=1722202507297^|1722199577982; dtPC=38^$200679804_52h27vKMOKDMMPUEFRKERFPPUUUULQVOHRCQAM-0e0^",
  "Origin" : "https://europa.eu",
  "Referer" : "https://europa.eu/europass/eportfolio/screen/profile?profileId=66a6493fa89ba072bef3badb&actionId=c8857fb6-7dce-4eee-a189-d5cf3eac6bb0&lang=en",
  "Sec-Fetch-Dest" : "empty",
  "Sec-Fetch-Mode" : "cors",
  "Sec-Fetch-Site" : "same-origin",
  "User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
  "X-Requested-With" : "XMLHttpRequest",
  "X-XSRF-TOKEN" : "9f0504ee-3f04-496f-91b7-30c5da9e03fc",
  "x-dtpc" : "38^$200679804_52h27vKMOKDMMPUEFRKERFPPUUUULQVOHRCQAM-0e0^"
  }

# response = requests.post(url, headers=headers, json=data)

# print(response.text)
# print(response.status_code)