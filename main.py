import requests
import json
import pandas as pd
import array as arr
import openpyxl

workbook = pd.read_excel('C:\\Users\\dwerelius\\Desktop\\US-Vmedi-nonContacts.xlsx')
ps = openpyxl.load_workbook('C:\\Users\\dwerelius\\Desktop\\US-Vmedi-nonContacts.xlsx')
sheet = ps['Sheet1']
npiList = arr.array('L')

rows = sheet.max_row
print(rows)
for i in range(500):
    npiList.append(workbook['NPI'].iloc[i])

URL = "https://npiregistry.cms.hhs.gov/api/?version=2.0"


def getInformation(npi):
    for x in npi:
        queryURL = URL + f"&number={x}"
        response = requests.get(queryURL)

        userdata = json.loads(response.text)

        firstName = (userdata["results"][0])["basic"]["first_name"]
        lastName = (userdata["results"][0])["basic"]["last_name"]
        city = (userdata["results"][0])["addresses"][0]["city"]
        state = (userdata["results"][0])["addresses"][0]["state"]

        print(f"\n{firstName} {lastName} lives in {city}, {state}")


getInformation(npiList)
