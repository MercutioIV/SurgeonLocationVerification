import requests
import json
import pandas as pd
import array as arr
import openpyxl
import time
from openpyxl.styles import PatternFill

start = time.time()

workbook = pd.read_excel('C:\\Users\\dwerelius\\Desktop\\US-Vmedi-nonContacts.xlsx')
xfile = openpyxl.load_workbook('C:\\Users\\dwerelius\\Desktop\\US-Vmedi-nonContacts.xlsx')
URL = "https://npiregistry.cms.hhs.gov/api/?version=2.0"
sheet = xfile['Sheet1']
xfile['Sheet1']
rows = sheet.max_row - 1
unregCount = 0
movedSurgeons = 0

def getInformation():
    global unregCount
    global movedSurgeons
    for x in range(rows):
        try:
            rowNum = x+2
            npi = workbook['NPI'].iloc[x]
            queryURL = URL + f"&number={npi}"
            response = requests.get(queryURL)

            userdata = json.loads(response.text)

            firstName = (userdata["results"][0])["basic"]["first_name"]
            lastName = (userdata["results"][0])["basic"]["last_name"]
            city = (userdata["results"][0])["addresses"][0]["city"]
            state = (userdata["results"][0])["addresses"][0]["state"]

            xcity = (workbook['Location/City'].iloc[x])
            xcity = xcity.upper()
            xstate = (workbook['State'].iloc[x])
            xstate = xstate.upper()

            if (xcity != city):
                sheet[f'Q{rowNum}'] = f'{city}'
                sheet[f'R{rowNum}'] = f'{state}'
                sheet[f'D{rowNum}'].fill = PatternFill(patternType='solid',
                                                       fgColor='FC2C03')
                sheet[f'E{rowNum}'].fill = PatternFill(patternType='solid',
                                                       fgColor='FC2C03')
                print(f"\n{firstName} {lastName} has moved to {city}, {state}")
                movedSurgeons += 1
            else:
                print(f"\n{firstName} {lastName} lives in {city}, {state}")

        except KeyError:
            sheet[f'D{rowNum}'].fill = PatternFill(patternType='solid',
                                                   fgColor='000000')
            sheet[f'E{rowNum}'].fill = PatternFill(patternType='solid',
                                                   fgColor='000000')
            print(f"\nNPI# {npi} is no longer registered")
            unregCount+=1


getInformation()
xfile.save('C:\\Users\\dwerelius\\Desktop\\US-Vmedi-nonContacts.xlsx')
print("\n#############################################")
print("# Inactive Surgeons: " + str(unregCount))
print("# New Surgeon Locations: " + str(movedSurgeons))
end = time.time()
print("# Finished in " + str(end - start) + " seconds")
print("#############################################")