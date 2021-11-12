from gspread_formatting import *
import gspread
import os
import csv
import sqlite3
conn = sqlite3.connect('C:\\BioRad\\BioRad.db')
cur = conn.cursor()
from ftplib import FTP





gc = gspread.service_account()

sh = gc.open("Tutorial")
worksheet = sh.sheet1
sheetName = "Sheet1"


print(sh.sheet1.get('A1'))
print(sh.sheet1.get('B1'))
sh.sheet1.update('B1', 'Bingo!')
print(sh.sheet1.get('B1'))

sh.sheet1.update_acell('F7','=SUM(F1:F6)')

usps1 = "https://tools.usps.com/go/TrackConfirmAction?tRef=fullpage&tLc=2&text28777=&tLabels="
track = '92001901755477000640516409'
usps2 = "%2C&tABt=false"
usps = usps1 + track + usps2
print (usps)
worksheet.format('A1:B1', {'textFormat': {'bold': True},
                           "horizontalAlignment": 'CENTER'
                           })

fmt = cellFormat(
    backgroundColor=color(1, 0.9, 0.9),
    textFormat=textFormat(bold=True, foregroundColor=color(1, 0, 1), fontFamily='Roboto', fontSize=18),
    horizontalAlignment='CENTER'
    )

format_cell_range(worksheet, 'A1', fmt)

#sh.update_cells(cell_list, value_input_option='USER_ENTERED')
#cell_values = [1,'=HYPERLINK("' + some_url + '","' + some_text + '")',3]
sh.sheet1.update_acell('F8','=HYPERLINK("'+usps+'")')
assert isinstance(sh, object)

sheetId = sh.worksheet(sheetName)._properties['sheetId']
print(sheetId)
worksheet_list = sh.worksheets()
print(worksheet_list)
body = {
    "requests": [
        {
            "mergeCells": {
                "mergeType": "MERGE_ALL",
                "range": {  # In this sample script, all cells of "A1:C3" of "Sheet1" are merged.
                    "sheetId": 0,
                    "startRowIndex": 0,
                    "endRowIndex": 3,
                    "startColumnIndex": 0,
                    "endColumnIndex": 3
                }
            }
        }
    ]
}
#res =
sh.batch_update(body)
listFile = []
listSplit = []
dict1 = {}
with open('C:\\BioRad\\Run25.txt', newline='') as a1cresults:
    results = csv.reader(a1cresults, delimiter='\t')
    count = 0

    for result in results:

        #print(result)

        if count > 3:
            if result[3] == "P":
                if result[7] == "A1c":
                    print(result[3] + " " + result[5] + " " + result[9])
                    a1c = result[9]
                    a1c = a1c.replace("*", "")
                    listSplit.append(result[3] + "," + result[5] + "," + a1c)

                    #cur.execute("INSERT INTO MachineResults (SampleID, Date, A1c) VALUES(?, ?, ?)",(result[4], result[5], a1c))
                    #conn.commit()

            listFile.append(results)
            print(type(result))
            print(type(results))
        count = count + 1

    #cur.execute("INSERT INTO MachineResults (SampleID, Date, A1c) VALUES ('1','2','3')")
    #conn.commit()



    cur.close()
    conn.close()
    print(len(result))
    print(len(listFile))
    print(len(listSplit))
    print(listSplit[0])

   # print(listFile(0))


#ftp = FTP('ftp.mya1cresults.com')
#"anonymous", "ftplib-example-1"
#ftp.login("Onduo@mya1cresults.com","ArThUrW91!")

ftp = FTP('ftp.benchmarknetworks.com')
#"anonymous", "ftplib-example-1"
ftp.login("biorad@benchmarknetworks.com","B10R@d21!")

data = []

ftp.dir(data.append)

ftp.quit()

for line in data:
    print("-", line)







