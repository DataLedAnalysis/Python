from urllib.request import HTTPBasicAuthHandler
import pandas as pd, requests

baseurl = "https://api.keyedinprojects.co.uk/V3/api/"
endpoint = "report"

usr = "abdul.wasay@thameswater.co.uk"
pwd = "FQ03tbvc!"
report = "771"

#endpoint is report, query parameter is key
#we still need to provide page numbers and loop through them

def req(baseurl,endpoint,report,pg):
    r = requests.get(baseurl + endpoint + f"?key={report}&pageNumber={pg}", auth=(usr,pwd))
    return r.json() #parses to a json file

def NoOfPages(response):
    return response['TotalPages']

def parseData(response):
    ls = []
    for item in response['Data']:
        ls.append(item)
    return ls

#total pages for 771 key are 12
d = req(baseurl,endpoint,report,1)

mainlist = []
for pg in range(1,NoOfPages(d)+1,1):
    mainlist.extend(parseData(req(baseurl,endpoint,report,pg)))

df = pd.DataFrame(mainlist)
df.to_excel('test.xlsx',index=False)

