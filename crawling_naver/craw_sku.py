from openpyxl import load_workbook
from urllib.request import urlopen
import re

xlsx_wb = load_workbook("./상품정보요청건(12).xlsx")
xlsx_ws = xlsx_wb["Sheet1"]

cvs_url = "http://165.244.243.83:8008"
hnb_url = "http://gshbhq.gsretail.com:8008"

for row in xlsx_ws.rows:
  match = re.search("\D", row[1].value)
  if (match == None and row[6].value != None):
    with urlopen(cvs_url+"/FileDown.do?attFileId="+row[6].value+"&attFileOrderSeq=1") as f:
      with open("./goodsimg/"+row[5].value ,"wb") as h:
        img = f.read()
        h.write(img)

xlsx_wb.close()
