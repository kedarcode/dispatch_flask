from ssh_pymongo import MongoSession
from datetime import datetime
import xlrd
from json_excel_converter import Converter
from json_excel_converter.xlsx import Writer
import pandas
import json

session = MongoSession(
    '128.199.16.195',
    port=22,
    user='root',
    uri='mongodb://localhost:27017')

db = session.connection['notshy']

sales_sheet = pandas.read_excel(
    'C:\\Users\\KEDAR\\PycharmProjects\\Dispatchflask\\salesorder\\Invoice.xlsx',
    sheet_name="Invoice", ).to_json()
sales_sheet = json.loads(sales_sheet)

# myquery = { "AWBNo": "9166165313" }
# newvalues = { "$set": { "invoice_no": "123" } }
#
# db['dartdatas'].update_one(myquery, newvalues)
#
buffer = []
for n, s in enumerate(sales_sheet['PurchaseOrder']):
    if sales_sheet['PurchaseOrder'][s].isnumeric():
        if sales_sheet['PurchaseOrder'][s] not in buffer:
            print(n, sales_sheet['PurchaseOrder'][s], sales_sheet['Invoice Number'][s])
            myquery = {"AWBNo": sales_sheet['PurchaseOrder'][s]}
            newvalues = {"$set": {"invoice_no": sales_sheet['Invoice Number'][s]}}

            db['dartdatas'].update_one(myquery, newvalues)

            buffer.append(sales_sheet['PurchaseOrder'][s])
