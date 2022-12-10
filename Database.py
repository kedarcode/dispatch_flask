from ssh_pymongo import MongoSession
from datetime import datetime
import xlrd
from json_excel_converter import Converter
from json_excel_converter.xlsx import Writer
from Path import PathResource


final_sheet = []


def validate(data):
    ks = data.keys()
    if 'salesOrderId' in ks and 'invoice_no' in ks and ('mobile_no1' in ks or 'mobile_no2' in ks)\
            and 'vendor_name' in ks and 'state_name' in ks and 'products' in ks and len(data['products']) > 0:
        return True
    else:
        return False


proser = []
products = ['TGEL', 'TABB-69(60)', 'PENIKOSETE', 'MR.BIG(60)', 'SOFTSPOTT', 'TG', 'PS', 'TB', 'MB', 'TABB', 'DAHAADOIL',
            'HORNYGOAT', 'DAHAAD', 'TABB-69F(30)', 'TRIPLEGINSENG']

prod_12 = ['PENIKOSETE', 'TABB-69(60)', 'TABB-69F(30)']


def caltax(data, sc, get_total):
    at = 0
    tax12 = 0
    tax18 = 0
    tax = 0
    obj = {}
    lab = ['cgst12', 'cgst18', 'sgst12', 'sgst18']
    lab2 = ['igst12', 'igst18']
    for n in data:
        at += n['price']
    if at == 0:
        return
    percentage_factor = get_total / at
    print(get_total)
    for d in data:
        name_str = d["name"].replace(' ', '').upper()
        d['qty'] = int(d['qty'])
        if name_str in prod_12:

            tax12 += float(format(((d["price"] * d['qty']) * percentage_factor) / d["qty"], ".2f"))
        elif name_str in products:
            tax18 += float(format(((d["price"] * d['qty']) * percentage_factor) / d["qty"], ".2f"))
        else:
            return

    if sc == 'MH':
        print(sc)
        for l in lab:
            if '12' in l:
                obj[l] = format(
                    (float(tax12) - (float(tax12) / 1.12)) / 2, '.2f')
                tax += (float(tax12) - (float(tax12) / 1.12)) / 2
            if '18' in l:
                obj[l] = format(
                    (float(tax18) - (float(tax18) / 1.18)) / 2, '.2f')
                tax += (float(tax18) - (float(tax18) / 1.18)) / 2

    else:
        for l in lab2:
            if '12' in l:
                obj[l] = format(
                    (float(tax12) - (float(tax12) / 1.12)), '.2f')
                tax += float(tax12) - (float(tax12) / 1.12)

            if '18' in l:
                obj[l] = format(
                    (float(tax18) - (float(tax18) / 1.18)), '.2f')
                tax += float(tax18) - (float(tax18) / 1.18)

    print(tax)
    obj['sub_total'] = format(get_total - tax, '.2f')

    print(obj, get_total, tax12, tax18)
    return obj


print(PathResource.resource_path(f'output/dispatch{str(datetime.today()).split(".")[0].replace(":", "")}.xlsx'))


def generate_sheet(start, end):
    session = MongoSession(
        '128.199.16.195',
        port=22,
        user='root',
        uri='mongodb://localhost:27017')

    db = session.connection['notshy']
    getdata = db['dartdatas'].find({"shipdate": {"$gte": start, "$lt": end}},
                                   {"salesOrderId": 1, "AWBNo": 1, "name": 1, "invoice_no": 1,
                                    "shipdate": 1, "add1": 1, "add2": 1, "add3": 1, "add4": 1,
                                    "consignee_pin": 1, "mobile_no1": 1, "state_code": 1,
                                    "mobile_no2": 1, "state_name": 1, "vendor_name": 1, "weight": 1, "pices": 1,
                                    "cod": 1, "upi": 1,
                                    "amount": 1, "products": 1, "agent_name": 1})

    for data in getdata:
        if str(data['AWBNo']).isnumeric() and validate(data):
            temp_pro = list(products)
            obj = {}
            obj['SALESORDER_ID'] = data['salesOrderId']
            obj['Sales Order'] = data['AWBNo']
            obj['invoice_no'] = data['invoice_no']
            if type(data["shipdate"]) != int and type(data["shipdate"]) != str:
                obj['invoice_date'] = f'{data["shipdate"].day}/{data["shipdate"].month}/{data["shipdate"].year}'
                obj['Date'] = f'{data["shipdate"].day}/{data["shipdate"].month}/{data["shipdate"].year}'
            obj['Customer Name'] = data["name"]
            obj['address'] = data['add1'] + data['add2'] + data['add3'] + data['add4']
            obj['pin_code'] = data['consignee_pin']
            try:
                obj['Phone No (customer)'] = data['mobile_no1']
            except:
                try:
                    obj['Phone No (customer)'] = data['mobile_no2']
                except:
                    pass
            obj['state'] = data['state_name']
            obj['vendor_name'] = data['vendor_name']
            obj['weight'] = int(data['weight']) / 1000
            obj['total_quantity'] = data['pices']
            obj['cod_amount'] = data['cod']
            obj['upi'] = data['upi']
            obj['Amount'] = data['amount']
            obj['state_code'] = data['state_code']
            tax = caltax(data['products'], data["state_code"], obj['Amount'])
            if tax:
                obj.update(tax)
            for pro in data['products']:
                name_str = pro['name'].replace(' ', '').upper()
                if name_str in temp_pro:
                    temp_pro.remove(name_str)
                    obj[name_str] = pro["qty"]
            for rem in temp_pro:
                obj[rem] = 0
            obj['agent'] = data["agent_name"]
            print(obj)
            final_sheet.append(obj)
    conv = Converter()
    conv.convert(final_sheet, Writer(file=PathResource.resource_path(
        f'output/dispatch{str(datetime.today()).split(".")[0].replace(":", "")}.xlsx')))

    session.stop()
    return f'output/dispatch{str(datetime.today()).split(".")[0].replace(":", "")}.xlsx'


generate_sheet(datetime(2022, 11, 24), datetime(2022, 12, 7))
