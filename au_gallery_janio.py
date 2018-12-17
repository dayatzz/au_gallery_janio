import xlrd
import datetime
import csv
import sys


try:
    filename = sys.argv[1]
except Exception:
    print(
        """
  $ python au_galler_janio.py filename.xlsx
""")
    exit()

wb = xlrd.open_workbook(filename)

ws = wb.sheet_by_index(0)
ws2 = wb.sheet_by_index(1)


def xldate_to_datetime(xldate):
    temp = datetime.datetime(1900, 1, 1)
    delta = datetime.timedelta(days=xldate - 2)
    return temp + delta


def convert_to_dict(sheet, convert_column={}):
    fields = []
    for column in range(sheet.ncols):
        # ws.cell_value(row, column)
        field = sheet.cell_value(0, column)
        fields.append(field)

    data = []
    for row in range(1, sheet.nrows):
        obj = {}
        for column in range(sheet.ncols):
            r = sheet.cell_value(row, column)
            if fields[column] in convert_column:
                r = convert_column[fields[column]](r)
            obj[fields[column]] = r
        data.append(obj)
    return data


sheet1_data = convert_to_dict(ws, {
    'BookTime': xldate_to_datetime
})
sheet2_data = convert_to_dict(ws2, {
    'BookTime': xldate_to_datetime,
    'PayTime': xldate_to_datetime
})
postals = {int(d['OrderId']): d['Zip Code'] for d in sheet2_data}


def find_duplicate(obj, data):
    filtered = filter(
        lambda x: x['consignee_name'].lower() == obj['Customer Name'].lower()
        and x['consignee_number'] == obj['Phone']
        and x['consignee_address'].lower() == obj['Address'].lower()
        and x['item_desc'].lower() == obj['SKU Name'].lower(),
        data)
    try:
        return list(filtered)
    except:
        return []


fields = [
    'shipper_order_id',
    'tracking_no',
    'item_desc',
    'item_quantity',
    'item_product_id',
    'item_sku',
    'item_category',
    'item_price_value',
    'item_price_currency',
    'consignee_name',
    'consignee_number',
    'consignee_address',
    'consignee_postal',
    'consignee_country',
    'consignee_state',
    'consignee_city',
    'consignee_province',
    'consignee_email',
    'order_length',
    'order_width',
    'order_height',
    'order_weight',
    'payment_type',
    'cod_amt_to_collect'
]
csv_data = []
for d in sheet1_data:
    f = find_duplicate(d, csv_data)
    if f:
        index = csv_data.index(f[0])
        csv_data[index]['shipper_order_id'] += ', {}'.format(
            str(int(d['OrderId'])))
        csv_data[index]['item_quantity'] += int(d.get('Quantity'))
    else:
        obj = {
            'shipper_order_id': str(int(d.get('OrderId', 0))),
            'tracking_no': '',
            'item_desc': d.get('SKU Name', ''),
            'item_quantity': int(d.get('Quantity', 0)),
            'item_product_id': '',
            'item_sku': int(d.get('SKU Number', 0)),
            'item_category': 'Lifestyle Products',
            'item_price_value': d.get('PaySubtotal', ''),
            'item_price_currency': 'IDR',
            'consignee_name': d.get('Customer Name', ''),
            'consignee_number': d.get('Phone', ''),
            'consignee_address': d.get('Address', ''),
            'consignee_postal': postals.get(d.get('OrderId', ''), ''),
            'consignee_country': 'Indonesia',
            'consignee_state': d.get('District', ''),
            'consignee_city': d.get('City', ''),
            'consignee_province': d.get('Province', ''),
            'consignee_email': '',
            'order_length': 50,
            'order_width': 50,
            'order_height': 2,
            'order_weight': 1,
            'payment_type': 'prepaid',
            'cod_amt_to_collect': '',
        }
        csv_data.append(obj)


with open('janio_excel_csv.csv', 'w') as f:
    w = csv.DictWriter(f, fieldnames=fields)
    w.writeheader()
    w.writerows(csv_data)
