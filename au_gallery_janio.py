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


def find_duplicate_consignee(obj, data, item=False):
    filtered = filter(
        lambda x: x['consignee_name'].lower() == obj['Customer Name'].lower()
        and x['consignee_number'] == obj['Phone']
        and x['consignee_address'].lower() == obj['Address'].lower(),
        data)
    try:
        return list(filtered)
    except:
        return []


def find_duplicate_item(obj, data):
    filtered = filter(
        lambda x: x['item_desc'].lower() == obj['SKU Name'].lower(), data
    )
    try:
        return list(filtered)
    except:
        return []


def convert_to_janio_object(raw_obj, ids=False):
    if not ids:
        ids = str(int(raw_obj.get('OrderId', 0)))

    obj = {
        'shipper_order_id': ids,
        'tracking_no': '',
        'item_desc': raw_obj.get('SKU Name', ''),
        'item_quantity': int(raw_obj.get('Quantity', 0)),
        'item_product_id': '',
        'item_sku': int(raw_obj.get('SKU Number', 0)),
        'item_category': 'Lifestyle Products',
        'item_price_value': raw_obj.get('PaySubtotal', ''),
        'item_price_currency': 'IDR',
        'consignee_name': raw_obj.get('Customer Name', ''),
        'consignee_number': raw_obj.get('Phone', ''),
        'consignee_address': raw_obj.get('Address', ''),
        'consignee_postal': postals.get(raw_obj.get('OrderId', ''), ''),
        'consignee_country': 'Indonesia',
        'consignee_state': raw_obj.get('District', ''),
        'consignee_city': raw_obj.get('City', ''),
        'consignee_province': raw_obj.get('Province', ''),
        'consignee_email': '',
        'order_length': 50,
        'order_width': 50,
        'order_height': 2,
        'order_weight': 1,
        'payment_type': 'prepaid',
        'cod_amt_to_collect': '',
    }
    return obj


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
    duplicate_consignee = find_duplicate_consignee(d, csv_data)
    if duplicate_consignee:
        id = str(int(d.get('OrderId', 0)))
        ids = set([i['shipper_order_id'] for i in duplicate_consignee])
        ids = ', '.join(ids)
        if id not in ids:
            ids += ', {}'.format(id)
            for i in duplicate_consignee:
                index = csv_data.index(i)
                csv_data[index]['shipper_order_id'] = ids

        obj = convert_to_janio_object(d, ids=ids)

        duplicate_item = find_duplicate_item(d, duplicate_consignee)
        if duplicate_item:
            index = csv_data.index(duplicate_item[0])
            csv_data[index]['item_quantity'] += int(d.get('Quantity', 0))
        else:
            csv_data.append(obj)
    else:
        obj = convert_to_janio_object(d)
        csv_data.append(obj)


with open('janio_excel_csv.csv', 'w') as f:
    w = csv.DictWriter(f, fieldnames=fields)
    w.writeheader()
    w.writerows(csv_data)
