#!/usr/bin/env python3

import csv
import sys
import os
import datetime
import time
from decimal import Decimal

basedir = os.path.dirname(__file__)
if basedir:
    os.chdir(basedir)

payments_filename = 'payments.csv'
invoices_filename = 'invoices.csv'

invoices = list(csv.reader(open(invoices_filename,'r')))
assert invoices[0][3] == 'BookingReference'
assert invoices[0][4] == 'Reference'

booking_to_invoice = {}
for invoice in invoices[1:]:
    if not len(invoice):
        continue
    if invoice[3].strip().lower() in booking_to_invoice:
        print("!!!! Booking %r has multiple invoices, at least %r and %r" % (
            invoice[3].strip().lower(),
            booking_to_invoice[invoice[3].strip().lower()],
            invoice[4]))
        #raise Exception("Multiple invoices for one booking")
        # Assume last invoice is the correct one for now :/
    booking_to_invoice[invoice[3].strip().lower()] = invoice[4]

cash_payments = []
wire_transfers = []

payments = [payment[16:] for payment in csv.reader(open(payments_filename,'r'))]
assert payments[0][0] == 'ReceivedDate'
assert payments[0][1] == 'Forename'
assert payments[0][5] == 'BookingReference'
assert payments[0][9] == 'PaymentMethod'
assert payments[0][15] == 'Direct1'

for idx,payment in enumerate(payments[1:]):
    if not len(payment):
        continue

    if payment[0] != '(see above)':
        date = datetime.datetime.strptime(payment[0].split(' ')[0], '%d/%m/%Y')

    name = payment[1].strip()
    reservation = payment[5]
    if reservation.strip().lower() in booking_to_invoice:
        reference = booking_to_invoice[reservation.strip().lower()]
    else:
        reference = reservation

    payment_method = payment[9]

    payment_value_raw = payment[15]
    if not payment_value_raw:
        continue

    if payment_value_raw.startswith('- '):
        payment_is_income = False
        payment_value_raw = payment_value_raw[2:]
    else:
        payment_is_income = True

    try:
        assert payment_value_raw.startswith('€ ')
    except AssertionError:
        print("Error for payment %d %r" % (idx+1,payment))
        print("payment_value_raw: %r" % payment_value_raw)
        raise
    payment_value_raw = payment_value_raw[2:].replace(',','')
    payment_value = Decimal(payment_value_raw)

    payment_info = {
            'date': date,
            'name': name,
            'ref': reference,
            'amount': payment_value if payment_is_income else -payment_value,
    }

    if payment_method in ['Banküberweisung','Überweisung']:
        wire_transfers.append(payment_info)
    elif payment_method == 'Bargeld':
        cash_payments.append(payment_info)
    else:
        if payment_method not in [
                'Kartenlesegerät',
                'OTA Vorauszahlung',
            ]:
            raise RuntimeError("Unhandled Payment method %r" % payment_method)

if __name__ == '__main__':
    import xlwt

    date_style = xlwt.easyxf(num_format_str='DD.MM.YYYY')
    cur_style = xlwt.easyxf(num_format_str='€ #.00')

    wb = xlwt.Workbook()
    for name,source in [
            ('Barzahlungen', cash_payments),
            ('Ueberweisungen', wire_transfers)
            ]:
        ws = wb.add_sheet(name)
        ws.col(2).width = 5000
        ws.col(3).width = 8000
        ws.col(4).width = 500
        ws.col(6).width = 500
        ws.col(7).width = 500
        ws.col(8).width = 500
        ws.col(9).width = 500
        for idx,payment in enumerate(source):
            ws.write(idx,0,payment['date'],style=date_style)
            # internal ref stays empty
            ws.write(idx,2,payment['ref'])
            ws.write(idx,3,payment['name'])
            if payment['amount'] < 0:
                amount = -payment['amount']
                dest_col = 5
            else:
                amount = payment['amount']
                dest_col = 10
            ws.write(idx,dest_col,amount, style=cur_style)

    wb.save(time.strftime('billing-%Y-%m-%d-%H-%M.xls'))

