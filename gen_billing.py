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

BOOKING_REFERENCE_COLUMN = 3
REFERENCE_COLUMN = 4
if (invoices[0][BOOKING_REFERENCE_COLUMN] != 'BookingReference'
        or invoices[0][REFERENCE_COLUMN] != 'Reference'):
    raise Exception("Invoice report changed format!")

booking_to_invoice = {}
for invoice in invoices[1:]:
    if not len(invoice):
        continue
    if invoice[BOOKING_REFERENCE_COLUMN].strip().lower() in booking_to_invoice:
        print("!!!! Booking %r has multiple invoices, at least %r and %r" % (
            invoice[BOOKING_REFERENCE_COLUMN].strip().lower(),
            booking_to_invoice[invoice[BOOKING_REFERENCE_COLUMN].strip().lower()],
            invoice[REFERENCE_COLUMN]))
        #raise Exception("Multiple invoices for one booking")
        # Assume last invoice is the correct one for now :/
    booking_to_invoice[invoice[BOOKING_REFERENCE_COLUMN].strip().lower()] = invoice[REFERENCE_COLUMN]

cash_payments = []
wire_transfers = []

payments = [payment[16:] for payment in csv.reader(open(payments_filename,'r'))]
RECEIVED_DATE_COLUMN = 0
FORENAME_COLUMN = 1
BOOKING_REFERENCE_COLUMN = 5
PAYMENT_METHOD_COLUMN = 9
DIRECT1_COLUMN = 15

if (payments[0][RECEIVED_DATE_COLUMN] != 'ReceivedDate'
        or payments[0][FORENAME_COLUMN] != 'Forename'
        or payments[0][BOOKING_REFERENCE_COLUMN] != 'BookingReference'
        or payments[0][PAYMENT_METHOD_COLUMN] != 'PaymentMethod'
        or payments[0][DIRECT1_COLUMN] != 'Direct1'):
    raise Exception("Payment report changed format!")

for idx,payment in enumerate(payments[1:]):
    if not len(payment):
        continue

    if payment[RECEIVED_DATE_COLUMN] != '(see above)':
        date = datetime.datetime.strptime(payment[RECEIVED_DATE_COLUMN].split(' ')[0], '%d/%m/%Y')

    name = payment[FORENAME_COLUMN].strip()
    reservation = payment[BOOKING_REFERENCE_COLUMN]
    if reservation.strip().lower() in booking_to_invoice:
        reference = booking_to_invoice[reservation.strip().lower()]
    else:
        reference = reservation

    payment_method = payment[PAYMENT_METHOD_COLUMN]

    payment_value_raw = payment[DIRECT1_COLUMN]
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
