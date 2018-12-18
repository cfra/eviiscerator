#!/usr/bin/env python3

import csv
import sys
import os
import datetime

basedir = os.path.dirname(__file__)
if basedir:
    os.chdir(basedir)

filename = 'bookings.csv'
guest_filename = 'guests.csv'

guests = list(csv.reader(open(guest_filename, 'r')))

EMAIL_COLUMN = None
PHONE_COLUMN = None
COUNTRY_COLUMN = None

for idx, field in enumerate(guests[0]):
    if field == 'Email':
        EMAIL_COLUMN = idx
    elif field == 'Telephone':
        PHONE_COLUMN = idx
    elif field == 'Country':
        COUNTRY_COLUMN = idx

if EMAIL_COLUMN is None \
   or PHONE_COLUMN is None \
   or COUNTRY_COLUMN is None:
    raise Exception("Guest report is missing a required column!")

mail_to_country = {}
phone_to_country = {}
name_to_country = {}
for guest in guests[1:]:
    if not len(guest):
        continue
    try:
        country = guest[COUNTRY_COLUMN]
        if not country:
            country = 'Germany'
        mail_to_country[guest[EMAIL_COLUMN].lower()] = country
        phone_to_country[guest[PHONE_COLUMN].lower()] = country
    except Exception:
        print("Could not read guest %r" % (guest))
        raise

if not os.path.exists('additions.csv'):
    with open('additions.csv','w') as f:
        f.write('Mail,Phone,Country,Name\n')
for addition in list(csv.reader(open('additions.csv', 'r')))[1:]:
    if not len(addition):
        continue
    mail = addition[0]
    phone = addition[1] if len(addition) >= 2 else ''
    country = addition[2] if len(addition) >= 3 else ''
    if not country:
        country = "Germany"

    name = addition[3] if len(addition) >= 4 else ''

    if mail:
        mail_to_country[mail] = country
    if phone:
        phone_to_country[phone] = country
    if name:
        name_to_country[name] = country

BOOKING_GIVEN_NAME_COLUMN = 4
BOOKING_SURNAME_COLUMN = 5
BOOKING_CHECKIN_COLUMN = 6
BOOKING_NIGHTS_COLUMN = 8
BOOKING_ADULTS_COLUMN = 13
BOOKING_CHILDREN_COLUMN = 14
BOOKING_STATUS_COLUMN = 24
BOOKING_EMAIL_COLUMN = 26
BOOKING_TELEPHONE_COLUMN = 27

bookings = list(csv.reader(open(filename,'r',encoding='utf16')))
if (bookings[1][BOOKING_GIVEN_NAME_COLUMN] != 'Vorname'
    or bookings[1][BOOKING_SURNAME_COLUMN] != 'Nachname'
    or bookings[1][BOOKING_CHECKIN_COLUMN] != 'Einchecken'
    or bookings[1][BOOKING_NIGHTS_COLUMN] != 'Nächte'
    or bookings[1][BOOKING_ADULTS_COLUMN] != 'Erwachsene'
    or bookings[1][BOOKING_CHILDREN_COLUMN] != 'Kinder'
    or bookings[1][BOOKING_STATUS_COLUMN] != 'Buchungsstatus'
    or bookings[1][BOOKING_EMAIL_COLUMN] != 'Gast Email'
    or bookings[1][BOOKING_TELEPHONE_COLUMN] != 'Telefon 1'):
    raise Exception("Booking report changed format!")

country_people = {}
country_stays = {}

first_checkin = None
last_checkin = None
unknown = []
for booking in bookings[2:]:
    if booking[BOOKING_STATUS_COLUMN] != 'Bestätigt':
        continue
    nights = int(booking[BOOKING_NIGHTS_COLUMN])
    people = int(booking[BOOKING_ADULTS_COLUMN]) + int(booking[BOOKING_CHILDREN_COLUMN])
    mail = booking[BOOKING_EMAIL_COLUMN]
    phone = booking[BOOKING_TELEPHONE_COLUMN]
    checkin = datetime.datetime.strptime(booking[BOOKING_CHECKIN_COLUMN], '%d-%b-%y')
    if first_checkin is None or checkin < first_checkin:
        first_checkin = checkin
    if last_checkin is None or checkin > last_checkin:
        last_checkin = checkin

    name = '%s %s' % (booking[BOOKING_GIVEN_NAME_COLUMN],booking[BOOKING_SURNAME_COLUMN])
    if mail in mail_to_country:
        country = mail_to_country[mail]
    elif phone in phone_to_country:
        country = phone_to_country[phone]
    elif name in name_to_country:
        name = name_to_country[name]
    else:
        #print("Could not lookup country for booking %r" % (booking))
        unknown_info = (mail,phone,name)
        if unknown_info not in unknown:
            unknown.append(unknown_info)

    if country not in country_people:
        country_people[country] = 0
        country_stays[country] = 0

    country_people[country] += people
    country_stays[country] += people * nights


if len(unknown):
    print("Could not determine country for the following people, please enter:")
    for person in unknown:
        mail = person[0] if not 'Bitte die Buchung entsperren' in person[0] else ''
        phone = person[1] if not 'Bitte die Buchung entsperren' in person[1] else ''
        name = person[2]
        country = input('Whats the country for %s (Mail: %r, Phone: %r)? ' % (name, mail, phone))
        with open('additions.csv','a') as f:
            line = '%s,%s,%s,%s\n' % (mail,phone,country,name)
            f.write(line)
    print("All done. Please restart.")
    sys.exit(1)

print("Information for %s to %s:\n\n" % (
      first_checkin.strftime('%d. %B %Y'),
      last_checkin.strftime('%d. %B %Y')))

print("Persons by Country")
for country,people in sorted(country_people.items(), key=lambda x:-x[1]):
    print("%s\t%d" % (country, people))

print("\n\nStays by Country")
for country,stays in sorted(country_stays.items(), key=lambda x:-x[1]):
    print("%s\t%d" % (country, stays))
