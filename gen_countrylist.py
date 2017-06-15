#!/usr/bin/env python3

import csv
import sys
import os

os.chdir(os.path.dirname(__file__))

filename = 'bookings.csv'
guest_filename = 'guests.csv'

guests = list(csv.reader(open(guest_filename, 'r')))
if (guests[0][7] != 'Email'
        or guests[0][8] != 'Telephone'
        or guests[0][12] != 'Country'):
    raise Exception("Guest report changed format!")

mail_to_country = {}
phone_to_country = {}
name_to_country = {}
for guest in guests[1:]:
    if not len(guest):
        continue
    try:
        country = guest[12]
        if not country:
            country = 'Germany'
        mail_to_country[guest[7].lower()] = country
        phone_to_country[guest[8].lower()] = country
    except Exception:
        print("Could not read guest %r" % (guest))
        raise

for addition in list(csv.reader(open('additions.csv', 'r')))[1:]:
    if not len(addition):
        continue
    mail = addition[0]
    phone = addition[1]
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

bookings = list(csv.reader(open(filename,'r',encoding='utf16')))
if (bookings[1][8] != 'Nächte'
    or bookings[1][13] != 'Erwachsene'
    or bookings[1][14] != 'Kinder'
    or bookings[1][23] != 'Buchungsstatus'
    or bookings[1][25] != 'Gast Email'
    or bookings[1][26] != 'Telefon 1'):
    raise Exception("Booking report changed format!")

country_people = {}
country_stays = {}

unknown = []
for booking in bookings[2:]:
    if booking[23] != 'Bestätigt':
        continue
    nights = int(booking[8])
    people = int(booking[13]) + int(booking[14])
    mail = booking[25]
    phone = booking[26]

    name = '%s %s' % (booking[4],booking[5])
    if mail in mail_to_country:
        country = mail_to_country[mail]
    elif phone in phone_to_country:
        country = phone_to_country[phone]
    elif name in name_to_country:
        name = name_to_country[name]
    else:
        #print("Could not lookup country for booking %r" % (booking))
        unknown_str = '%s %s Mail: %s Phone: %s' % (booking[4],booking[5],mail,phone)
        if unknown_str not in unknown:
            unknown.append(unknown_str)

    if country not in country_people:
        country_people[country] = 0
        country_stays[country] = 0

    country_people[country] += people
    country_stays[country] += people * nights


if len(unknown):
    print("Could not determine country for the following people:")
    for person in unknown:
        print(person)
    print("Please add a mapping to additions.csv")
    raise Exception("Need additions")

print("Persons by Country")
for country,people in sorted(country_people.items(), key=lambda x:-x[1]):
    print("%s\t%d" % (country, people))

print("\n\nStays by Country")
for country,stays in sorted(country_stays.items(), key=lambda x:-x[1]):
    print("%s\t%d" % (country, stays))
