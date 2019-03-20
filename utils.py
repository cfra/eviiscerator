import datetime

month_to_number = {
    'jan': '01',
    'feb': '02',
    'mrz': '03',
    'apr': '04',
    'mai': '05',
    'jun': '06',
    'jul': '07',
    'aug': '08',
    'sep': '09',
    'okt': '10',
    'nov': '11',
    'dez': '12',
}

def replace_by_dict(input_string, replacement_dict):
    for search, replace in replacement_dict.items():
        input_string = input_string.replace(search, replace)
    return input_string

def parse_old_booking_format(date_input):
    return datetime.datetime.strptime(date_input.split(' ')[0], '%d/%m/%Y')

def parse_new_booking_format(date_input):
    date_input = replace_by_dict(date_input.lower(), month_to_number)
    return datetime.datetime.strptime(date_input.split(' ')[0], '%d-%m-%Y')

date_parsers = [
    parse_old_booking_format,
    parse_new_booking_format
]

def parse_date(date_input):
    for parser in date_parsers:
        try:
            return parser(date_input)
        except ValueError:
            continue
    else:
        raise ValueError("Could not parse date %r" % date_input)
