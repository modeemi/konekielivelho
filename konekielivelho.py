#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
A small program for parsing (Nordea formatted) TITO output from business online bank.
TITO is the Finnish account statement format for storing bank account information in files.

This software is licensed under the MIT license and is provided as-is;
author assumes no responsibility for the use of this software by other parties.

Feel free to customize to your own needs; this is a very specific implementation.

Thanks to Danskebank for almost working manual (not 1:1 match to Nordea output):

    http://www-2.danskebank.com/Link/Tito/$file/Tito.pdf

The full specification can be found at:

    https://www.fkl.fi/materiaalipankki/ohjeet/Dokumentit/Konekielinen_tiliote_palvelukuvaus.pdf

Default behaviour is to read input.tito from current folder and parse output.csv from it:

    ./konekielivelho.py

Is equal to:

    ./konekielivelho.py --input input.tito --output output.csv
"""

import re
import csv
from datetime import datetime
from optparse import OptionParser

__author__ = 'Aleksi Häkli'
__copyright__ = 'Aleksi Häkli'
__license__ = 'MIT'
__version__ = '0.1.0'
__date__ = '2015-04-02'
__status__ = 'MVP'


class BankEventType():
    HEADER = 'T00'
    TRANSACTION = 'T10'
    SPECIFICATION = 'T11'
    BALANCE = 'T40'
    CUMULATIVE = 'T50'

    DESCRIPTION_MATCHER = re.compile(r'^T11...00.*$', flags=re.MULTILINE)
    NOT_INTERESTING_MATCHER = re.compile(  # Sanitize some unnecessary fields aside payments
        r'^.*(PALVELUMAKSU|S.*ST.*LIPAS|ITSEPALVELU).*$'
        , flags=re.IGNORECASE | re.MULTILINE)

    @staticmethod
    def is_transaction(string):
        return string.startswith(BankEventType.TRANSACTION)

    @staticmethod
    def is_description(string):
        return re.match(BankEventType.DESCRIPTION_MATCHER, string)

    @staticmethod
    def is_interesting(string):
        return not re.match(BankEventType.NOT_INTERESTING_MATCHER, string)


class Payment():
    """
    A simple class for manazing serialization and deserialization of payment objects.
    Add any processing logic to parse_from_list you might require.
    """

    DICT_FIELDS = ['datetime', 'amount', 'payer', 'reference', 'message']

    def __init__(self):
        self.datetime = datetime(year=1970, month=1, day=1)
        self.amount = int(0)
        self.payer = str()
        self.reference = str()
        self.message = str()

    def __repr__(self):  # for readable print() with the object
        return '%s,%5.2f,%s,%s,"%s"' % (
            self.datetime, self.amount, self.payer, self.reference or '', self.message)

    @staticmethod
    def parse_from_list(rows):
        payment_row = rows[0]

        str_datetime = payment_row[42:48].strip()
        str_amount = re.sub(r'[,.]', '', payment_row[87:106].strip())
        str_payer = payment_row[108:143].strip()
        str_reference = payment_row[160:179].strip().lstrip('0')

        payment = Payment()
        payment.datetime = datetime.strptime(str_datetime, '%y%m%d')
        payment.amount = int(str_amount)
        payment.payer = str_payer
        payment.reference = str_reference

        # Parse possible multiline message,
        # assumes that input beyond first row
        # of bank transaction is always just the message.
        if len(rows) > 1:
            def parse_message(msg_string):
                if len(msg_string) <= 8:
                    return ''
                return msg_string[8::].strip()

            payment.message = '\n'.join(map(parse_message, rows[1::]))

        return payment


def convert_ascii_alphabets(string):
    """
    Parse and convert good old ISO 646 ASCII encoded character string.
    Characters in @[\]{|} range are reserved to language-specific special characters,
    this function converts them to the Finnish / Scandic interpretation of the subset.
    """

    substitutions = [(r'\]', 'Å'), (r'\[', 'Ä'), (r'\\', 'Ö'), (r'\}', 'å'), (r'\{', 'ä'), (r'\|', 'ö')]
    for s in substitutions:
        string = re.sub(s[0], s[1], string)
    return string

def main():
    parser = OptionParser(usage='%prog [options]\n' + __doc__)
    parser.add_option('-i', '--input', dest='input', help='TITO file to parse', default='input.tito')
    parser.add_option('-o', '--output', dest='output', help='CSV file to output', default='output.csv')
    (options, args) = parser.parse_args()

    with open(options.input, mode='r') as f:
        filecontents = f.read()

    contents = convert_ascii_alphabets(filecontents).splitlines()
    transactions = []

    for (i, v) in enumerate(contents):
        if BankEventType.is_transaction(v) and BankEventType.is_interesting(v):
            transaction = list()
            transaction.append(v)

            # Find following message lines; this is specific to only message strings
            # At the moment no other types of specification string are parsed
            # from the statement, so add any logic you need for that here.
            j = i
            while True:
                j += 1
                if BankEventType.is_description(contents[j]):
                    transaction.append(contents[j])
                else:
                    break

            transactions.append(transaction)

    if transactions:
        payments = map(Payment.parse_from_list, transactions)
        with open(options.output, encoding='utf-8', mode='w') as f:
            writer = csv.DictWriter(f, fieldnames=Payment.DICT_FIELDS)
            writer.writeheader()
            for p in payments:
                writer.writerow(p.__dict__)

        print('Wrote %d transactions parsed from %s to %s' % (
            len(transactions), options.input, options.output))
    else:
        print('\n'.join([
            'Input file did not have any interesting transactions, '
            , 'cowardly refusing to produce an output file from such data.']))

if __name__ == '__main__':
    main()
