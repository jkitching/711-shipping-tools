#!/bin/env python

import argparse
import csv
import logging
import sys

from openpyxl import load_workbook


FIELD_NAMES = [
    'sender_name',
    'sender_phone',
    'sender_email',
    'package_value',
    'store_name',
    'store_id',
    'receiver_name',
    'receiver_phone',
    'receiver_email',
    'return_store_name',
    'return_store_id'
]


def get_csv_reader(filename, fields):
    if fields:
        fields = fields.split(',')

    if filename == '-':
        csv_file = sys.stdin
    elif filename:
        csv_file = open(filename, 'r', encoding='utf-8')
    else:
        csv_file = open('/dev/null', 'r', encoding='utf-8')

    if fields:
        reader = csv.DictReader(csv_file, fieldnames=fields)
    else:
        reader = csv.DictReader(csv_file)

    return reader


def fill_row(row, field_names, default_values):
    output = []
    for field in field_names:
        if field in row:
            output.append(row[field])
        else:
            output.append(default_values[field])
    return output


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument('--template', help='The source xlsx file', required=True)
    parser.add_argument('--output', help='The output xlsx file', required=True)
    parser.add_argument('--csv', help='The source csv file')
    parser.add_argument('--csv-fields', help='Comma-separated csv field names')
    parser.add_argument('--verbose', action='store_true', help='Verbosity level')

    # Generate CLI arguments programmatically
    for field in FIELD_NAMES:
        parser.add_argument('--' + field, help='The ' + field.replace('_', ' '))

    # Parse arguments
    args = parser.parse_args()

    # Set up logging
    logging_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=logging_level,
                        format='%(asctime)s [%(levelname)s] %(message)s')

    csv_reader = get_csv_reader(args.csv, args.csv_fields)
    default_values = {field: getattr(args, field) for field in FIELD_NAMES}

    # Create a list of rows, filling in missing fields from field_names
    rows = []
    for row in csv_reader:
        rows.append([None] + fill_row(row, FIELD_NAMES, default_values))
    if not rows:
        rows.append([None] + fill_row([], FIELD_NAMES, default_values))

    wb = load_workbook(filename=args.template)
    ws = wb.active
    row_id = 4
    for row in rows:
        logging.debug('Fill row: %s', row[1:])
        for cell_id, cell in enumerate(ws[row_id]):
            cell.value = row[cell_id]
        row_id += 1
    wb.save(args.output)


if __name__ == '__main__':
    main()
