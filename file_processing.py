
import re
import csv
import json
from datetime import datetime
from collections import OrderedDict

import xlrd
import requests

month_dict = {
    'Jan': 1,
    'Feb': 2,
    'Mar': 3,
    'Apr': 4,
    'May': 5,
    'Jun': 6,
    'Jul': 7,
    'Aug': 8,
    'Sep': 9,
    'Oct': 10,
    'Nov': 11,
    'Dec': 12,
}


class ProcessFile(object):
    """Process only the new records since last update and save them in an
    output file."""

    def __init__(
            self,
            file_path,
            header_properties,
            check_date,
            last_saved_date,
            offset,
    ):

        self.file_path = file_path
        self.offset = offset
        self.header_properties = header_properties
        self.check_date = check_date
        self.last_saved_date = datetime.strptime(last_saved_date, '%Y-%m-%d')
        self.worksheet = None

        self.col_start_from = 2

    def process(self):
        """Processing file."""

        file_name = self.download_file()

        workbook = xlrd.open_workbook(file_name, on_demand=True)
        self.worksheet = workbook.sheet_by_index(0)

        header = ['Date'] + [self.header_properties.get('prefix')] * (
            self.worksheet.ncols - self.col_start_from)
        self.generate_header(header)

        rows = [header]
        last_date_processed = self.read_file_data(rows)

        # Close the workbook
        workbook.release_resources()
        del workbook

        self.write_file_data(rows)

        return str(last_date_processed.date())

    def download_file(self):
        """download file from the link to the present folder"""

        file_name = self.file_path.get('input').split("/")[-1]
        response = requests.get(self.file_path.get('input'))
        with open(file_name, 'w') as output_file_obj:
            output_file_obj.write(response.content)

        return file_name

    def write_file_data(self, rows):
        """Writes the data in csv format."""

        with open(self.file_path.get('output'), 'w') as output_file_obj:
            writer = csv.writer(output_file_obj)
            writer.writerows(rows)

    def generate_header(self, header):
        """Generates the header of the file by clubing all the header rows."""

        for row in xrange(self.offset.get('top'), self.offset.get('header')):

            old_data = ''
            for col in xrange(self.worksheet.ncols - self.col_start_from):
                data = re.sub(
                    r'\d\/',
                    '',
                    self.worksheet.cell_value(row, col + self.col_start_from)
                )
                data = re.sub(r' ', '_', data)

                remove_line_from_headers = '|'.join(
                    '\\' + _ for _ in self.header_properties.get(
                        "remove_line_from_headers"))

                if (remove_line_from_headers and
                        re.findall(remove_line_from_headers, data)):
                    data = ''
                if (header[col + 1] == self.header_properties.get('prefix')
                        and not data):
                    data = old_data
                if data:
                    header[col + 1] += '{}{}'.format(
                        '_' if header[col + 1] else '', data)
                old_data = data

    def read_file_data(self, rows):
        """Read the selected data from the workbook object."""

        current_year = 0
        current_month = 0
        current_day = 0 if self.check_date else 1

        for row in xrange(
                self.offset.get('header'),
                self.worksheet.nrows - self.offset.get('bottom')):

            if month_dict.get(self.worksheet.cell_value(row, 1), 0):
                current_month = month_dict.get(
                    self.worksheet.cell_value(row, 1), 0)
                if self.worksheet.cell_value(row, 0):
                    current_year = int(self.worksheet.cell_value(row, 0))
                if self.check_date:
                    continue
            elif (self.check_date and
                  isinstance(self.worksheet.cell_value(row, 1), float)):
                current_day = int(self.worksheet.cell_value(row, 1))
            else:
                continue

            if not (current_year and current_month):
                # to check if dates are populating can be removed
                continue

            current_date = datetime(
                year=current_year, month=current_month, day=current_day)
            if self.last_saved_date >= current_date:
                continue

            r_list = ['{d.month}/{d.day}/{d.year}'.format(d=current_date), ]
            for col in xrange(self.worksheet.ncols):
                if col >= self.col_start_from:
                    r_list.append(self.worksheet.cell_value(row, col))

            rows.append(r_list)

        return current_date


if __name__ == '__main__':
    with open('file_configs.json', 'r') as file_obj:

        file_configrations = json.load(file_obj, object_pairs_hook=OrderedDict)

    for file_config in file_configrations:
        file_obj = ProcessFile(**file_config)
        file_config['last_saved_date'] = file_obj.process()

    with open('file_configs.json', 'w') as file_obj:
        json.dump(file_configrations, file_obj, indent=4, sort_keys=True)
