from html.parser import HTMLParser
import csv
import argparse
import time
import datetime

row_params = ['code', 'start_date', 'cal_weeks', 'start', 'end', 'module_name', 'description', 'type', 'locations',
              'staff', 'comment']


def main():
    parser = argparse.ArgumentParser(description='Convert rooster VU pages into .ics files')
    parser.add_argument('--i', type=str, help="input html file name", required=True)
    parser.add_argument('--o', type=str, help="output csv file name", required=True)
    args = parser.parse_args()

    html_string = file_to_string(args.i)

    html_parser = VUHtmlParser()

    html_parser.feed(html_string)

    with open(args.o, 'w', newline='') as output_csv:
        writer = csv.writer(output_csv)
        writer.writerow(['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time', 'All Day Event', 'Description',
                         'Location', 'Private'])

        for row in html_parser.rows:
            parse_row(writer, row)


def file_to_string(file):
    with open(file, 'r') as input_file:
        return input_file.read().replace('\n', '')


def parse_row(writer, row):
    weeks = []
    for cal_week in row['cal_weeks'].split(','):
        if '-' in cal_week:
            start_week = cal_week.split('-')[0]
            end_week = cal_week.split('-')[1]

            for i in range(int(start_week), int(end_week) + 1):
                weeks.append(i)
        else:
            weeks.append(int(cal_week))

    for week in weeks:
        add_event(writer, row, weeks[0], week)


def add_event(writer, row, start_week, week):
    day = datetime.datetime.strptime(row['start_date'], '%d/%m/%y')

    d = datetime.timedelta(weeks=week - start_week)
    day = day + d

    day = datetime.datetime.strftime(day, '%m/%d/%Y')

    writer.writerow([
        row['module_name'] + ' ' + row['type'] + ' ' + row['code'],
        day,
        time24to12(row['start']),
        day,
        time24to12(row['end']),
        False,
        get_description(row),
        row['locations'],
        False
    ])


def time24to12(str):
    return time.strftime('%I:%M %p', time.strptime(str, '%H:%M'))


def get_description(row):
    return 'staff: ' + row['staff'] + ' description: ' + row['description'] + ' comment: ' + row['comment']


def format_date(str):
    return time.strftime('%m/%d/%Y', )


class VUHtmlParser(HTMLParser):
    in_table = False
    ignore_row = False
    expect_data = False
    td_index = 0

    rows = []
    current_row = {}

    def error(self, message):
        pass

    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            class_tags = [attr for attr in attrs if attr[0] == 'class' and attr[1] == 'spreadsheet']

            if len(class_tags) > 0:
                self.in_table = True

        elif tag == 'tr':
            self.td_index = 0
            if self.in_table:
                self.current_row = {}

                class_tags = [attr for attr in attrs if attr[0] == 'class' and attr[1] == 'columnTitles']
                if len(class_tags) > 0:
                    self.ignore_row = True

        elif tag == 'td':
            if self.in_table and not self.ignore_row:
                self.expect_data = True

    def handle_endtag(self, tag):
        if tag == 'table':
            self.in_table = False

        elif tag == 'tr':
            if self.in_table and not self.ignore_row:
                self.rows.append(self.current_row)
            self.ignore_row = False

        elif tag == 'td':
            self.expect_data = False

    def handle_data(self, data):
        if self.expect_data:
            self.current_row[row_params[self.td_index]] = data.replace(u'\xa0', u' ')
            self.td_index += 1


if __name__ == '__main__':
    main()
