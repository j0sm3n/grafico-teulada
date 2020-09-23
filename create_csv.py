import csv
import locale
from datetime import datetime


# Needed to work properly with spanish months
locale.setlocale(locale.LC_ALL, 'es_ES.utf-8')
input_file = 'turnos copia.csv'
output_file = 'turnos.csv'


def split_list(a_list, elements):
    list_of_lists = []
    for i in range(0, len(a_list), elements):
        b_list = a_list[i:i + elements]
        list_of_lists.append(b_list)
    return list_of_lists


def create_csv():
    with open(input_file) as csv_file:
        csv_reader = csv.reader(csv_file)
        days = []
        for row in list(csv_reader):
            month_list = split_list(row, 21)
            for day in month_list:
                days.append(day)

    with open(output_file, mode='w') as f:
        f_writer = csv.writer(f)
        days[0].pop(1)
        f_writer.writerow(days[0])
        for row in days[1:]:
            row[0] = row[0].replace('.', '-2020')
            row_date = datetime.strptime(row[0], '%d-%b-%Y')
            row[0] = row_date.strftime('%Y-%m-%d')
            row.pop(1)
            f_writer.writerow(row)


if __name__ == '__main__':
    create_csv()
