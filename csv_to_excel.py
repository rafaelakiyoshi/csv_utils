import os
import glob
import csv
from xlsxwriter.workbook import Workbook


for csvfile in glob.glob(os.path.join('.', '*.csv')):
    workbook = Workbook(csvfile + '_2.xlsx')
    worksheet = workbook.add_worksheet()
    c = 0
    with open(csvfile, 'r') as f:
        spamreader = csv.reader(f, delimiter=',', quotechar='"')
        for r, row in enumerate(spamreader):
            if (r > 230000):
                break
            if (r>=130000 and r<=230000):
                for c, col in enumerate(row):
                    print('Lendo linha...' + str(r))
                    worksheet.write(r, c, col)
    workbook.close()
