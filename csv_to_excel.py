import os
import glob
import csv
from xlsxwriter.workbook import Workbook


for csvfile in glob.glob(os.path.join('.', '*.csv')):
    workbook = Workbook(csvfile + '_2.xlsx')
    worksheet = workbook.add_worksheet()
    offset = 130000
    limit = 230000
    with open(csvfile, 'r') as f:
        spamreader = csv.reader(f, delimiter=',', quotechar='"')
        for r, row in enumerate(spamreader):

            if (r!=0 and r > limit):
                print('breko')
                break

            if (r == 0 or (r>=offset and r<=limit)):
                print('Lendo linha...' + str(r))
                for c, col in enumerate(row):
                    rNum = r if r==0 else r-offset+1
                    worksheet.write(rNum, c, col)
    workbook.close()
