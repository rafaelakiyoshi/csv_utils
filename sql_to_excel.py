# -*- coding: utf-8 -*-
import os
import glob
import csv
import MySQLdb
from xlsxwriter.workbook import Workbook
import sys
reload(sys)
sys.setdefaultencoding('latin1')

user = 'root' # your username
passwd = 'root' # your password
host = 'localhost' # your host
db = 'lexana' # database where your table is stored
table = 'users' # table you want to save

connect = MySQLdb.connect(user=user, passwd=passwd, host=host, db=db)
cursor = connect.cursor()

query = "SELECT * FROM %s;" % table
cursor.execute(query)
workbook = Workbook('outfile.xlsx')
worksheet = workbook.add_worksheet()
for r, row in enumerate(cursor.fetchall()):
    print('Lendo linha...' + str(r))
    for c, col in enumerate(row):
        worksheet.write(r, c, col)
workbook.close();
