# -*- coding: utf-8 -*-
import sys, codecs, datetime, os
from openpyxl import load_workbook

def main():
    datapoints = {}
    datapoints[-1] = 0
    newface = datetime.date.today().year - 1949 # 西暦から回期を求める
    for i in range(1,newface):
        datapoints[i] = 0

    stat = os.stat(sys.argv[1])
    dt = datetime.datetime.fromtimestamp(stat.st_mtime)
    last_modified = "\nvar last_modified = \"%s\"\n" % dt.strftime("%Y-%m-%d %H:%M:%S")

    book = load_workbook(filename=sys.argv[1])
    sheet = book[u'参加者']
    maxrow = sheet.max_row
    entries = 0
    rows = sheet.iter_rows('H2:H'+str(maxrow))
    for row in rows:
        for cell in row:
            if cell.value in range(1,newface):
                datapoints[cell.value] += 1
            elif isinstance(cell.value, unicode):
                print '不正な回期: %s' % str(cell.value)
            elif str(cell.value) == 'None':
                datapoints[-1] += 1
            else:
                print '不正な回期: %s' % str(cell.value)
            entries += 1
            if cell.value in range(newface--3,newface--1):
                print '学生: %s' % str(cell.value)

    total = "var total = %d\n" % entries

    dataPlot = "var dataPlot1 =["

    for k,v in sorted(datapoints.items(), key=lambda x: x[0]):
        if k == -1 :
            data = "{ label: %d, y: %d, indexLabel: " % (k, v)
            data += u"\"不明\" },"
        else:
            data = "{ label: %d, y: %d }," % (k, v)
        dataPlot += data

    dataPlot += "];\n"

    dataPlot += "var dataPlot2 =["

    for k,v in sorted(datapoints.items(), key=lambda x: x[1], reverse=True):
        if k == -1 :
            data = "{ label: %d, y: %d, indexLabel: " % (k, v)
            data += u"\"不明\" },"
        else:
            data = "{ label: %d, y: %d }," % (k, v)
        dataPlot += data

    dataPlot += "];"

    f = codecs.open('data.js', 'w', 'utf-8')
    f.write(dataPlot)
    f.write(last_modified)
    f.write(total)
    f.close()

if __name__ == "__main__":
    main()
