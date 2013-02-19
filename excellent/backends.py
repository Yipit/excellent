#!/usr/bin/env python
# -*- coding: utf-8 -*-
import csv


class CSV(object):
    def __init__(self, delimiter=','):
        self.delimiter = delimiter

    def write(self, data, output):
        csv_writer = csv.writer(output, delimiter=self.delimiter)
        for i, row in enumerate(data):
            if i is 0:
                csv_writer.writerow(row.keys())

            csv_writer.writerow(row.values())


from xlwt import Font, XFStyle, Borders, Alignment, Workbook


class XL(object):
    def __init__(self, delimiter=','):
        self.workbook = Workbook()

    def get_header_style(self):
        style = XFStyle()
        style.alignment = Alignment()
        style.alignment.horz = Alignment.HORZ_RIGHT
        style.borders = Borders()
        style.font = Font()
        style.font.bold = True
        return style

    def write_row(self, row, values, style=None):
        for index, value in enumerate(values):
            row.write(index, value, style or XFStyle())

    def write(self, data, output):
        sheet = self.workbook.add_sheet('Sheet1')
        for i, row in enumerate(data):
            if i is 0:
                self.write_row(sheet.row(0), row.keys(), self.get_header_style())

            self.write_row(sheet.row(i + 1), row.values())

        self.workbook.save(output)
