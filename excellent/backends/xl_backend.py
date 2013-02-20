#!/usr/bin/env python
# -*- coding: utf-8 -*-
from xlwt import Font, XFStyle, Borders, Alignment, Workbook


class XL(object):
    def __init__(self, delimiter=','):
        self.workbook = Workbook()
        self.current_sheet = None

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
        if not self.current_sheet:
            self.current_sheet = self.workbook.add_sheet('Sheet1')
        sheet = self.current_sheet
        for i, row in enumerate(data):
            if i is 0:
                self.write_row(sheet.row(0), row.keys(), self.get_header_style())

            self.write_row(sheet.row(i + 1), row.values())

    def use_sheet(self, name):
        self.current_sheet = self.workbook.add_sheet(name)

    def save(self, output):
        self.workbook.save(output)
        output.close()
