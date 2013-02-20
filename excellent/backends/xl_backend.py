#!/usr/bin/env python
# -*- coding: utf-8 -*-
from xlwt import XFStyle, Alignment, Workbook


class XL(object):
    def __init__(self, workbook=None):
        self.workbook = workbook or Workbook()
        self.current_sheet = None
        self.current_row = 0

    def get_header_style(self):
        style = XFStyle()
        style.alignment.horz = Alignment.HORZ_RIGHT
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
            if i is 0 and self.current_row is 0:
                self.write_row(sheet.row(0), row.keys(), self.get_header_style())
            self.write_row(sheet.row(i + self.current_row + 1), row.values())

    def get_sheets(self):
        return self.workbook._Workbook__worksheets

    def get_or_create_sheet(self, name):
        for sheet in self.get_sheets():
            if sheet.name == name:
                return sheet, sheet.rows and max(sheet.rows.keys()) or 0
        return self.workbook.add_sheet(name), 0

    def use_sheet(self, name):
        self.current_sheet, self.current_row = self.get_or_create_sheet(name)

    def save(self, output):
        self.workbook.save(output)
        output.close()
