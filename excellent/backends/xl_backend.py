#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Copyright <2013> Gabriel Falcao <gabriel@yipit.com>
# Copyright <2013> Suneel Chakravorty <suneel@yipit.com>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

from xlwt import XFStyle, Alignment, Workbook
from .base import BaseBackend


class XL(BaseBackend):
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
            self.use_sheet('Sheet1')

        sheet = self.current_sheet
        header_style = self.get_header_style()

        for i, row in enumerate(data, self.current_row):
            if self.current_row is 0:
                self.write_row(sheet.row(0), row.keys(), header_style)
            self.write_row(sheet.row(i + 1), row.values())
            self.current_row = i + 1

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
        super(XL, self).save(output)
