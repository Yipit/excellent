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

import copy
from xlwt import XFStyle, Alignment, Workbook
from .base import BaseBackend

default_style = XFStyle()
bold_style = XFStyle()
bold_style.alignment.horz = Alignment.HORZ_RIGHT
bold_style.font.bold = True

# Excel has issues when creating too many styles/fonts, hence we use
# a cache to reuse font instances (see FAQ#13 http://poi.apache.org/faq.html)
STYLE_CACHE = {}


def hash_style(style):
    """
    This ugly function allows us to get a hash for xlwt Style instances. The
    hash allows us to determine that two Style instances are the same, even if
    they are different objects.
    """
    font_attrs = ["font", "alignment", "borders", "pattern", "protection"]
    attrs_hashes = [hash(frozenset(getattr(style, attr).__dict__.items())) for attr in font_attrs]
    return hash(sum(attrs_hashes + [hash(style.num_format_str)]))


class XL(BaseBackend):
    def __init__(self, workbook=None, default_style=default_style):
        self.workbook = workbook or Workbook()
        self.current_sheet = None
        self.current_row = 0
        self.default_style = default_style

    def get_header_style(self):
        return bold_style

    def write_row(self, row, values, style=None, **kwargs):
        style = style or self.default_style

        if kwargs:
            # If there are additional changes in kwargs, we don't want to modify
            # the original style, so we make a copy
            style = copy.deepcopy(style)

            if 'bold' in kwargs:
                style.font.bold = kwargs['bold']
            if 'bottom_border' in kwargs:
                style.borders.bottom = 2
            if 'format_string' in kwargs:
                style.num_format_str = kwargs['format_string']

        style_hash = hash_style(style)
        if style_hash in STYLE_CACHE:
            style = STYLE_CACHE[style_hash]
        else:
            STYLE_CACHE[style_hash] = style

        for index, value in enumerate(values):
            row.write(index, value, style)

    def write(self, data, output, style=None, **kwargs):
        if not self.current_sheet:
            self.use_sheet('Sheet1')

        sheet = self.current_sheet
        header_style = self.get_header_style()

        for i, row in enumerate(data, self.current_row):
            if self.current_row is 0:
                self.write_row(sheet.row(0), row.keys(), header_style, **kwargs)
            self.write_row(sheet.row(i + 1), row.values(), style=style, **kwargs)
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
