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
from __future__ import unicode_literals


import csv

from .base import BaseBackend


class CSV(BaseBackend):
    def __init__(self, delimiter=','):
        self.delimiter = bytes(delimiter)
        self.current_row = 0

    def write(self, data, output):
        csv_writer = csv.writer(output, delimiter=self.delimiter)
        for i, row in enumerate(data):
            if self.current_row is 0:
                csv_writer.writerow(row.keys())

            csv_writer.writerow(row.values())
            self.current_row = i + 1
