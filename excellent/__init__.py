#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
from StringIO import StringIO


class Writer(object):
    def __init__(self, format):
        self.format = format
        self.buffer = StringIO()

    def write(self, data):
        csv_writer = csv.writer(self.buffer)
        for i, row in enumerate(data):
            if i is 0:
                csv_writer.writerow(row.keys())

            csv_writer.writerow(row.values())

    def get_value(self):
        return self.buffer.getvalue()
