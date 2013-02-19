#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
from StringIO import StringIO


class Writer(object):
    def __init__(self, format, delimiter=',', output=None):
        self.format = format
        if not output:
            output = StringIO()

        self.buffer = output
        self.delimiter = delimiter

    def write(self, data):
        csv_writer = csv.writer(self.buffer, delimiter=self.delimiter)
        for i, row in enumerate(data):
            if i is 0:
                csv_writer.writerow(row.keys())

            csv_writer.writerow(row.values())

    def save(self):
        self.buffer.close()

    def get_value(self):
        return self.buffer.getvalue()
