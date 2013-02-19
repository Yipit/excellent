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
