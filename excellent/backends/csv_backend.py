#!/usr/bin/env python
# -*- coding: utf-8 -*-
import csv

from .base import BaseBackend


class CSV(BaseBackend):
    def __init__(self, delimiter=','):
        self.delimiter = delimiter
        self.current_row = 0

    def write(self, data, output):
        csv_writer = csv.writer(output, delimiter=self.delimiter)
        for i, row in enumerate(data):
            if self.current_row is 0:
                csv_writer.writerow(row.keys())

            csv_writer.writerow(row.values())
            self.current_row = i + 1
