#!/usr/bin/env python
# -*- coding: utf-8 -*-
from StringIO import StringIO


class Writer(object):
    def __init__(self, backend, output=None):
        self.backend = backend
        self.buffer = output or StringIO()

    def write(self, data):
        self.backend.write(data, self.buffer)

    def save(self):
        self.buffer.close()

    def get_value(self):
        return self.buffer.getvalue()
