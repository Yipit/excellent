#!/usr/bin/env python
# -*- coding: utf-8 -*-
from collections import OrderedDict
from StringIO import StringIO
from excellent.backends import CSV, XL

__all__ = ['CSV', 'XL', 'Writer']


class Writer(object):
    def __init__(self, backend, output=None):
        self.backend = backend
        self.buffer = output or StringIO()

    def write(self, data):
        data = map(OrderedDict, data)
        self.backend.write(data, self.buffer)

    def save(self):
        self.backend.save(self.buffer)
