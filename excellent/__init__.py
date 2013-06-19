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
version = '0.0.2'

from collections import OrderedDict
from StringIO import StringIO
from excellent.backends import CSV, XL

__all__ = ['CSV', 'XL', 'Writer']


class Writer(object):
    def __init__(self, backend, output=None):
        self.backend = backend
        self.buffer = output or StringIO()

    def write(self, data, **kwargs):
        data = map(OrderedDict, data)
        self.backend.write(data, self.buffer, **kwargs)

    def save(self):
        self.backend.save(self.buffer)
