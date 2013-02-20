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

import os
from collections import OrderedDict
from excellent import Writer, CSV
from sure import expect


def test_write_data_with_header_to_csv():
    "Writer should write data with header correctly in CSV format"
    # Given a CSV writer
    writer = Writer(backend=CSV())

    # When we write some data
    data = [OrderedDict([("Country", "Argentina"), ("Revenue", 14500025)])]
    writer.write(data)

    # Then the written to data should match our expectation
    (expect(writer.buffer.getvalue())
     .to.equal("Country,Revenue\r\nArgentina,14500025\r\n"))


def test_write_data_with_header_and_multiple_rows_to_csv():
    ("Writer should write data with header and multiple "
     "rows correctly in CSV format")
    # Given a CSV writer
    writer = Writer(backend=CSV())

    # When we write some data
    data = [
        OrderedDict([("Country", "Argentina"), ("Revenue", 14500025)]),
        OrderedDict([("Country", "Brazil"), ("Revenue", 145002495)]),
    ]
    writer.write(data)

    # Then the written to data should match our expectation
    expect(writer.buffer.getvalue()).to.equal("\r\n".join([
        "Country,Revenue",
        "Argentina,14500025",
        "Brazil,145002495",
        "",
    ]))


def test_writing_with_nondefault_delimiter_byte():
    "Writer should allow nondefault delimiter (byte)"
    # Given a CSV writer
    writer = Writer(backend=CSV(delimiter=b";"))

    # When we write some data
    data = [
        OrderedDict([("Country", "Argentina"), ("Revenue", 14500025)]),
        OrderedDict([("Country", "Brazil"), ("Revenue", 145002495)]),
    ]
    writer.write(data)

    # Then the written to data should match our expectation
    expect(writer.buffer.getvalue()).to.equal("\r\n".join([
        "Country;Revenue",
        "Argentina;14500025",
        "Brazil;145002495",
        "",
    ]))


def test_writing_with_nondefault_delimiter_unicode():
    "Writer should allow nondefault delimiter (unicode)p"
    # Given a CSV writer
    writer = Writer(backend=CSV(delimiter=";"))

    # When we write some data
    data = [
        OrderedDict([("Country", "Argentina"), ("Revenue", 14500025)]),
        OrderedDict([("Country", "Brazil"), ("Revenue", 145002495)]),
    ]
    writer.write(data)

    # Then the written to data should match our expectation
    expect(writer.buffer.getvalue()).to.equal("\r\n".join([
        "Country;Revenue",
        "Argentina;14500025",
        "Brazil;145002495",
        "",
    ]))


def test_writing_to_a_file():
    "Writer should be able to write to a file"
    # Given a temporary file
    tmpfile = open('/tmp/test_1.csv', 'w+')

    # And CSV writer pointing to that tempfile
    writer = Writer(backend=CSV(delimiter=";"), output=tmpfile)

    # And some data
    writer.write([
        OrderedDict([("Country", "USA"), ("Revenue", 22222)]),
        OrderedDict([("Country", "Canada"), ("Revenue", 33333)]),
    ])

    # When the writer gets saved
    writer.save()

    # Then the file exists
    os.path.exists(tmpfile.name).should.be.true
    expect(open(tmpfile.name).read()).to.equal("\r\n".join([
        "Country;Revenue",
        "USA;22222",
        "Canada;33333",
        "",
    ]))


def test_writing_twice():
    "Writer should be able to write to a file"
    # Given a temporary file
    tmpfile = open('/tmp/test_1.csv', 'w+')

    # And CSV writer pointing to that tempfile
    writer = Writer(backend=CSV(delimiter=";"), output=tmpfile)

    # And some data
    writer.write([
        OrderedDict([("Country", "USA"), ("Revenue", 22222)]),
    ])
    writer.write([
        OrderedDict([("Country", "Canada"), ("Revenue", 33333)]),
    ])

    writer.write([
        OrderedDict([("Country", "Italia"), ("Revenue", 1)]),
    ])

    # When the writer gets saved
    writer.save()

    # Then the file exists
    os.path.exists(tmpfile.name).should.be.true
    expect(open(tmpfile.name).read()).to.equal("\r\n".join([
        "Country;Revenue",
        "USA;22222",
        "Canada;33333",
        "Italia;1",
        "",
    ]))
