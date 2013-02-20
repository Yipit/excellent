#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
from excellent import Writer, CSV
from sure import expect


def test_write_data_with_header_to_csv():
    "Writer should write data with header correctly in CSV format"
    # Given a CSV writer
    writer = Writer(backend=CSV())

    # When we write some data
    data = [{"Country": "Argentina", "Revenue": 14500025}]
    writer.write(data)

    # Then the written to data should match our expectation
    (expect(writer.buffer.getvalue())
     .to.equal("Country,Revenue\r\nArgentina,14500025\r\n"))


def test_write_data_with_header_and_multiple_rows_to_csv():
    "Writer should write data with header and multiple rows correctly in CSV format"
    # Given a CSV writer
    writer = Writer(backend=CSV())

    # When we write some data
    data = [{"Country": "Argentina", "Revenue": 14500025},
            {"Country": "Brazil", "Revenue": 145002495}]
    writer.write(data)

    # Then the written to data should match our expectation
    expect(writer.buffer.getvalue()).to.equal("\r\n".join([
        "Country,Revenue",
        "Argentina,14500025",
        "Brazil,145002495",
        "",
    ]))


def test_writing_with_nondefault_delimiter():
    "Writer should allow nondefault delimiter"
    # Given a CSV writer
    writer = Writer(backend=CSV(delimiter=";"))

    # When we write some data
    data = [{"Country": "Argentina", "Revenue": 14500025},
            {"Country": "Brazil", "Revenue": 145002495}]
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
        {"Country": "USA", "Revenue": 22222},
        {"Country": "Canada", "Revenue": 33333},
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
