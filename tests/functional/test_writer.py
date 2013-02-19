#!/usr/bin/env python
# -*- coding: utf-8 -*-

from excellent import Writer
from sure import expect


def test_write_data_with_header_to_csv():
    "Writer should write data with header correctly in CSV format"
    # Given a CSV writer
    writer = Writer(format="csv")

    # When we write some data
    data = [{"Country": "Argentina", "Revenue": 14500025}]
    writer.write(data)

    # Then the written to data should match our expectation
    (expect(writer.get_value())
     .to.equal("Country,Revenue\r\nArgentina,14500025\r\n"))


def test_write_data_with_header_and_multiple_rows_to_csv():
    "Writer should write data with header and multiple rows correctly in CSV format"
    # Given a CSV writer
    writer = Writer(format="csv")

    # When we write some data
    data = [{"Country": "Argentina", "Revenue": 14500025},
            {"Country": "Brazil", "Revenue": 145002495}]
    writer.write(data)

    # Then the written to data should match our expectation
    expect(writer.get_value()).to.equal("\r\n".join([
        "Country,Revenue",
        "Argentina,14500025",
        "Brazil,145002495",
        "",
    ]))


def test_writing_with_nondefault_delimiter():
    "Writer should allow nondefault delimiter"
    # Given a CSV writer
    writer = Writer(format="csv", delimiter=";")

    # When we write some data
    data = [{"Country": "Argentina", "Revenue": 14500025},
            {"Country": "Brazil", "Revenue": 145002495}]
    writer.write(data)

    # Then the written to data should match our expectation
    expect(writer.get_value()).to.equal("\r\n".join([
        "Country;Revenue",
        "Argentina;14500025",
        "Brazil;145002495",
        "",
    ]))
