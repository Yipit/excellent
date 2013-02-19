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
