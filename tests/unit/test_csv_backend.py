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

from collections import OrderedDict
from mock import Mock, call, patch
from excellent import CSV


@patch('excellent.backends.csv_backend.csv')
def test_write(csv):
    "CSV backend writes data to output"

    # Given a csv writer
    csv_writer = csv.writer.return_value

    # And a mocked output and CSV backend
    csv_backend = CSV()
    output = Mock()

    # And data

    data = [
        OrderedDict([('Superhero', 'Superman'), ('Weakness', 'Kryptonite')]),
        OrderedDict([('Superhero', 'Spiderman'), ('Weakness', 'Maryjane')]),
    ]

    # When we write data to output
    csv_backend.write(data, output)

    # Then csv.writer was called once
    csv.writer.assert_called_once_with(output, delimiter=',')

    # And writerow received the appropriate calls

    csv_writer.writerow.assert_has_calls([
        call(["Superhero", "Weakness"]),
        call(["Superman", "Kryptonite"]),
        call(["Spiderman", "Maryjane"]),
    ])
