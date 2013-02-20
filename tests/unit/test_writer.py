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

from collections import OrderedDict
from excellent import Writer
from mock import Mock


def test_write_init_with_no_output():
    "Initializing Writer with no output defaults to StringIO"
    # Given a backend
    backend = 'dummy-backend1'

    # When I create a writer with the backend
    writer = Writer(backend)

    # Then the writer's backend should be backend
    writer.should.have.property('backend').being.equal('dummy-backend1')

    # And the writer's buffer should be a StringIO instance
    writer.should.have.property('buffer').being.a('StringIO.StringIO')


def test_write_init_with_output():
    "Initializing Writer with an output assigns it accordingly"
    # Given a backend
    backend = 'dummy-backend2'

    # And an output
    output = 'some output'

    # When I create a writer with the backend
    writer = Writer(backend, output)

    # Then the writer's backend should be backend
    writer.should.have.property('backend').being.equal('dummy-backend2')

    # And the writer's buffer should equal 'some output'
    writer.should.have.property('buffer').being.equal('some output')


def test_writer_writes():
    "Writer should write data via the backend.write method"

    # Given a writer with a mocked backend
    backend = Mock()
    writer = Writer(backend)

    # When I write data
    data = [[('some awesome key', 'some awesome value')]]
    writer.write(data)

    # Then the backend.write should have been called once with
    # data and writer.buffer
    backend.write.assert_called_once_with([OrderedDict(data[0])],
                                          writer.buffer)


def test_writer_save():
    "Writer should close the buffer upon save"

    # Given a writer with a mocked backend
    backend = Mock()
    writer = Writer(backend)

    writer.save()
    backend.buffer.close.assert_called_once()
