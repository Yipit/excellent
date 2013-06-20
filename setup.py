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

import os
from setuptools import setup


def get_packages():
    # setuptools can't do the job :(
    packages = []
    for root, dirnames, filenames in os.walk('excellent'):
        if '__init__.py' in filenames:
            packages.append(".".join(os.path.split(root)).strip("."))

    return packages

setup(name='excellent',
    version='0.0.3',
    description='Python library for writing CSV and XLS files',
    author=u'Yipit',
    author_email='coders@yipit.com',
    url='http://tech.yipit.com/excellent',
    packages=get_packages(),
    install_requires=[
        "xlwt",
        "xlrd",
    ],
    package_data={
        'excellent': ['COPYING', '*.md'],
    },
)
