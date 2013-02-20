#!/usr/bin/env python
# -*- coding: utf-8 -*-


class BaseBackend(object):
    def save(self, output):
        output.close()
