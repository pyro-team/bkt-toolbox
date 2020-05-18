# -*- coding: utf-8 -*-
'''
Created on 13.05.2020

@author: fstallmann
'''

from __future__ import absolute_import

import unittest
import os

import bkt.helpers as helpers


class StringConversions(unittest.TestCase):
    def test_lower_camelcase(self):
        self.assertEqual(helpers.snake_to_lower_camelcase("nothingToChange"), "nothingToChange")
        self.assertEqual(helpers.snake_to_lower_camelcase("FirstLower"), "firstLower")
        self.assertEqual(helpers.snake_to_lower_camelcase("this_IS_a_teST"), "thisIsATest")
        self.assertEqual(helpers.snake_to_lower_camelcase("THIS"), "tHIS")
        self.assertEqual(helpers.snake_to_lower_camelcase("CHANGE_THIS"), "changeThis")

    def test_upper_camelcase(self):
        self.assertEqual(helpers.snake_to_upper_camelcase("NothingToChange"), "NothingToChange")
        self.assertEqual(helpers.snake_to_upper_camelcase("firstUpper"), "FirstUpper")
        self.assertEqual(helpers.snake_to_upper_camelcase("this_IS_a_teST"), "ThisIsATest")
        self.assertEqual(helpers.snake_to_upper_camelcase("this"), "This")
        self.assertEqual(helpers.snake_to_upper_camelcase("CHANGE_THIS"), "ChangeThis")

class CacheTest(unittest.TestCase):
    def setUp(self):
        self.cache = helpers.caches.get("unittests")
        self.cache_file = self.cache._filename

    def tearDown(self):
        helpers.caches.close("unittests")
        base_path, cache_file = os.path.split(self.cache_file)
        files = os.listdir(base_path)
        for file in files:
            if file.startswith(cache_file):
                os.remove(os.path.join(base_path, file))

    def test_cache_rw(self):
        with self.assertRaises(KeyError):
            self.cache["testvalue"]

        self.cache["testvalue"] = set([1,2,3])
        self.assertSetEqual(self.cache["testvalue"], set([1,2,3]))

        ustring = u"ÖÄÜß 😘 \r\n\tABC"
        self.cache["unicode_value"] = ustring
        self.cache[ustring] = "unicode_key"
        self.cache.sync()
        self.assertEqual(self.cache["unicode_value"], ustring)
        self.assertEqual(self.cache[ustring], "unicode_key")

    def test_cache_openclose(self):
        self.cache["testvalue"] = set([1,2,3])

        ustring = u"ÖÄÜß 😘 \r\n\tABC"
        self.cache["unicode_value"] = ustring
        self.cache[ustring] = "unicode_key"
        
        helpers.caches.close("unittests")
        self.cache = helpers.caches.get("unittests")

        self.assertSetEqual(self.cache["testvalue"], set([1,2,3]))
        self.assertEqual(self.cache["unicode_value"], ustring)
        self.assertEqual(self.cache[ustring], "unicode_key")
