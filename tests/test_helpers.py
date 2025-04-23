# -*- coding: utf-8 -*-
'''
Created on 13.05.2020

@author: fstallmann
'''



import unittest
import os

import bkt.helpers as helpers


class IterTools(unittest.TestCase):
    def test_iterable(self):
        self.assertFalse(helpers.iterable(None))
        self.assertFalse(helpers.iterable(True))
        self.assertFalse(helpers.iterable(0))
        self.assertFalse(helpers.iterable(42.123))

        self.assertTrue(helpers.iterable("ABCD"))
        self.assertTrue(helpers.iterable([]))
        self.assertTrue(helpers.iterable([1, 2, 3]))
        self.assertTrue(helpers.iterable(list(range(3))))
        self.assertTrue(helpers.iterable(x for x in range(3)))

    def test_flatten(self):
        self.assertListEqual( list(helpers.flatten([list(range(2)), list(range(2,4))])), [0,1,2,3] )
        self.assertListEqual( list(helpers.flatten(list(range(i)) for i in range(4))), [0,0,1,0,1,2] )

    def test_all_equal(self):
        self.assertFalse(helpers.all_equal( list(range(2)) ))
        self.assertFalse(helpers.all_equal( [1,1,0,1,1] ))
        self.assertFalse(helpers.all_equal( i<3 for i in range(4) ))
        self.assertFalse(helpers.all_equal( i>0 for i in range(4) ))

        self.assertTrue(helpers.all_equal( [] ))
        self.assertTrue(helpers.all_equal( [1] ))
        self.assertTrue(helpers.all_equal( [1,1,1,1,1] ))
        self.assertTrue(helpers.all_equal( i<5 for i in range(4) ))
        self.assertTrue(helpers.all_equal( "ABC" for _ in range(4) ))

    def test_nth(self):
        self.assertEqual(helpers.nth(list(range(4,8)), 2), 6)
        self.assertEqual(helpers.nth(list(range(4,8)), 25), None)
        self.assertEqual(helpers.nth(["a", "b", "c"], 1, "x"), "b")
        self.assertEqual(helpers.nth(["a", "b", "c"], 5, "x"), "x")

    def test_ambiguity_tuple(self):
        self.assertTupleEqual(helpers.get_ambiguity_tuple([3,3,3,3,3]), (False, 3))
        self.assertTupleEqual(helpers.get_ambiguity_tuple([3,3,3,4,3]), (True, 3))
        self.assertTupleEqual(helpers.get_ambiguity_tuple(["a","a","a"]), (False, "a"))
        self.assertTupleEqual(helpers.get_ambiguity_tuple(["b","a","a"]), (True, "b"))

        self.assertTupleEqual(helpers.get_ambiguity_tuple(list(range(2,6))), (True, 2))
        self.assertTupleEqual(helpers.get_ambiguity_tuple(i>0 for i in range(2,6)), (False, True))



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

    def test_endings_to_windows(self):
        self.assertEqual(helpers.endings_to_windows("\nthis\nis\n\nsimple\n\n"), "\r\nthis\r\nis\r\n\r\nsimple\r\n\r\n")
        self.assertEqual(helpers.endings_to_windows("\rthis\ris\r\rsimple\r\r"), "\r\nthis\r\nis\r\n\r\nsimple\r\n\r\n")
        self.assertEqual(helpers.endings_to_windows("\r\nthis\r\nis\r\nsimple\r\n\r\n"), "\r\nthis\r\nis\r\nsimple\r\n\r\n")
        self.assertEqual(helpers.endings_to_windows("this\nis\nsimple", "B", "A"), "Athis\r\nBis\r\nBsimple")

    def test_endings_to_unix(self):
        self.assertEqual(helpers.endings_to_unix("\r\nthis\r\nis\r\n\r\nsimple\r\n\r\n"), "\nthis\nis\n\nsimple\n\n")
        self.assertEqual(helpers.endings_to_unix("\rthis\ris\r\rsimple\r\r"), "\nthis\nis\n\nsimple\n\n")
        self.assertEqual(helpers.endings_to_unix("\nthis\nis\nsimple\n\n"), "\nthis\nis\nsimple\n\n")



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

        ustring = "ÖÄÜß 😘 \r\n\tABC"
        self.cache["unicode_value"] = ustring
        self.cache[ustring] = "unicode_key"
        self.cache.sync()
        self.assertEqual(self.cache["unicode_value"], ustring)
        self.assertEqual(self.cache[ustring], "unicode_key")

    def test_cache_openclose(self):
        self.cache["testvalue"] = set([1,2,3])

        ustring = "ÖÄÜß 😘 \r\n\tABC"
        self.cache["unicode_value"] = ustring
        self.cache[ustring] = "unicode_key"
        
        helpers.caches.close("unittests")
        self.cache = helpers.caches.get("unittests")

        self.assertSetEqual(self.cache["testvalue"], set([1,2,3]))
        self.assertEqual(self.cache["unicode_value"], ustring)
        self.assertEqual(self.cache[ustring], "unicode_key")

class BitvalueAccessorTest(unittest.TestCase):
    def setUp(self):
        helpers.settings = dict()
    
    def test_with_attr_notation(self):
        bitvalue = helpers.BitwiseValueAccessor(8, ["test1", "test2", "test3", "test4"]) #1, 2, 4, 8

        self.assertFalse(bitvalue.test1)
        self.assertFalse(bitvalue.test2)
        self.assertFalse(bitvalue.test3)
        self.assertTrue(bitvalue.test4)

        self.assertEqual(bitvalue.get_bitvalue(), 8)

        bitvalue.test1 = True
        bitvalue.test2 = True
        bitvalue.test3 = True
        bitvalue.test4 = False

        self.assertTrue(bitvalue.test1)
        self.assertTrue(bitvalue.test2)
        self.assertTrue(bitvalue.test3)
        self.assertFalse(bitvalue.test4)

        self.assertEqual(bitvalue.get_bitvalue(), 7)

        with self.assertRaises(AttributeError):
            bitvalue.does_not_exist = False

        self.assertTrue(hasattr(bitvalue, "test3"))
        self.assertFalse(hasattr(bitvalue, "does_not_exist"))

        self.assertDictEqual(bitvalue.as_dict(), {"test1": True, "test2": True, "test3": True, "test4": False})

    def test_with_list_notation(self):
        helpers.settings["test.bitvalue"] = 1
        bitvalue = helpers.BitwiseValueAccessor(settings_key="test.bitvalue", attributes=["test1", "test2", "test3", "test4"]) #1, 2, 4, 8

        self.assertTrue(bitvalue["test1"])
        self.assertFalse(bitvalue["test2"])
        self.assertFalse(bitvalue["test3"])
        self.assertFalse(bitvalue["test4"])

        self.assertEqual(bitvalue.get_bitvalue(), 1)
        self.assertEqual(helpers.settings["test.bitvalue"], 1)

        bitvalue["test1"] = True
        bitvalue["test2"] = True
        bitvalue["test3"] = True
        bitvalue["test4"] = True

        self.assertTrue(bitvalue["test1"])
        self.assertTrue(bitvalue["test2"])
        self.assertTrue(bitvalue["test3"])
        self.assertTrue(bitvalue["test4"])

        self.assertEqual(bitvalue.get_bitvalue(), 15)
        self.assertEqual(helpers.settings["test.bitvalue"], 15)

        with self.assertRaises(KeyError):
            bitvalue["does_not_exist"] = False
        
        self.assertEqual(len(bitvalue), 4)

        self.assertTrue("test1" in bitvalue)
        self.assertFalse("does_not_exist" in bitvalue)

        self.assertListEqual(list(bitvalue), [("test1",True), ("test2",True), ("test3",True), ("test4",True)])

    def test_add_option(self):
        bitvalue = helpers.BitwiseValueAccessor(8, ["test1", "test2", "test3", "test4"]) #1, 2, 4, 8

        bitvalue.add_option("test5") #16
        self.assertFalse(bitvalue.test5)
        self.assertEqual(bitvalue.get_bitvalue(), 8)

        bitvalue.test5 = True
        self.assertTrue(bitvalue.test5)
        self.assertEqual(bitvalue.get_bitvalue(), 8+16)

        bitvalue.add_option("test6", True) #32
        self.assertTrue(bitvalue.test5)
        self.assertEqual(bitvalue.get_bitvalue(), 8+16+32)
    
    def test_repr(self):
        bitvalue = helpers.BitwiseValueAccessor(8, ["test1", "test2", "test3", "test4"]) #1, 2, 4, 8
        self.assertEqual(repr(bitvalue), "<BitwiseValueAccessor bitvalue=8 attributes=['test1', 'test2', 'test3', 'test4']>")