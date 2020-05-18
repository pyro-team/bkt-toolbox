# -*- coding: utf-8 -*-
'''
Created on 13.05.2020

@author: fstallmann
'''

from __future__ import absolute_import

import unittest

from collections import namedtuple

from bkt.library import search

search.settings = dict() #bkt.settings would return closed dict!


def get_list_by_indices(mylist, *args):
    return [mylist[i] for i in args]


class SearchDefaultDocuments(unittest.TestCase):
    def setUp(self):
        search.settings = dict()
        self.search_engine = search.get_search_engine("tests")
    
    def tearDown(self):
        del search.search_engines["tests"]
        self.search_engine = None

    def test_add_documents_with_string_keywords(self):
        search_writer = self.search_engine.writer()
        search_writer.add_document(module="test", name="doc 1", keywords="this is a TEST with string keywords.\nThis is, also a test with multiple-key-wörds.")
        search_writer.add_document(module="test", name="doc 1", keywords="this is a TEST with string keywords.\nThis is, also a test with multiple-key-wörds.") #duplicate with same hash
        search_writer.add_document(module="test", name="doc 2", keywords="THIS is a test with string keywords.\nThis is, also a test with multiple-key-wörds.")
        search_writer.add_document(module="test", name="doc 3", keywords="this is a test with string keywords.\nThis is, also a test with multiple-key-wörds.")

        self.assertEqual(self.search_engine.count_documents(), 0)
        self.assertEqual(self.search_engine.count_keywords(), 0)

        search_writer.commit()
        self.assertEqual(self.search_engine.count_documents(), 3)
        self.assertEqual(self.search_engine.count_keywords(), 11)

    def test_add_documents_with_list_keywords(self):
        search_writer = self.search_engine.writer()
        added_doc = search_writer.add_document(module="test", name="doc 1", keywords=search_writer.get_keywords_from_string("this is a TEST with list keywords.\nThis is, also a test with multiple-key-wörds."))
        self.assertEqual(added_doc.module, "test")
        self.assertEqual(added_doc.name, "doc 1")
        self.assertSetEqual(added_doc.keywords, set(["this", "is", "a", "test", "with", "list", "keywords", "also", "multiple", "key", "wörds"]))

        added_doc = search_writer.add_document(module="test", name="doc 1", keywords=set(["this", "is", "a", "test", "with", "list", "keywords", "also", "multiple", "key", "wörds"])) #duplicate with same hash
        search_writer.add_document(module="test", name="doc 2", keywords=["THIS", "is", "a", "test", "with", "list", "keywords", "also", "multiple", "key", "wörds"])
        search_writer.add_document(module="test", name="doc 3", keywords=["this", "is", "a", "test", "with", "list", "keywords", "also", "multiple", "key", "wörds"])

        self.assertEqual(self.search_engine.count_documents(), 0)
        self.assertEqual(self.search_engine.count_keywords(), 0)

        search_writer.commit()
        self.assertEqual(self.search_engine.count_documents(), 3)
        self.assertEqual(self.search_engine.count_keywords(), 11)
    
    def test_add_documents_and_cancel(self):
        search_writer = self.search_engine.writer()
        search_writer.add_document(module="test", name="doc 1", keywords="this is a test")
        search_writer.add_document(module="test", name="doc 2", keywords="this is a test")
        search_writer.add_document(module="test", name="doc 3", keywords="this is a test")

        self.assertEqual(self.search_engine.count_documents(), 0)
        self.assertEqual(self.search_engine.count_keywords(), 0)

        search_writer.cancel()
        self.assertEqual(self.search_engine.count_documents(), 0)
        self.assertEqual(self.search_engine.count_keywords(), 0)


class SearchCustomDocuments(unittest.TestCase):
    CustomDoc = namedtuple("CustomDoc", "id path keywords text")
    CustomDocWrong = namedtuple("CustomDocWrong", "id path text")

    def setUp(self):
        search.settings = dict()
    
    def tearDown(self):
        try:
            del search.search_engines["tests2"]
            self.search_engine = None
        except KeyError:
            pass

    def test_add_custom_documents(self):
        self.search_engine = search.get_search_engine("tests2", self.CustomDoc)
        search_writer = self.search_engine.writer()

        added_doc = search_writer.add_document(id="test", path="test2/doc 1", text="Lorem ipsum umlauts öäü #+-.,/1!?", keywords="test keywords 123")
        self.assertEqual(added_doc.id, "test")
        self.assertEqual(added_doc.path, "test2/doc 1")
        self.assertEqual(added_doc.text, "Lorem ipsum umlauts öäü #+-.,/1!?")
        self.assertEqual(added_doc.keywords, "test keywords 123")

        search_writer.add_document(id="test", path="test2/doc 1", text="Lorem ipsum umlauts öäü #+-.,/1!?", keywords="test keywords 123") #duplicate with same hash
        search_writer.add_document(id="test", path="test2/doc 2", text="Lorem ipsum umlauts öäü #+-.,/1!?", keywords=["test", "keywords", "123"])
        
        # add existing document instance
        doc = self.CustomDoc(id="test", path="test2/doc 3", text="Lorem ipsum umlauts öäü #+-.,/1!?", keywords="test keywords 123")
        added_doc = search_writer.add_document(doc)
        self.assertEqual(doc, added_doc)

        # add without keywords
        with self.assertRaises(ValueError):
            search_writer.add_document(id="test", path="test2/doc 4", text="Lorem ipsum umlauts öäü #+-.,/1!?")

        self.assertEqual(self.search_engine.count_documents(), 0)
        self.assertEqual(self.search_engine.count_keywords(), 0)

        search_writer.commit()
        self.assertEqual(self.search_engine.count_documents(), 3)
        self.assertEqual(self.search_engine.count_keywords(), 3)

    def test_add_wrong_custom_documents(self):
        with self.assertRaises(TypeError):
            self.search_engine = search.get_search_engine("tests2", self.CustomDocWrong)


class SearchSearcherTests(unittest.TestCase):
    def setUp(self):
        search.settings = dict()
        self.search_engine = search.get_search_engine("tests3")

        search_writer = self.search_engine.writer()
        self.documents = [
            search_writer.add_document(module="test1", name="doc 1.1", keywords="test1 doc lorem ipsum dolor"),
            search_writer.add_document(module="test1", name="doc 1.2", keywords="test1 doc ipsum umlautöäüßs"),
            search_writer.add_document(module="test2", name="doc 2.1", keywords="test2 doc lorem umlautöäüßs"),
            search_writer.add_document(module="test2", name="doc 2.2", keywords="test2 doc lorem dolor"),
            search_writer.add_document(module="test2", name="doc 2.3", keywords="test2 doc ipsum dolor"),
        ]
        search_writer.commit()
    
    def tearDown(self):
        del search.search_engines["tests3"]
        self.search_engine = None
    
    def test_search_exact_and(self):
        with self.search_engine.searcher() as searcher:
            result1 = searcher.search_exact("TEST2", True)
            self.assertEqual(len(result1), 3)
            self.assertListEqual(list(result1), get_list_by_indices(self.documents, 2,3,4))

            result2 = searcher.search_exact("lorem dolor", True)
            self.assertEqual(len(result2), 2)
            self.assertListEqual(list(result2), get_list_by_indices(self.documents, 0,3))

            result3 = searcher.search_exact("test", True)
            self.assertEqual(len(result3), 0)
            self.assertListEqual(list(result3), [])

            result4 = searcher.search_exact("umlautöäüßs", True)
            self.assertEqual(len(result4), 2)
            self.assertListEqual(list(result4), get_list_by_indices(self.documents, 1,2))

        self.assertEqual(self.search_engine.count_recent_searches(), 4)
        self.assertSequenceEqual(self.search_engine.get_recent_searches(), ["umlautöäüßs", "test", "lorem dolor", "TEST2"])
    
    def test_search_exact_or(self):
        with self.search_engine.searcher() as searcher:
            result1 = searcher.search_exact("TEST2", False)
            self.assertEqual(len(result1), 3)
            self.assertListEqual(list(result1), get_list_by_indices(self.documents, 2,3,4))

            result2 = searcher.search_exact("lorem dolor", False)
            self.assertEqual(len(result2), 4)
            self.assertListEqual(list(result2), [self.documents[0], self.documents[2], self.documents[3], self.documents[4]])

            result3 = searcher.search_exact("test", False)
            self.assertEqual(len(result3), 0)
            self.assertListEqual(list(result3), [])

            result4 = searcher.search_exact("umlautöäüßs", False)
            self.assertEqual(len(result4), 2)
            self.assertListEqual(list(result4), [self.documents[1], self.documents[2]])

        self.assertEqual(self.search_engine.count_recent_searches(), 4)
        self.assertSequenceEqual(self.search_engine.get_recent_searches(), ["umlautöäüßs", "test", "lorem dolor", "TEST2"])
    
    def test_search_wildcard_and(self):
        with self.search_engine.searcher() as searcher:
            result1 = searcher.search("TEST2", True)
            self.assertEqual(len(result1), 3)
            self.assertListEqual(list(result1), get_list_by_indices(self.documents, 2,3,4))

            result2 = searcher.search("orem olo", True)
            self.assertEqual(len(result2), 2)
            self.assertListEqual(list(result2), get_list_by_indices(self.documents, 0,3))

            result3 = searcher.search("test", True)
            self.assertEqual(len(result3), 5)
            self.assertListEqual(list(result3), self.documents)

            result4 = searcher.search("umlautöäüßs", True)
            self.assertEqual(len(result4), 2)
            self.assertListEqual(list(result4), get_list_by_indices(self.documents, 1,2))

        self.assertEqual(self.search_engine.count_recent_searches(), 4)
        self.assertSequenceEqual(self.search_engine.get_recent_searches(), ["umlautöäüßs", "test", "orem olo", "TEST2"])
    
    def test_search_wildcard_or(self):
        with self.search_engine.searcher() as searcher:
            result1 = searcher.search("TEST2", False)
            self.assertEqual(len(result1), 3)
            self.assertListEqual(list(result1), get_list_by_indices(self.documents, 2,3,4))

            result2 = searcher.search("orem olo", False)
            self.assertEqual(len(result2), 4)
            self.assertListEqual(list(result2), get_list_by_indices(self.documents, 0,2,3,4))

            result3 = searcher.search("test", False)
            self.assertEqual(len(result3), 5)
            self.assertListEqual(list(result3), self.documents)

            result4 = searcher.search("umlautöäüßs", False)
            self.assertEqual(len(result4), 2)
            self.assertListEqual(list(result4), get_list_by_indices(self.documents, 1,2))

        self.assertEqual(self.search_engine.count_recent_searches(), 4)
        self.assertSequenceEqual(self.search_engine.get_recent_searches(), ["umlautöäüßs", "test", "orem olo", "TEST2"])


class SearchResultTests(unittest.TestCase):
    def setUp(self):
        search.settings = dict()
        self.search_engine = search.get_search_engine("tests4")

        search_writer = self.search_engine.writer()
        self.documents = [
            search_writer.add_document(module="test1", name="doc 1.1", keywords="test1 doc lorem ipsum dolor"), #0
            search_writer.add_document(module="test2", name="doc 2.3", keywords="test2 doc ipsum dolor"),       #1
            search_writer.add_document(module="test2", name="doc 2.1", keywords="test2 doc lorem umlautöäüßs"), #2
            search_writer.add_document(module="test1", name="doc 1.2", keywords="test1 doc ipsum umlautöäüßs"), #3
            search_writer.add_document(module="test2", name="doc 2.2", keywords="test2 doc lorem dolor"),       #4
        ]
        search_writer.commit()
    
    def tearDown(self):
        del search.search_engines["tests4"]
        self.search_engine = None
    
    def test_ordering(self):
        with self.search_engine.searcher() as searcher:
            result1 = searcher.search_exact("doc")
            self.assertListEqual(list(result1), self.documents)
            self.assertListEqual(list(result1.reverse()), list(reversed(self.documents)))

            self.assertListEqual(list(result1.sortedby("name")), get_list_by_indices(self.documents, 0,3,2,4,1))
            self.assertListEqual(list(result1.sortedby("name", reverse=True)), get_list_by_indices(self.documents, 1,4,2,3,0))
            self.assertListEqual(list(reversed(result1.sortedby("name"))), get_list_by_indices(self.documents, 1,4,2,3,0))

            with self.assertRaises(AssertionError):
                result1.sortedby("name").reverse()
    
    def test_limit(self):
        with self.search_engine.searcher() as searcher:
            result1 = searcher.search_exact("doc")

            self.assertListEqual(list(result1.limit(3)), get_list_by_indices(self.documents, 0,1,2))
            self.assertListEqual(list(result1.limit(start=2, stop=4)), get_list_by_indices(self.documents, 2,3))
    
    def test_group(self):
        with self.search_engine.searcher() as searcher:
            result1 = searcher.search_exact("doc")
            self.assertListEqual(list(result1.sortedby("name").groupedby("module")), [("test1", get_list_by_indices(self.documents, 0,3)), ("test2", get_list_by_indices(self.documents, 2,4,1))])
            self.assertListEqual(list(result1.sortedby("name").groupedby("module", limit_groups=2)), [("test1", get_list_by_indices(self.documents, 0,3)), ("test2", get_list_by_indices(self.documents, 2,4))])
    
    def test_paginate(self):
        with self.search_engine.searcher() as searcher:
            result1 = searcher.search_exact("doc")
            self.assertListEqual(list(result1.paginate(2)), [get_list_by_indices(self.documents, 0,1), get_list_by_indices(self.documents, 2,3), get_list_by_indices(self.documents, 4)])