# -*- coding: utf-8 -*-
'''
Simple search engine for indexing and searching for keywords

Created on 09.03.2020
@author: fstallmann
'''

from __future__ import absolute_import

### Search should be compatible to Whoosh
import logging

import re #for extracting keywords from string
import fnmatch #for fuzzy search in list

from collections import defaultdict, namedtuple, OrderedDict
from contextlib import contextmanager


### default search document
SearchDocument = namedtuple("SearchDocument", "module name keywords")


class SearchResults(object):
    def __init__(self, doc_db, document_hashs=None):
        self._doc_db = doc_db
        self._result_hashes = document_hashs or set()

        self._chained_iterators = None

    def __len__(self):
        return len(self._result_hashes)
    
    def __iter__(self):
        return self._iterator
    
    def _get_default_iterator(self):
        for doc_hash,doc in self._doc_db.iteritems():
                if doc_hash in self._result_hashes:
                    yield doc
    
    @property
    def _iterator(self):
        if self._chained_iterators is None:
            self._chained_iterators = self._get_default_iterator()
        return self._chained_iterators
    @_iterator.setter
    def _iterator(self, value):
        self._chained_iterators = value

    def groupedby(self, field):
        from itertools import groupby
        self._iterator = [(k,list(g)) for k, g in groupby(self._iterator, key=lambda d: getattr(d, field))]
        return self
        # result_dict = OrderedDict()
        # for doc in self._iterator:
        #     try:
        #         result_dict[getattr(doc, self._group_field)].append(doc)
        #     except KeyError:
        #         result_dict[getattr(doc, self._group_field)] = [doc]
        # return result_dict.iteritems()

    def sortedby(self, field, reverse=False):
        self._iterator = sorted(self._iterator, key=lambda d: getattr(d, field), reverse=reverse)
        return self

    def limit(self, stop, start=0):
        from itertools import islice
        self._iterator = islice(self._iterator, stop=stop, start=start)
        return self


class SearchWriter(object):
    def __init__(self, engine):
        self._documents = []
        self._engine = engine

    ### WRITER ###
    def add_document(self, **kwargs):
        if "keywords" not in kwargs:
            raise ValueError("Documents need a keywords field")
        doc = self._engine._schema(**kwargs)
        self._documents.append(doc)
    
    def commit(self):
        for doc in self._documents:
            try:
                doc_hash = hash(doc)
            except TypeError:
                #if doc is not hashable (e.g. contains a set), use string representation
                doc_hash = hash(str(doc))

            #do not create duplicates
            if doc_hash in self._engine._docs:
                logging.debug("SEARCH: duplicate doc hash "+str(doc_hash))
                continue

            #split comma-seperated values in list
            if type(doc.keywords) not in [list, set]:
                # keywords = doc.keywords.lower().replace(",", " ").split()
                keywords = re.findall(r'\w+', doc.keywords.lower())
            else:
                keywords = doc.keywords
            
            self._engine._docs[doc_hash] = doc
            logging.debug("SEARCH: commit document {} for keywords: {}".format(doc_hash, keywords))
            for keyword in keywords:
                self._engine._keywords.add(keyword)
                self._engine._db[hash(keyword)].add(doc_hash)

        del self
    
    def cancel(self):
        del self


class SearchSearcher(object):
    def __init__(self, engine):
        self._engine = engine

    ### SEARCH ENGINE ###
    def search(self, query):
        ''' search with wildcards around each keyword, multiple keywords are connected with OR '''
        logging.debug("SEARCH: for "+query)

        # #check min length
        # if len(query) < 3:
        #     return result #empty result
        
        #split keyword by whitespace
        search_terms = set(["*{}*".format(s) for s in query.lower().split()])

        #perform search with wildcards using fnmatch
        #FIXME: consider using regex instead of fnmatch?
        result_keywords = set()
        for search_term in search_terms:
            result_keywords = result_keywords.union(fnmatch.filter(self._engine._keywords, search_term))
        
        logging.debug("SEARCH: found keywords "+",".join(result_keywords))

        #convert found keywords to document hashs
        result_doc_hashes = set()
        for keyword in result_keywords:
            # result_doc_hashes = result_doc_hashes.union(self._engine._db[hash(keyword)])
            result_doc_hashes.update(self._engine._db[hash(keyword)])

        #add results in the order that documents have been added
        # for doc_hash,doc in self._engine._docs.iteritems():
        #     if doc_hash in result_doc_hashes:
        #         result.add_result(doc)
        
        return SearchResults(self._engine._docs, result_doc_hashes)
    
    def search_exact(self, query):
        ''' exact search without wildcards, multiple keywords are connected with AND '''
        logging.debug("SEARCH EXACT: for "+query)
        
        #split keyword by whitespace
        search_terms = set([s for s in query.lower().split()])

        #add document hashes that have all defined keywords
        #NOTE: as _db is a defaultdict, unkown search_terms do not throw an error but return an empty set
        result_doc_hashes = self._engine._db[hash(search_terms.pop())]
        for keyword in search_terms:
            # result_doc_hashes = result_doc_hashes.intersection(self._engine._db[hash(keyword)])
            result_doc_hashes.intersection_update(self._engine._db[hash(keyword)])

        #add results in the order that documents have been added
        # for doc_hash,doc in self._engine._docs.iteritems():
        #     if doc_hash in result_doc_hashes:
        #         result.add_result(doc)

        # for doc_hash in self._engine._db[hash(query.lower())]:
        #     result.add_result(self._engine._docs[doc_hash])
        return SearchResults(self._engine._docs, result_doc_hashes)


class SearchEngine(object):
    def __init__(self, name, schema):
        self._name = name
        self._schema = schema

        self._db = defaultdict(set)
        self._docs = OrderedDict()
        self._keywords = set()

    ### INDEXING AND SEARCHING ###
    def count_documents(self):
        return len(self._docs)

    def count_keywords(self):
        return len(self._keywords)

    def writer(self):
        return SearchWriter(self)
    
    @contextmanager
    def searcher(self):
        try:
            yield SearchSearcher(self)
        finally:
            pass


### FACTORY ###
search_engines = dict()

def get_search_engine(name, schema=SearchDocument):
    try:
        return search_engines[name]
    except:
        search_engines[name] = SearchEngine(name, schema)
        return search_engines[name]