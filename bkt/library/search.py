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
# import fnmatch #for fuzzy search in list

from itertools import groupby, islice, chain
from collections import namedtuple, OrderedDict, deque
from contextlib import contextmanager

from bkt import settings

### default search document
SearchDocument = namedtuple("SearchDocument", "module name keywords")


class SearchResults(object):
    def __init__(self, doc_db, document_hashs=None):
        self._doc_db = doc_db
        self._result_hashes = document_hashs or set()

        self._reversed_order = False #reverse order of default iterator
        self._chained_iterators = None #chain iterators when sort, group etc. are called

    def __len__(self):
        return len(self._result_hashes)
    
    def __iter__(self):
        #self is an iterator, see next()
        self._iterator = iter(self._iterator)
        return self
    
    def __reversed__(self):
        #create list from normal forward iteration and then use reversed function
        return reversed(list(self))

    def next(self):
        try:
            return next(self._iterator)
        except StopIteration as e:
            #after iterator is exhausted, reset chained-iterators to start over again
            self._chained_iterators = None
            raise e

    __next__ = next #for python 3
    
    def _get_default_iterator(self):
        if self._reversed_order:
            iterator = reversed(self._doc_db) #OrderedDict supports reversed()
            #python3: iterator = reversed(self._doc_db.keys())
        else:
            iterator = self._doc_db.iterkeys()
        
        for doc_hash in iterator:
                if doc_hash in self._result_hashes:
                    yield self._doc_db[doc_hash]
    
    @property
    def _iterator(self):
        if self._chained_iterators is None:
            self._chained_iterators = self._get_default_iterator()
        return self._chained_iterators
    @_iterator.setter
    def _iterator(self, value):
        self._chained_iterators = value

    def groupedby(self, field, limit_groups=None):
        if limit_groups is not None:
            self._iterator = ( (k, list(islice(g, limit_groups))) for k, g in groupby(self._iterator, key=lambda d: getattr(d, field)) )
        else:
            self._iterator = ( (k, list(g)) for k, g in groupby(self._iterator, key=lambda d: getattr(d, field)) )
        
        #NOTE: the following line also works in "normal" situation, but len() on group and reversed(self) does not work!
        # self._iterator = groupby(self._iterator, key=lambda d: getattr(d, field))
        return self

    def sortedby(self, field, reverse=False):
        self._iterator = sorted(self._iterator, key=lambda d: getattr(d, field), reverse=reverse)
        return self
    
    def reverse(self):
        assert self._chained_iterators is None, "reverse() must be called before any other operation"
        #reverse order of dictionatory iteration in default iterator
        self._reversed_order = True
        return self

    def limit(self, stop, start=0):
        self._iterator = islice(self._iterator, stop=stop, start=start)
        return self

    def paginate(self, results_per_page):
        def chunk_iterator(iterable):
            iterable = iter(iterable)
            while True:
                result = list(islice(iterable, results_per_page))
                if result:
                    yield result
                else:
                    raise StopIteration
        
        self._iterator = chunk_iterator(self._iterator)
        return self


class SearchWriter(object):
    def __init__(self, engine):
        self._documents = []
        self._engine = engine

    ### WRITER ###
    def get_keywords_from_string(self, keywords):
        return set(re.findall(r'\w+', keywords.lower()))

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
                keywords = self.get_keywords_from_string(doc.keywords)
            else:
                keywords = doc.keywords
            
            self._engine._docs[doc_hash] = doc
            logging.debug("SEARCH: commit document {} for keywords: {}".format(doc_hash, keywords))
            for keyword in keywords:
                self._engine._keywords.add(keyword)
                self._engine._db.setdefault(hash(keyword), set()).add(doc_hash)

        del self
    
    def cancel(self):
        del self


class SearchSearcher(object):
    def __init__(self, engine):
        self._engine = engine

    ### SEARCH ENGINE ###
    def search(self, query, join_and=True):
        ''' search with wildcards around each keyword
            join_and=True > multiple keywords are connected with AND
            join_and=False > multiple keywords are connected with OR
        '''
        logging.debug("SEARCH: for "+query)

        #add to search history
        self._engine.add_to_recent(query)

        # #check min length
        # if len(query) < 3:
        #     return result #empty result
        
        #split keyword by whitespace
        # search_terms = set(["*{}*".format(s) for s in query.lower().split()])
        # search_patterns = [re.compile(".*"+s+".*") for s in query.lower().split()]
        search_terms = set(query.lower().split())

        #perform search with wildcards using fnmatch
        # result_keywords = []
        # for search_term in search_terms:
        #     result_keywords.append(
        #         fnmatch.filter(self._engine._keywords, search_term)
        #     )
        #     #result_keywords = result_keywords.union(fnmatch.filter(self._engine._keywords, search_term))
        
        #perform search using regex
        # result_keywords = {}
        # for keyword in self._engine._keywords:
        #     for pattern in search_patterns:
        #         if pattern.match(keyword):
        #             result_keywords.setdefault(pattern.pattern, []).append(keyword)

        #perform search using in operator
        # result_keywords = []
        # for search_term in search_terms:
        #     result_keywords.append([
        #         k for k in self._engine._keywords if search_term in k
        #     ])
        #fastest version with generator expression
        result_keywords = (
            [
                k for k in self._engine._keywords if search_term in k
            ] for search_term in search_terms
        )
        
        # logging.debug("SEARCH: found keywords "+",".join(chain.from_iterable(result_keywords)))

        #convert found keywords to document hashs
        #AND-search
        if join_and:
            #add document hashes that have all defined keywords
            try:
                keyword_doc_hashes = []
                for search_terms in result_keywords:
                    #search doc hashes for current keyword
                    current_doc_hashes = set()
                    for keyword in search_terms:
                        current_doc_hashes.update(self._engine._db[hash(keyword)])
                    keyword_doc_hashes.append(current_doc_hashes)
                #add doc hashes of all keyword doc hashes using intersection (AND-search)
                result_doc_hashes = set.intersection(*keyword_doc_hashes)
            except KeyError:
                logging.debug("SEARCH: empty result set due to KeyError")
                result_doc_hashes = set() #emptyset
        
        #OR-search
        else:
            #add document hashes that have any defined keywords
            result_doc_hashes = set()
            for keyword in chain.from_iterable(result_keywords): 
                #chain unpacks nesting of lists
                result_doc_hashes.update(self._engine._db.get(hash(keyword), set()))

        #add results in the order that documents have been added
        # for doc_hash,doc in self._engine._docs.iteritems():
        #     if doc_hash in result_doc_hashes:
        #         result.add_result(doc)

        return SearchResults(self._engine._docs, result_doc_hashes)
    
    def search_exact(self, query, join_and=True):
        ''' exact search without wildcards
            join_and=True > multiple keywords are connected with AND
            join_and=False > multiple keywords are connected with OR
        '''
        logging.debug("SEARCH EXACT: for "+query)

        #add to search history
        self._engine.add_to_recent(query)
        
        #split keyword by whitespace
        search_terms = set(query.lower().split())

        #AND-search
        if join_and:
            #add document hashes that have all defined keywords
            try:
                result_doc_hashes = self._engine._db[hash(search_terms.pop())]
                for keyword in search_terms:
                    result_doc_hashes.intersection_update(self._engine._db[hash(keyword)])
            except KeyError:
                logging.debug("SEARCH EXACT: empty result set due to KeyError")
                result_doc_hashes = set() #emptyset
        
        #OR-search
        else:
            #add document hashes that have any defined keywords
            result_doc_hashes = set()
            for keyword in search_terms:
                result_doc_hashes.update(self._engine._db.get(hash(keyword), set()))

        #add results in the order that documents have been added
        # for doc_hash,doc in self._engine._docs.iteritems():
        #     if doc_hash in result_doc_hashes:
        #         result.add_result(doc)

        return SearchResults(self._engine._docs, result_doc_hashes)


class SearchEngine(object):
    def __init__(self, name, schema):
        self._name = name
        self._schema = schema

        self._db = dict()
        self._docs = OrderedDict()
        self._keywords = set()

        self._settings_key = "bkt.search."+name
        self._recent_searches = deque(settings.get(self._settings_key+".recent_searches", []), maxlen=10)

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

    ### SEARCH HISTORY ###
    def add_to_recent(self, query):
        try:
            #try to remove if already exists and add to beginning
            self._recent_searches.remove(query)
            self._recent_searches.appendleft(query)
        except ValueError:
            self._recent_searches.appendleft(query)
        settings[self._settings_key+".recent_searches"] = self._recent_searches

    def get_recent_searches(self):
        return self._recent_searches

    def count_recent_searches(self):
        return len(self._recent_searches)


### FACTORY ###
search_engines = dict()

def get_search_engine(name, schema=SearchDocument):
    try:
        return search_engines[name]
    except:
        search_engines[name] = SearchEngine(name, schema)
        return search_engines[name]