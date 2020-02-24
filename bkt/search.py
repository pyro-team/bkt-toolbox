# -*- coding: utf-8 -*-

### Search should be compatible to Whoosh

import shelve #for resources cache
import os.path #for resources cache
import logging

import bkt.helpers as _h

import re
import fnmatch #for fuzzy search
from collections import defaultdict, namedtuple, OrderedDict
from contextlib import contextmanager


### default search document
SearchDocument = namedtuple("SearchDocument", "module name keywords")


class SearchResults(object):
    def __init__(self):
        self._result_docs = []
    
    def __len__(self):
        return len(self._result_docs)
    
    def __iter__(self):
        for doc in self._result_docs:
            yield doc

    def groupedby(self, field):
        result_dict = OrderedDict()
        for doc in self._result_docs:
            try:
                result_dict[getattr(doc, field)].append(doc)
            except KeyError:
                result_dict[getattr(doc, field)] = [doc]
        return result_dict

    def sortedby(self, field):
        return sorted(self._result_docs, key=lambda d: getattr(d, field))
    
    def add_result(self, document):
        self._result_docs.append(document)


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
        self._engine.cache_sync()
        del self
    
    def cancel(self):
        del self


class SearchSearcher(object):
    def __init__(self, engine):
        self._engine = engine

    ### SEARCH ENGINE ###
    def search(self, query):
        ''' search with wildcards around each keyword, multiple keywords are connected with OR '''
        result = SearchResults()
        logging.debug("SEARCH: for "+query)

        # #check min length
        # if len(query) < 3:
        #     return result #empty result
        
        #split keyword by whitespace
        search_terms = set(["*{}*".format(s) for s in query.lower().split()])

        #perform search with wildcards
        result_keywords = set()
        for search_term in search_terms:
            result_keywords = result_keywords.union(fnmatch.filter(self._engine._keywords, search_term))
        
        logging.debug("SEARCH: found keywords "+",".join(result_keywords))

        #convert found keywords to document hashs
        result_doc_hashes = set()
        for keyword in result_keywords:
            result_doc_hashes = result_doc_hashes.union(self._engine._db[hash(keyword)])

        #add results in the order that documents have been added
        # for doc_hash in result_doc_hashes:
        #     result.add_result(self._engine._docs[doc_hash])
        for doc_hash,doc in self._engine._docs.iteritems():
            if doc_hash in result_doc_hashes:
                result.add_result(self._engine._docs[doc_hash])
        
        return result
    
    def search_exact(self, query):
        ''' exact search without wildcards, multiple keywords are connected with AND '''
        result = SearchResults()
        logging.debug("SEARCH EXACT: for "+query)
        
        #split keyword by whitespace
        search_terms = set([s for s in query.lower().split()])

        #add document hashes that have all defined keywords
        result_doc_hashes = self._engine._db[hash(search_terms.pop())]
        for keyword in search_terms:
            result_doc_hashes = result_doc_hashes.intersection(self._engine._db[hash(keyword)])

        #add results in the order that documents have been added
        for doc_hash,doc in self._engine._docs.iteritems():
            if doc_hash in result_doc_hashes:
                result.add_result(self._engine._docs[doc_hash])

        # for doc_hash in self._engine._db[hash(query.lower())]:
        #     result.add_result(self._engine._docs[doc_hash])
        return result


class SearchEngine(object):
    def __init__(self, name, schema):
        # cache_file = os.path.join( _h.get_cache_folder(), "search.%s.cache"%name )

        self._name = name
        self._schema = schema

        self._db = defaultdict(set)
        self._docs = OrderedDict()
        self._keywords = set()

        #load index from cache #:FIXME: caching doesnt work right now
        # self._cache = shelve.open(cache_file, protocol=2)
        # if "index" in self._cache:
        #     self._keywords = self._keywords.union(self._cache["keywords"])
        #     for k,v in self._cache["documents"].iteritems():
        #         self._docs[k].append(v)
        #     for k,v in self._cache["index"].iteritems():
        #         self._db[k].append(v)
    

    ### CACHE HANDLING ###
    def cache_sync(self):
        #:FIXME: caching doesnt work right now
        pass
        # self._cache["keywords"] = self._keywords
        # self._cache["documents"] = self._docs
        # self._cache["index"] = self._db
        # self._cache.sync()
    
    def cache_clear(self):
        #:FIXME: caching doesnt work right now
        pass
        # self._cache.clear()
        # self._cache.sync()

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