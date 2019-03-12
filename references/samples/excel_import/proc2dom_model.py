# -*- coding: utf-8 -*-
'''
Created on 09.09.2014

@author: cschmitt
'''
from bkt.library.excel.model import ModelBuilder, Column, S_TEXT

def create_model():
    builder = ModelBuilder()
    
    with builder.define_entity("proc", "Processes") as proc:
        proc.id = Column("ID", S_TEXT, int, skip_none=True)
        proc.name = Column("Prozess", S_TEXT)
    
    with builder.define_entity("proc2dom", "Processes2Domains") as p2d:
        p2d.id = Column("ID", S_TEXT, int, skip_none=True)
        p2d.proc_id = Column("ProzessID", S_TEXT, int, skip_none=True)
        p2d.used_function = Column("genutzte Funktion", S_TEXT)
        p2d.domain_id = Column(u"ID Domäne", S_TEXT, skip_none=True)
    
    with builder.define_entity("domain", "DomainStructure") as dom:
        dom.id = Column("Domain ID", S_TEXT, skip_none=True)
        dom.name = Column("Domain Name", S_TEXT)
        dom.parent_id = Column("Parent ID", S_TEXT)
        dom.aggregate_id = Column("Aggregate Domain ID", S_TEXT)
        
    with builder.define_expansion('foo') as expansion:
        expansion.forward('proc.domains')
        expansion.backward('domain.procs')

        expansion.join('proc.id', 'proc2dom.proc_id')
        expansion.join('proc2dom.domain_id', 'domain.id')
        
    with builder.define_expansion('foo2') as expansion:
        expansion.forward('domain.parent', unique=True)
        expansion.backward('domain.children')

        expansion.join('domain.parent_id', 'domain.id')
    
    return builder.get_model()

if __name__ == '__main__':
    create_model()