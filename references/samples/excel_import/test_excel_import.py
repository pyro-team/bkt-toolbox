# -*- coding: utf-8 -*-
'''
Created on 09.09.2014

@author: cschmitt
'''
from __future__ import print_function
import unittest
import clr
import os.path

clr.AddReference('Microsoft.Office.Interop.Excel')
from Microsoft.Office.Interop import Excel  # @UnresolvedImport

from bkt.library.excel import importer


import proc2dom_model
_model = proc2dom_model.create_model()


class ImportTest(unittest.TestCase):
    def setUp(self):
        self.model = _model 
        self.excel = Excel.ApplicationClass()
        
        
    def test_import(self):
        
        path = os.path.join(os.path.dirname(__file__), '2014-09-03 Funktionaler Footprint der Prozesse2.xlsx')
        workbook = self.excel.Workbooks.Open(path)
            
        imp = importer.ModelImporter(self.model, workbook)
        modeldata = imp.import_data()
        for dom in modeldata.domain:
            print(dom)
        for proc in modeldata.proc:
            print(proc)
        for p2d in modeldata.proc2dom:
            print(p2d)
            
        print(modeldata.domain.by_id['QM-06-02'])
        print(modeldata.domain.by_id['QM-06-02'].name)
        
        def domains_of_proc(proc):
            p2d = modeldata.proc2dom.select_proc_id(proc.id)
            domains = modeldata.domain.join_id(p2d, 'domain_id')
            return domains
            #return [modeldata.domain.by_id[p.domain_id] for p in p2d]
        
        proc11 = modeldata.proc.by_id[11]
        print('domains of proc 11')
        for dom in domains_of_proc(proc11):
            print(dom)
        
        #print('domains of proc 11 -- by expansion')
        #expansion = self.model.entities.proc.expansion
        #for dom in expansion.expand(modeldata, [proc11]):
        #    print(dom)
            
        print('domains of proc 11 -- by property')
        for dom in proc11.domains:
            print(dom)
            print(dom.parent)
            print(dom.procs)
            print('-------------')
            
        
        #fd.close()
        
    def tearDown(self):
        self.excel.Visible = True
        self.excel.Quit()
        
if __name__ == '__main__':
    unittest.main()