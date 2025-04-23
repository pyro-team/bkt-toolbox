'''
Created on 09.09.2014

@author: cschmitt
'''



from collections import OrderedDict

from bkt.library.excel.model import ModelData, EntityData

def import_all_sheets(workbook):
    sheets = OrderedDict()
    for sheet in workbook.Sheets:
        sheets[sheet.Name] = import_table_text(sheet)
    return sheets
    
def import_table_text(sheet):
    def iter_rows():
        for row in sheet.UsedRange.Rows:
            _row = []
            for i in range(sheet.UsedRange.Columns.Count):
                cell = row.Cells(i+1)
                _row.append(cell.Text)
            yield _row
    return list(iter_rows())

def import_table_values(sheet):
    def iter_rows():
        for row in sheet.UsedRange.Rows:
            _row = []
            for i in range(sheet.UsedRange.Columns.Count):
                cell = row.Cells(i+1)
                _row.append(cell.Value())
            yield _row
    return list(iter_rows())

def iter_rows_as_text(sheet):
    for row in sheet.UsedRange.Rows:
        _row = []
        for i in range(sheet.UsedRange.Columns.Count):
            cell = row.Cells(i+1)
            _row.append(cell.Text)
        yield _row
        
class Header(object):
    def __init__(self, row_index, col_index_by_name):
        self.row_index = row_index
        self.col_index_by_name = col_index_by_name

class ModelImporter(object):
    def __init__(self, model, workbook):
        self.model = model
        self.workbook = workbook
        
    def search_header(self, entity, sheet):
        header_items = [c.excel_name for c in entity.columns]
        #print(header_items)
        
        def contains_all_items(table_row):
            #print(table_row)
            for h in header_items:
                if not h in table_row:
                    return False
            return True
                    
        for j, row in enumerate(iter_rows_as_text(sheet)):
            if contains_all_items(row):
                return Header(j, {h:row.index(h) for h in header_items})
            
    def import_data(self):
        mdata = {}
        for entity in self.model.entities:
            objects = self.import_entity(entity)
            edata = EntityData(entity, objects)
            mdata[entity.name] = edata
        return ModelData(mdata)
    
    def import_entity(self, entity):
        sheet = self.workbook.Sheets[entity.excel_name]
        header = self.search_header(entity, sheet)
        if header is None:
            raise ValueError('header for %s in sheet %s not found' % (entity.name, entity.excel_name))
        
        imported_objects = []
        for j, row in enumerate(sheet.UsedRange.Rows):
            if j <= header.row_index:
                continue
            kwargs = {}
            skip = False
            for column in entity.columns:
                col_index = header.col_index_by_name[column.excel_name]
                cell = row.Cells(col_index+1)
                value = column.get_content(cell)
                if value is None and column.skip_none:
                    skip = True
                kwargs[column.name] = value
            
            if not skip:
                imported_objects.append(entity(**kwargs))
            
        return imported_objects
            