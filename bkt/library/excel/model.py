'''
Created on 08.09.2014

@author: cschmitt
'''

from __future__ import absolute_import

from collections import OrderedDict

class TemporaryBuildObject(object):
    pass

class AttributeContainer(object):
    def __init__(self, dictionary=None):
        if dictionary is None:
            self._elems = OrderedDict()
        else:
            self._elems = OrderedDict(dictionary)
    
    def __getattr__(self, attr):
        try:
            return self._elems[attr]
        except KeyError:
            raise AttributeError(attr)
    
    def __getitem__(self, key):
        return self._elems[key]
    
    def __iter__(self):
        return self._elems.values()
    
class Model(object):
    def __init__(self):
        self.entities = None
        
class BaseType(object):
    _entity = None
    
    def __init__(self, **kwargs):
        entity = type(self)._entity
        for col in entity.columns:
            setattr(self, col.name, None)
        for attr, value in kwargs.iteritems():
            col = entity.col_by_name.get(attr)
            if col is None:
                raise TypeError('unknown argument %s for entity %s' % (attr, entity.name))
            setattr(self, col.name, value)
            
    def __repr__(self):
        entity = type(self)._entity
        content = ', '.join([('%s=%r' % (c.name, getattr(self, c.name))) for c in entity.columns])
        return '<%s: %s> ' % (entity.name, content)
        

class Entity(object):
    def __init__(self, entity_name, excel_name, columns):
        self.name = entity_name
        self.excel_name = excel_name
        self.columns = columns
        self.col_by_name = {c.name:c for c in columns}
        self.type = type(self.name, (BaseType,), dict(_entity=self))
        
    def __call__(self, **kwargs):
        return self.type(**kwargs)

class EntityBuilder(object):
    def __init__(self, builder, entity_name, excel_name):
        self.builder = builder
        self.name = entity_name
        self.excel_name = excel_name
        self.temp = TemporaryBuildObject()
        
    def __enter__(self):
        return self.temp
    
    def __exit__(self, exc_type, value, traceback):
        cols = []
        for attr, value in self.temp.__dict__.iteritems():
            if isinstance(value, Column):
                value.name = attr
                cols.append(value)
        cols.sort(key=lambda c : c.order)
        entity = Entity(self.name, self.excel_name, cols)
        self.builder.entities[self.name] = entity

class ExpansionBuilder(object):
    def __init__(self, model_builder, name):
        self.model_builder = model_builder
        self.name = name
        self._forward = None
        self._backward = None
        self._forward_unique = False
        self._backward_unique = False
        self.joins = []
        
    def forward(self, attr, unique=False):
        self._forward = attr
        self._forward_unique = unique
        
    def backward(self, attr, unique=False):
        self._backward = attr
        self._backward_unique = unique
        
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, value, traceback):
        self.build_expansion()
        
    def split_attr2(self, s):
        parts = s.split('.')
        if len(parts) != 2:
            raise ValueError(s)
        
        entity_name, col_name = parts
        entity = self.model_builder.entities[entity_name]
        if col_name in entity.col_by_name:
            raise ValueError('expansion attribute %s already present' % col_name)
        return entity, col_name
    
    def split_attr(self, s):
        parts = s.split('.')
        if len(parts) != 2:
            raise ValueError(s)
        
        entity_name, col_name = parts
        entity = self.model_builder.entities[entity_name]
        column = entity.col_by_name[col_name]
        return entity, column
    
    def join(self, source, target):
        s = self.split_attr(source)
        t = self.split_attr(target)
        self.joins.append((s, t))
    
    def build_expansion(self):
        self.source, self.source_attr = self.split_attr2(self._forward)
        self.target, self.target_attr = self.split_attr2(self._backward)
        self.check_entities()

        fwd = Expansion(self.source, self.target, self.joins, self._forward_unique)
        setattr(self.source.type, self.source_attr, fwd.as_property())
        
        back = fwd.reverse()
        back.unique = self._backward_unique
        setattr(self.target.type, self.target_attr, back.as_property())
        
    def check_entities(self):
        current = self.source
        for (s,_),(t,_) in self.joins:
            if current is not s:
                raise ValueError
            current = t
        if current is not self.target:
            raise ValueError
        
class Expansion(object):
    def __init__(self, source, target, joins, unique):
        self.source = source
        self.target = target
        self.joins = joins
        self.unique = unique
        
    def reverse(self):
        rj = []
        for src, dst in reversed(self.joins):
            rj.append((dst, src))
        return Expansion(self.target, self.source, rj, False)
        
    def expand(self, modeldata, objects):
        result_set = list(objects)
        
        for o in result_set:
            if not isinstance(o, self.source.type):
                raise TypeError('expected %r, got %r' % (self.source.type, o))
        
        #print(result_set)
        for (src, src_col), (dst, dst_col) in self.joins:
            joiner = modeldata[dst.name].join(dst_col.name)
            next_set = joiner(result_set, src_col.name)
            result_set = next_set
            #print(result_set)
        
        if self.unique:
            if len(result_set) == 0:
                return None
            elif len(result_set) == 1:
                return result_set[0]
            else:
                raise ValueError('expansion not unique')
        else:
            return result_set
    
    def as_property(self):
        def getter(obj):
            return self.expand(obj._modeldata, [obj])
        return property(getter)

class ModelBuilder(object):
    def __init__(self):
        self.entities = OrderedDict()
        self.expansions = OrderedDict()
    
    def define_entity(self, entity_name, excel_name):
        return EntityBuilder(self, entity_name, excel_name)
    
    def define_expansion(self, expansion_name):
        return ExpansionBuilder(self, expansion_name)
    
    def get_model(self):
        model = Model()
        model.entities = AttributeContainer(self.entities)
        return model

S_VALUE = 1
S_TEXT = 2
S_FORMULA = 3
S_FORMULA_LOCAL = 4

class Column(object):
    _declaration_order = 0
    
    def __init__(self, excel_name, source=S_TEXT, converter=None, python_name=None, convert_none=False, empty_string_as_none=True, skip_none=False):
        self.name = python_name
        self.excel_name = excel_name
        self.source = S_TEXT
        self.converter = converter
        self.convert_none = convert_none
        self.empty_string_as_none = empty_string_as_none
        self.skip_none = skip_none

        self.order = self._declaration_order
        Column._declaration_order += 1
        
    def get_content(self, cell):
        if self.source == S_VALUE:
            val = cell.Value
        elif self.source == S_TEXT:
            val = cell.Text
        elif self.source == S_FORMULA:
            val = cell.Formula
        elif self.source == S_FORMULA_LOCAL:
            val = cell.FormulaLocal
        
        try:
            if self.empty_string_as_none and val == '':
                val = None
            if self.converter is not None:
                if self.convert_none or val is not None:
                    val = self.converter(val)
        except:
            print('ERROR: could not convert %r with %r' % (val, self.converter))
            raise
            
        return val

class Selector(object):
    def __init__(self, entity_data, attr):
        self.entity_data = entity_data
        self.attr = attr
    
    def __call__(self, value):
        index = self.entity_data.index(self.attr)
        if index.unique:
            val = index.storage.get(value)
            if val is None:
                return []
            else:
                return [val]
        else:
            return [o for o in self.entity_data if getattr(o, self.attr) == value]
        
class Joiner(object):
    def __init__(self, entity_data, attr):
        self.entity_data = entity_data
        self.attr = attr
        
    def __call__(self, other_objects, other_attr):
        result = []
        index = self.entity_data.index(self.attr)
        if not index.unique:
            for other_obj in other_objects:
                for obj in self.entity_data:
                    if getattr(obj, self.attr) == getattr(other_obj, other_attr):
                        result.append(obj)
        else:
            for obj in other_objects:
                target = index.storage.get(getattr(obj, other_attr))
                if target is not None:
                    result.append(target)
        return result
    
class Index(object):
    def __init__(self):
        self.unique = True
        self.storage = {}
        
    def __getitem__(self, value):
        return self.storage[value]

class IndexNotUniqueError(Exception):
    pass

class VoidStorage(object):
    def __getitem__(self, key):
        raise IndexNotUniqueError

    def get(self, key):
        raise IndexNotUniqueError
     
class EntityData(object):
    def __init__(self, entity, objects):
        self.entity = entity
        self.objects = objects
        self.indices = {}
        
    def __iter__(self):
        return self.objects
    
    def select(self, select_attr):
        if not select_attr in self.entity.col_by_name:
            raise AttributeError
        return Selector(self, select_attr)
    
    def join(self, join_attr):
        if not join_attr in self.entity.col_by_name:
            raise AttributeError(join_attr)
        return Joiner(self, join_attr)
    
    def index(self, by_attr):
        if not by_attr in self.entity.col_by_name:
            raise AttributeError
        
        index = self.indices.get(by_attr)
        if index is None:
            index = Index()
            storage = index.storage
            for obj in self:
                key = getattr(obj, by_attr)
                if key in storage:
                    print('WARNING: index for %s on %s is not unique: multiple occurrence of %r' % (by_attr, self.entity.name, key))
                    index.unique = False
                    index.storage = VoidStorage()
                    break
                storage[key] = obj
            self.indices[by_attr] = index
        
        return index
    
    def __getattr__(self, attr):
        if attr.startswith('select_'):
            select_attr = attr[7:]
            return self.select(select_attr)
        elif attr.startswith('by_'):
            by_attr = attr[3:]
            return self.index(by_attr)
        elif attr.startswith('join_'):
            join_attr = attr[5:]
            return self.join(join_attr)
        else:
            raise AttributeError(attr)

class ModelData(object):
    def __init__(self, model_data):
        self._model_data = model_data
        for edata in model_data.values():
            for obj in edata:
                obj._modeldata = self

    def __getitem__(self, entity_name):
        return self._model_data[entity_name]
        
    def __getattr__(self, entity_name):
        try:
            return self._model_data[entity_name]
        except KeyError:
            raise AttributeError(entity_name)