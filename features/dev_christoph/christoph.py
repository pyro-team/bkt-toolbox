# -*- coding: utf-8 -*-
'''
Created on 21.01.2014

@author: cschmitt
'''

import json
import math

import bkt
from bkt import mso
from bkt.library.algorithms import get_bounds_shapes, median
from bkt.library.table import TableRecognition
#from bkt.ribbon import mso, Group, ForeignControl
from bkt import helpers
import traceback

# FIXME: benötigter Namespace zur Referenzierung fremder Controls unklar (ProgID vs. URI)
# register_namespace('bktdev', 'BKT.Dev.DevAddIn')
#register_namespace('bktdev', 'http://www.business-kasper-toolbox.com/toolbox/dev')

class TagDescriptor(object):
    def __init__(self,name,convert_json=None,convert_back=None):
        self.name = name
        self.convert_json = convert_json
        self.convert_back = convert_back
        
    def __get__(self,obj,cls=None):
        if obj is None:
            return self
        val = obj.data.get(self.name)
        if self.convert_back is not None:
            val = self.convert_back(val)
        return val
    
    def __set__(self,obj,value):
        if self.convert_json:
            value = self.convert_json(value)
        obj.data[self.name] = value

def add_properties(cls):
    def get_conversion(conversion_name):
        if conversion_name == 'float':
            return (float,None)
        elif conversion_name == None:
            return (None,None)
        else:
            raise AssertionError
    
    for name in cls.PROPERTIES:
        if '|' in name:
            conversion, name = name.split('|')
        else:
            conversion = None
        descriptor = TagDescriptor(name, *get_conversion(conversion))
        setattr(cls, name, descriptor)
    return cls


class BKTTag(object):
    TAG = "BKT"

    def __init__(self, tags):
        self.tags = tags
        self.data = {}
        
    def load(self):
        try:
            tag_data = self.tags.Item(self.TAG)
            if not tag_data:
                self.data = {}
            self.data = json.loads(tag_data)
        except:
            self.data = {}
        
    def save(self):
        try:
            tag_data = json.dumps(self.data)
            self.tags.Add(self.TAG,tag_data)
        except Exception:
            helpers.exception_as_message()
        
    def __enter__(self):
        self.load()
        return self

    def __exit__(self, cls, value, traceback):
        self.save()

@add_properties
class BKTPresentationTag(BKTTag):
    PROPERTIES = ['float|contentarea_left',
                  'float|contentarea_top',
                  'float|contentarea_width',
                  'float|contentarea_height',
                  'float|default_spacing']
    
    @property
    def is_area_set(self):
        return self.contentarea_left is not None

class PPTContainer(bkt.FeatureContainer):
    def presentation_tag(self, presentation):
        return BKTPresentationTag(presentation.Tags)

@bkt.configure(label='Table Operations')
@bkt.group
class TableGroup(PPTContainer):
    
    #@box(vertical=True,name='box1')    
    @bkt.arg_shapes_limited(2)
    @bkt.image('align table 32')
    @bkt.large_button("Align Table (Default Spacing)")
    def dev_align_table(self, shapes):
        tr = TableRecognition(shapes)
        tr.run()
        tr.align()

    @bkt.arg_shapes_limited(2)
    @bkt.image('align table median 32')
    @bkt.large_button("Align Table (Median Spacing)")
    def dev_align_table_median(self, shapes):
        tr = TableRecognition(shapes)
        tr.run()
        tr.align(tr.median_spacing())

    @bkt.arg_shapes_limited(2)
    @bkt.spinner_box
    def dev_align_table_set_spacing(self, shapes, value):
        try:
            v = float(value)
        except ValueError:
            traceback.print_exc()
            return
        
        tr = TableRecognition(shapes)
        tr.run()
        tr.align(v)
        
    @bkt.arg_shapes_limited(2)
    @dev_align_table_set_spacing.get_text
    def dev_align_table_set_spacing_get_text(self, shapes):
        tr = TableRecognition(shapes)
        tr.run()
        res = tr.median_spacing()
        print res
        return str(res)
    
    @bkt.arg_shapes_limited(2)
    @dev_align_table_set_spacing.increment
    def dev_inc_spacing(self, shapes):
        self.dev_align_table_median_increase(shapes)

    @bkt.arg_shapes_limited(2)
    @dev_align_table_set_spacing.decrement
    def dev_dec_spacing(self, shapes):
        self.dev_align_table_median_decrease(shapes)

    @bkt.arg_shapes_limited(2)
    @bkt.image('Zero Spacing 32')
    @bkt.large_button("Align Zero")
    def dev_align_table_zero(self,shapes):
        tr = TableRecognition(shapes)
        tr.run()
        tr.align(0)
        
    #@bkt.arg_shapes_limited(2)
    #@bkt.image('Decrease Spacing')
    #@bkt.large_button("Decrease Spacing")
    def dev_align_table_median_decrease(self, shapes):
        tr = TableRecognition(shapes)
        tr.run()
        tr.align(tr.median_spacing()-1)

    #@bkt.arg_shapes_limited(2)
    #@bkt.image('Increase Spacing')
    #@bkt.large_button("Increase Spacing")
    def dev_align_table_median_increase(self, shapes):
        tr = TableRecognition(shapes)
        tr.run()
        tr.align(tr.median_spacing()+1)
        
    @bkt.arg_presentation
    @bkt.arg_shapes_limited(2)
    @bkt.image('Fit content area')
    @bkt.large_button("Fit Table To Content Area")
    def dev_fit_table_to_content(self, shapes, presentation):
        with self.presentation_tag(presentation) as tag:
            if tag.is_area_set:
                left = tag.contentarea_left
                top = tag.contentarea_top
                width = tag.contentarea_width
                height = tag.contentarea_height
            else:
                left = 0
                top = 0
                setup = presentation.PageSetup
                width = setup.SlideWidth
                height = setup.SlideHeight
        
        tr = TableRecognition(shapes)
        tr.run()
        spacing = tr.median_spacing()
        tr.fit_content(left, top, width, height, spacing)

@bkt.configure(label='In-Place Table Operations')
@bkt.group
class TableInPlaceGroup(PPTContainer):
    @bkt.arg_shapes_limited(2)
    @bkt.image('Fit Cells 32')
    @bkt.large_button('Fit Cells')
    def dev_fit_cells_inplace(self,shapes):
        tr = TableRecognition(shapes)
        tr.run()
        spacing = tr.median_spacing()
        bounds = tr.get_bounds()
        tr.fit_content(*bounds, spacing=spacing, fit_cells=True)

    @bkt.arg_shapes_limited(2)
    @bkt.image('Increase InPlace Spacing')
    @bkt.large_button('Increase Spacing')
    def dev_increase_spacing(self,shapes):
        tr = TableRecognition(shapes)
        tr.run()
        tr.change_spacing_in_bounds(1)
        
    @bkt.arg_shapes_limited(2)
    @bkt.image('Decrease InPlace Spacing')
    @bkt.large_button('Decrease Spacing')
    def dev_decrease_spacing(self,shapes):
        tr = TableRecognition(shapes)
        tr.run()
        tr.change_spacing_in_bounds(-1)

    @bkt.arg_shapes_limited(2)
    @bkt.image('Tabelle transponieren 32')
    @bkt.large_button('Transpose')
    def dev_transpose_table(self,shapes):
        tr = TableRecognition(shapes)
        tr.run()
        spacing = tr.median_spacing()
        left, top, width, height = tr.get_bounds()
        tr.transpose()
        tr.transpose_cell_size();
        tr.fit_content(left, top, width, height, spacing)
    
@bkt.configure(label='Misc', uuid='54e10f83-b661-4f1a-b4d7-e2e98a14e9db')
@bkt.group
class MiscGroup(PPTContainer):
    
    @bkt.arg_shapes_limited(2)
    @bkt.image('Info')
    @bkt.large_button('Table Info')
    def dev_table_info(self,shapes):
        tr = TableRecognition(shapes)
        tr.run()
        msg = u""
        msg += "dimension: rows=%d, cols=%d\r\n" % tr.dimension
        msg += "median spacing: %r\r\n" % tr.median_spacing()
             
        helpers.log(msg)
    
    @bkt.arg_presentation
    @bkt.arg_shapes_limited(1, 1)
    @bkt.image('Set content area')
    @bkt.large_button('Set Content Area')
    def dev_set_content_area(self, presentation, shapes):
        shape = shapes[0]
        with self.presentation_tag(presentation) as tag:
            tag.contentarea_left = shape.Left
            tag.contentarea_top = shape.Top
            tag.contentarea_width = shape.Width
            tag.contentarea_height = shape.Height
        shape.Delete()
        
    @bkt.arg_shapes_limited(2)
    @bkt.image('Fit Cells 32')
    @bkt.large_button('Fit Cells (destructive)')
    def dev_fit_cells(self,shapes):
        tr = TableRecognition(shapes)
        tr.run()
        tr.align(spacing=tr.median_spacing(), fit_cells=True)
    
    
    @bkt.arg_shapes_limited(2,2)        
    @bkt.image('Swap Shapes 32')
    @bkt.large_button('Swap Shapes')
    def devswap(self, shapes):
        print "Test"
        s1, s2 = shapes
        s1.Left, s2.Left = s2.Left, s1.Left
        s1.Top, s2.Top = s2.Top, s1.Top
    
    @bkt.arg_shapes_limited(2,2)        
    @bkt.image('Swap Cells 32')
    @bkt.large_button('Swap Cells 2')
    def devswap2(self, shapes):
        s1, s2 = shapes
        s1.Left, s2.Left = s2.Left, s1.Left
        s1.Top, s2.Top = s2.Top, s1.Top
        s1.Width, s2.Width = s2.Width, s1.Width
        s1.Height, s2.Height = s2.Height, s1.Height
        
    @bkt.arg_shapes_limited(2)
    @bkt.image('Equalize Cells 32')
    @bkt.large_button('Equalize With Last Selected')
    def equalize(self, shapes):
        ref = shapes[-1]
        for s in shapes[:-1]:
            s.Width = ref.Width
            s.Height = ref.Height

    @bkt.arg_shapes_limited(2)
    @bkt.image('Fit height 32')
    @bkt.large_button('Equalize Height')
    def equalize_height(self, shapes):
        ref = shapes[-1]
        for s in shapes[:-1]:
            s.Height = ref.Height

    @bkt.arg_shapes_limited(2)
    @bkt.image('Fit width 32')
    @bkt.large_button('Equalize Width')
    def equalize_width(self, shapes):
        ref = shapes[-1]
        for s in shapes[:-1]:
            s.Width = ref.Width
            
#@group(label='Misc')
@bkt.group
class TestGroup(PPTContainer):

    @bkt.arg_shapes_limited(1)        
    @bkt.image_mso('HappyFace')
    @bkt.large_button('Align/Resize on Raster')
    def align_on_grid(self, shapes):
        raster = 4
        
        def r(dim):
            return round(float(dim)/float(raster))*raster
        
        for s in shapes:
            for attr in ('Top', 'Left', 'Width', 'Height'):
                val = getattr(s, attr)
                setattr(s, attr, r(val))
            
    @bkt.arg_shapes_limited(1)        
    @bkt.image_mso('HappyFace')
    @bkt.large_button('vstack')
    def vstack(self, shapes):
        shapes = sorted(shapes, key=lambda s : s.Top)
        y = shapes[0].Top
        for s in shapes:
            s.Top = y
            y += s.Height
    
    picked_color = None
    
    @bkt.arg_shapes_limited(1)        
    @bkt.image_mso('HappyFace')
    @bkt.large_button('pick color')
    def pick_color(self, shapes):
        self.picked_color = shapes[0].Fill.ForeColor.RGB
        
    @bkt.arg_shapes        
    @bkt.image_mso('HappyFace')
    @bkt.large_button('apply color')
    def apply_color(self, shapes):
        for shape in shapes:
            shape.Fill.ForeColor.RGB = self.picked_color
            
    def get_center_and_radius(self, shapes):
        midpoints = [(s.Left + 0.5*s.Width, s.Top + 0.5*s.Height) for s in shapes]
        N = float(len(midpoints))
        #x,y,w,h = get_bounds_shapes(midpoints)
        
        cx = sum(p[0] for p in midpoints) / N
        cy = sum(p[1] for p in midpoints) / N
        
        def midpoint_dist(p):
            x, y = p
            return math.hypot(cx-x, cy-y)
        
        r = sum(midpoint_dist(p) for p in midpoints) / N
        return cx, cy, r
    
    def align_circle(self, shapes, cx, cy, r):
        def arg(s):
            scx = s.Left + 0.5*s.Width
            scy = s.Top + 0.5*s.Height
            s_arg = math.atan2(scy-cy, scx-cx)
            if s_arg < 0:
                s_arg += math.pi*2
            return s_arg
        
        shapes = sorted(shapes, key=arg)
        
        n = len(shapes)
        omega = (2*math.pi)/float(n)
        
        if n%2 == 1:
            alpha0 = -omega/4
        else:
            alpha0 = 0
            
        for i, shape in enumerate(shapes):
            alpha = alpha0 + i*omega
            x = cx + r*math.cos(alpha)
            y = cy + r*math.sin(alpha)
            
            shape.Left = x - 0.5*shape.Width
            shape.Top = y - 0.5*shape.Height

    @bkt.arg_shapes_limited(2)
    @bkt.image('Test32')
    @bkt.large_button('Circle')
    def circle(self, shapes):
        cx, cy, r = self.get_center_and_radius(shapes)
        self.align_circle(shapes, cx, cy, r)

    @bkt.arg_shapes_limited(2)
    @bkt.image('Test32')
    @bkt.large_button('Decrease Radius')
    def decrease_radius(self, shapes):
        cx, cy, r = self.get_center_and_radius(shapes)
        r -= 5
        self.align_circle(shapes, cx, cy, r)

    @bkt.arg_shapes_limited(2)
    @bkt.image('Test32')
    @bkt.large_button('Increase Radius')
    def increase_radius(self, shapes):
        cx, cy, r = self.get_center_and_radius(shapes)
        r += 5
        self.align_circle(shapes, cx, cy, r)

@bkt.configure(label='Test Group 2', uuid='5e2fdb4a-d139-4f34-bd46-57b1e404cf4d')
@bkt.group
class TestGroup2(PPTContainer):
    @bkt.image_mso('HappyFace')
    @bkt.large_button('Split Bullets')
    @bkt.arg_context
    @bkt.arg_shapes
    def split_bullets(self, context, shapes):
        
        def par_height(par, with_spaces=True):
            h = par.Lines().Count * par.Font.Size * (par.ParagraphFormat.SpaceWithin + 0.2)
            if with_spaces:
                h += par.ParagraphFormat.SpaceBefore + par.ParagraphFormat.SpaceAfter
            return h
        
        def trim(par):
            while par.Characters(par.Length, 1).Text in '\r\n':
                par.Characters(par.Length, 1).Delete()

        for shp in shapes:
            if not shp.TextFrame.TextRange.Text:
                continue
            
            shp.Select(True)
            
            for parIndex in range(2, shp.TextFrame.TextRange.Paragraphs().Count + 1):
                shpCopy = shp.Duplicate()
                shpCopy.Select(False)
                shpCopy.Top = shp.Top
                shpCopy.Left = shp.Left
                
                for index in range(1, parIndex):  # @UnusedVariable
                    shpCopy.Top = shpCopy.Top + par_height(shpCopy.TextFrame.TextRange.Paragraphs(1))
                    shpCopy.TextFrame.TextRange.Paragraphs(1).Delete()
                
                for index in  range(parIndex, shp.TextFrame.TextRange.Paragraphs().Count+1):  # @UnusedVariable
                    shpCopy.TextFrame.TextRange.Paragraphs(2).Delete()
                    
                trim(shpCopy.TextFrame.TextRange)
                shp.Height = par_height(shp.TextFrame.TextRange.Paragraphs(1)) + shp.TextFrame.MarginTop + shp.TextFrame.MarginBottom
            
            shpCopy.Top = max(shpCopy.Top, shp.Top + shp.Height - shpCopy.Height)
            for index in range(1, shp.TextFrame.TextRange.Paragraphs().Count):  # @UnusedVariable
                shp.TextFrame.TextRange.Paragraphs(2).Delete()
                
            trim(shp.TextFrame.TextRange)
            shp.Height = par_height(shp.TextFrame.TextRange.Paragraphs(1)) + shp.TextFrame.MarginTop + shp.TextFrame.MarginBottom
            context.app.ActiveWindow.selection.ShapeRange.Distribute(1, False)

@bkt.configure(label='Ausrichten', uuid='a1d52fb5-a9e4-48ed-8b1c-c3f4c6f87462')  
@bkt.group
class Ausrichten(bkt.FeatureContainer):
    al_left = mso.button.ObjectsAlignLeftSmart
    al_right = mso.button.ObjectsAlignRightSmart
    al_top = mso.button.ObjectsAlignTopSmart
    al_bottom = mso.button.ObjectsAlignBottomSmart
    
    dist_h = mso.button.ObjectsAlignBottomSmart
    dist_v = mso.button.AlignDistributeVertically
    
    center_v = mso.button.ObjectsAlignMiddleVerticalSmart
    center_h = mso.button.ObjectsAlignCenterHorizontalSmart
    
def Group(*args, **kwargs):
    pass

#def create_ausrichten():
#    children = [mso.button.ObjectsAlignLeftSmart,
#                mso.button.ObjectsAlignRightSmart,
#                mso.button.ObjectsAlignTopSmart,
#                mso.button.ObjectsAlignBottomSmart,
#                mso.button.AlignDistributeHorizontally,
#                mso.button.AlignDistributeVertically,
#                mso.button.ObjectsAlignMiddleVerticalSmart,
#                mso.button.ObjectsAlignCenterHorizontalSmart]
#    return Group(id='group_bkt_align', label='Ausrichten', children=children)
    
@bkt.configure(label='Reihenfolge', uuid='843cd2d9-e81a-49d8-ac1a-888eb29d0deb')  
@bkt.group
class Reihenfolge(bkt.FeatureContainer):
    group = mso.button.ObjectsGroup
    ungroup = mso.button.ObjectsUngroup
    send_to_back = mso.button.ObjectSendToBack
    bring_to_front =  mso.button.ObjectBringToFront

@bkt.configure(label='Operationen', uuid='e267a046-6132-4ebb-9ffc-4deccd9679b8')
@bkt.group
class Operationen(bkt.FeatureContainer):
    interset = mso.button.ShapesIntersect
    combine = mso.button.ShapesCombine
    substract = mso.button.ShapesSubtract
    union = mso.button.ShapesUnion


#def create_bkt_dev():
#    # FIXME: geht nicht und wird vermutlich auch nie gehen, idQ ist buggy :(
#    children = [ForeignControl(id='ReloadBKT', namespace='bktdev'),
#                ForeignControl(id='UnloadBKT', namespace='bktdev')]
#    return Group(id='GroupBKTDev', label="BKT Dev", children=children)


@bkt.powerpoint
@bkt.configure(label='Toolbox CS')
@bkt.tab
class CSToolboxTab(bkt.FeatureContainer):
    g_table = bkt.use(TableGroup)
    g_table_ip = bkt.use(TableInPlaceGroup)
    g_misc = bkt.use(MiscGroup)
    g_ausrichten = bkt.use(Ausrichten)
    g_reihenfolge = bkt.use(Reihenfolge)
    g_Operationen = bkt.use(Operationen)


@bkt.powerpoint
@bkt.configure(label='Toolbox CS Test', uuid='7247e0bb-3a1c-4705-a1fe-2a8f43853e4c')
@bkt.tab
class CSTestTab(bkt.FeatureContainer):
    test1 = bkt.use(TestGroup)
    test2 = bkt.use(TestGroup2)

'''
@bkt.powerpoint
class CSToolboxTab(bkt.Tab):
    label = 'Toolbox CS'
    
    groups = [TableGroup,
              TableInPlaceGroup,
              MiscGroup,
              create_ausrichten, 
              create_reihenfolge,
              create_operationen
              ]
    
'''