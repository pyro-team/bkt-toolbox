# -*- coding: utf-8 -*-
'''
Created on 13.08.2015

@author: rdebeerst
'''



import bkt
import bkt.library.powerpoint as powerpoint
import bkt.library.bezier as bezier
import bkt.library.algorithms
import System
import bkt.ui
import json
import math

import ruben as toolbox_rd



class Adjustments(object):

    @staticmethod
    def adjustment_edit_box(num):
        editbox= bkt.ribbon.EditBox(
            id    ='adjustment-' + str(num),
            label =' value ' + str(num),
            sizeString = '#######',
            
            on_change   = bkt.Callback(
                lambda shapes, value: map( lambda shape: Adjustments.set_adjustment(shape, num, value), shapes),
                shapes=True),
            
            get_text    = bkt.Callback(
                lambda shapes       : Adjustments.get_adjustment(shapes[0], num),
                shapes=True),
            
            get_enabled = bkt.Callback(
                lambda shapes : Adjustments.is_enabled(shapes[0], num),
                shapes=True)
        )
        return editbox


    @staticmethod
    def set_adjustment(shape, num, value):
        if shape.adjustments.count >= num:
            shape.adjustments.item[num] = value

    @staticmethod
    def get_adjustment(shape, num):
        try:
            if shape.adjustments.count >= num:
                return shape.adjustments.item[num]
            else:
                return None
        except:
            return None

    @staticmethod
    def is_enabled(shape, num):
        try:
            return (shape.adjustments.count >=num)
        except:
            return False



group_adjustments = bkt.ribbon.Group(
    label = "Shape Adjustments",
    
    children=[
        Adjustments.adjustment_edit_box(num)
        
        for num in range(1,10)
    ]
)





class VariousShapes(object): 
    
    @staticmethod
    def create_line(slide, shapes):
        shapeCount = slide.shapes.count
        reset_selection = True
        for shape in shapes:
            line = slide.shapes.addLine( shape.left, shape.top+shape.height, shape.left+shape.width, shape.top+shape.height )
            line.line.weight = 1.5
            line.select(reset_selection)
            reset_selection = False
        




class ShapePoints(object):
    
    @classmethod
    def display_points(cls, shape):
        
        if not shape.Type == powerpoint.MsoShapeType['msoFreeform']:
            shape.Nodes.SetPosition(1, shape.Left, shape.Top)
        
        pointlist = "["
        first = True
        for node in shape.nodes:
            if not first:
                pointlist += ","
            pointlist += "\r\n"
            pointlist += '  {"x":' + str(node.points[0,0])
            pointlist += ', "y":' + str(node.points[0,1])
            # pointlist += ', "segmentType": ' + str(node.segmentType)
            # pointlist += ', "editingType": ' + str(node.editingType)
            pointlist += '}'
            first = False
        pointlist += "\r\n]"
        #bkt.console.show_message(json)
        
        def json_callback(json_points):
            cls.change_points(shape, json_points=json_points)
        
        bkt.console.show_input(pointlist, json_callback)
        #shape.textframe.textrange.text = json
    
    
    @staticmethod
    def change_points(shape, json_points=None):
        points = json.loads(json_points)
        
        # richtige Anzahl Punkte
        while len(points) > shape.nodes.count:
            shape.nodes.insert(shape.nodes.count, 0,0,  0.0, 0.0)
        while len(points) < shape.nodes.count:
            shape.nodes.delete(shape.nodes.count)
        
        index = 0
        for p in points:
            shape.nodes.setPosition(index+1, p['x'], p['y'])
            # shape.nodes.setEditingType(index+1, p['editingType'])
            # shape.nodes.setSegmentType(index+1, p['segmentType'])
            index += 1
        
        
        # index = 0
        # for p in points:
        #     shape.nodes.insert(index+1, p['segmentType'], p['editingType'],  p['x'], p['y'])
        #     index += 1
        #
        # # while len(points) > shape.nodes.count:
        # #     shape.nodes.insert(shape.nodes.count, 0,0,  0.0, 0.0)
        # #
        # while len(points) < shape.nodes.count:
        #     shape.nodes.delete(shape.nodes.count)
        
group_shape_points = bkt.ribbon.Group(
    label=u"Shape Points",
    children=[
        bkt.ribbon.Button(
            label=u'Shape Points',
            imageMso='ObjectEditPoints', show_label=False,
            on_action=bkt.Callback(ShapePoints.display_points)
        )
    ]
)






class CopyPasteStyle(object):
    
    ref_shape = None
    copy_settings = {
        'background': True,
        'img': True,
        'line': True,
        'position': False,
        'size': True,
    }
    
    def copy_style(self, shape):
        if shape.type == powerpoint.MsoShapeType['msoPicture']:
            self.ref_shape = shape
        else:
            self.ref_shape = None
    
    # def image_selected(self, shape):
    #     return shape.type == powerpoint.MsoShapeType['msoPicture'];
    
    def paste_style(self, shapes):
        if self.ref_shape == None:
            return
        for shape in shapes:
            if self.copy_settings['size']:
                shape.width = self.ref_shape.width
                shape.height = self.ref_shape.height

            if self.copy_settings['position']:
                shape.left = self.ref_shape.left
                shape.top = self.ref_shape.top
                shape.rotation = self.ref_shape.rotation

            if self.copy_settings['background']:
                #shape.Fill.Type = self.ref_shape.Fill.Type
                shape.Fill.ForeColor.RGB = self.ref_shape.Fill.ForeColor.RGB
                shape.Fill.ForeColor.SchemeColor = self.ref_shape.Fill.ForeColor.SchemeColor
                shape.Fill.ForeColor.Brightness = self.ref_shape.Fill.ForeColor.Brightness
                shape.Fill.ForeColor.TintAndShade = self.ref_shape.Fill.ForeColor.TintAndShade
            
            if self.copy_settings['line']:
                shape.Line.Style = self.ref_shape.Line.Style
                shape.Line.DashStyle = self.ref_shape.Line.DashStyle
                shape.Line.Weight = self.ref_shape.Line.Weight
                shape.Line.Transparency = self.ref_shape.Line.Transparency
                shape.Line.ForeColor.RGB = self.ref_shape.Line.ForeColor.RGB
                shape.Line.ForeColor.SchemeColor = self.ref_shape.Line.ForeColor.SchemeColor
                shape.Line.ForeColor.Brightness = self.ref_shape.Line.ForeColor.Brightness
                shape.Line.ForeColor.TintAndShade = self.ref_shape.Line.ForeColor.TintAndShade
            
            if self.copy_settings['img'] and shape.type == powerpoint.MsoShapeType['msoPicture']:
                shape.PictureFormat.crop.ShapeHeight = self.ref_shape.PictureFormat.crop.ShapeHeight 
                shape.PictureFormat.crop.ShapeWidth  = self.ref_shape.PictureFormat.crop.ShapeWidth  
                shape.PictureFormat.crop.ShapeTop    = self.ref_shape.PictureFormat.crop.ShapeTop    
                shape.PictureFormat.crop.ShapeLeft   = self.ref_shape.PictureFormat.crop.ShapeLeft   
            
                shape.PictureFormat.crop.PictureHeight  = self.ref_shape.PictureFormat.crop.PictureHeight
                shape.PictureFormat.crop.PictureWidth   = self.ref_shape.PictureFormat.crop.PictureWidth
                shape.PictureFormat.crop.PictureOffsetX = self.ref_shape.PictureFormat.crop.PictureOffsetX
                shape.PictureFormat.crop.PictureOffsetY = self.ref_shape.PictureFormat.crop.PictureOffsetY
    
    # def paste_style_enabled(self, shape):
    #     return shape.type == powerpoint.MsoShapeType['msoPicture'] and self.ref_shape != None;
    
    def setting_size(self, pressed):
        self.copy_settings['size'] = (pressed == True)
    
    def setting_size_pressed(self):
        return self.copy_settings['size'] == True
    
    def setting_background(self, pressed):
        self.copy_settings['background'] = (pressed == True)
    
    def setting_background_pressed(self):
        return self.copy_settings['background'] == True
    
    
    def setting_img(self, pressed):
        self.copy_settings['img'] = (pressed == True)
    
    def setting_img_pressed(self):
        return self.copy_settings['img'] == True



copy_paste_style = CopyPasteStyle()

group_copy_paste_style = bkt.ribbon.Group(
    label=u"Copy Style",
    children=[
        bkt.ribbon.Button(label='copy', screentip='copy image settings: Zuschneideposition, ...', 
            on_action=bkt.Callback(copy_paste_style.copy_style),
            #get_enabled=bkt.Callback(copy_paste_style.image_selected)
        ),
        bkt.ribbon.Button(label='paste', screentip='paste image settings: Zuschneideposition, ...',
            on_action=bkt.Callback(copy_paste_style.paste_style),
            #get_enabled=bkt.Callback(copy_paste_style.paste_style_enabled)
        ),
        bkt.ribbon.ToggleButton(label="SIZE",
            on_toggle_action=bkt.Callback(copy_paste_style.setting_size),
            get_pressed=bkt.Callback(copy_paste_style.setting_size_pressed)
        ),
        bkt.ribbon.ToggleButton(label="BACKGROUND",
            on_toggle_action=bkt.Callback(copy_paste_style.setting_background),
            get_pressed=bkt.Callback(copy_paste_style.setting_background_pressed)
        ),
        bkt.ribbon.ToggleButton(label="IMG",
            on_toggle_action=bkt.Callback(copy_paste_style.setting_img),
            get_pressed=bkt.Callback(copy_paste_style.setting_img_pressed)
        )
    ]
)
    
    
    
    
    

class ShapeMetaStyle(object):
    
    META_STYLE_NAME = "RD-METASTYLE"
    META_STYLE_SETTINGS = "RD-METASTYLE-SETTING"
    
    style_names = None
    settings = None
    
    
    def set_meta_style(self, shapes, value, slides):
        for shape in shapes:
            if value == '':
                shape.tags.delete(self.META_STYLE_NAME)
            else:
                shape.tags.add(self.META_STYLE_NAME, value)
        # Nach Aktion die Liste aktualisieren
        self.update_style_names(slides)
    
    def set_meta_style_item_count(self, slides):
        # Liste bei jedem Aufruf aktualisieren, damit Wechsel zwischen PowerPoint-Files möglich wird
        self.update_style_names(slides)
        return len(self.style_names)
    
    def set_meta_style_item_label(self, index):
        return self.style_names[index-1]


    def get_meta_style(self, shapes):
        shapetypes = map( lambda shape: self.get_tag_value(shape, self.META_STYLE_NAME, ''), shapes )
        shapetypes = list(set(shapetypes))
        if len(shapetypes) == 1:
            return shapetypes[0]
        else:
            return ''
        #return self.get_tag_value(shapes[0], 'SEN-RD-METASTYLE', '')
    
    
    def get_tag_value(self, obj, tagname, default=''):
        for idx in range(1,obj.tags.count+1):
            if obj.tags.name(idx) == tagname:
                return obj.tags.value(idx)
        return default
    
    def apply_style(self, shape, slides):
        style = self.get_tag_value(shape, self.META_STYLE_NAME)
        if style == '':
            return

        master_shape = shape
        settings = self.get_style_settings(slides[0].parent, style)
        
        for slide in slides:
            for shape in slide.shapes:
                if shape != master_shape:
                    if self.get_tag_value(shape, self.META_STYLE_NAME) == style:
                        self.copy_style(master_shape, shape, self.default_style)
        
        self.update_style_names(slides)
    
    
    def get_style_settings(self, presentation, stylename):
        if self.settings == None:
            # deserialize json
            self.settings = {}
            try:
                self.settings = json.loads(get_tag_value(presentation, self.META_STYLE_SETTINGS, None)) or {}
            except:
                pass
                        
        else:
            if self.settings.has_key(stylename):
                return self.settings[stylename]
            else:
                return self.default_style
    
    
    default_style = {
        "type": True,
        "size": True,
        "background": True,
        "linestyle": True,
        "orientation": True,
        "textformat": True,
        "margin": True
    }
    
    def copy_style(self, origin, target, settings):
        
        # Shape Type
        if settings['type']:
            if target.Type == origin.Type and origin.Type == powerpoint.MsoShapeType['msoAutoShape']:
                target.AutoShapeType = origin.AutoShapeType
        
        # Size
        if settings['size']:
            target.width = origin.width
            target.height = origin.height
        
        # Background
        if settings['background']:
            target.Fill.ForeColor.RGB = origin.Fill.ForeColor.RGB

        # Line Style
        if settings['linestyle']:
            target.Line.Style = origin.Line.Style
            target.Line.DashStyle = origin.Line.DashStyle
            target.Line.Transparency = origin.Line.Transparency
            target.Line.ForeColor.RGB = origin.Line.ForeColor.RGB
            target.Line.Weight = origin.Line.Weight
        
        # Text Ausrichtung
        if settings['orientation']:
            target.TextFrame.TextRange.ParagraphFormat.Alignment = origin.TextFrame.TextRange.ParagraphFormat.Alignment
            target.Textframe.HorizontalAnchor = origin.Textframe.HorizontalAnchor
            target.Textframe.VerticalAnchor = origin.Textframe.VerticalAnchor
        
        # Text Format
        if settings['textformat']:
            target.Textframe.TextRange.Font.Color.RGB   = origin.Textframe.TextRange.Font.Color.RGB
            target.Textframe.TextRange.Font.Size        = origin.Textframe.TextRange.Font.Size
            target.Textframe.TextRange.Font.Bold        = origin.Textframe.TextRange.Font.Bold
            target.Textframe.TextRange.Font.Italic      = origin.Textframe.TextRange.Font.Italic     
            target.Textframe.TextRange.Font.Underline   = origin.Textframe.TextRange.Font.Underline  
        
        # Margins
        if settings['margin']:
            target.Textframe.marginLeft   =  origin.Textframe.marginLeft  
            target.Textframe.marginRight  =  origin.Textframe.marginRight 
            target.Textframe.marginTop    =  origin.Textframe.marginTop   
            target.Textframe.marginBottom =  origin.Textframe.marginBottom
        
        #FIXME: more to come...
    
    
    
    
    def update_style_names(self, slides):
        style_names = set()
        
        for slide in slides:
            for shape in slide.shapes:
                style_names =style_names.union( set( [ self.get_tag_value(shape, self.META_STYLE_NAME, '') ] ))
        
        self.style_names = list(style_names)



shape_meta_style = ShapeMetaStyle()
group_meta_style = bkt.ribbon.Group(
	label=u'Master Styles',
	children=[
		bkt.ribbon.ComboBox(
			label='Name', size_string='###############', show_label=True,
			on_change=bkt.Callback(shape_meta_style.set_meta_style),
			get_item_count=bkt.Callback(shape_meta_style.set_meta_style_item_count),
			get_item_label=bkt.Callback(shape_meta_style.set_meta_style_item_label),
			get_text=bkt.Callback(shape_meta_style.get_meta_style),
		),
		bkt.ribbon.Button(label='Style übertragen', show_label=True,
			on_action=bkt.Callback(shape_meta_style.apply_style)
		)
	]
)




    
class Diverses(object):
    
    @classmethod
    def make_black_white(cls, shapes):
        for shape in shapes:
            shape.Fill.ForeColor.RGB = cls.black_white_rgb(shape.Fill.ForeColor.RGB)
            shape.Line.ForeColor.RGB = cls.black_white_rgb(shape.Line.ForeColor.RGB)
            shape.Textframe.TextRange.Font.Color.RGB = cls.black_white_rgb(shape.Textframe.TextRange.Font.Color.RGB)
    
    @staticmethod
    def black_white_rgb(rgb):
        r= rgb % 256
        g= ((rgb-r)/256) % 256
        b=(rgb-r-g*256)/256/256
        bw = int(round(0.21*r+0.72*g+0.07*b))
        return bw+bw*256+bw*256*256
    
    
    # @staticmethod
    # def circ_w_connectors(slide):
    #     kurven = bezier.kreisSegmente(4*10, 100, [200,200])
    #     # Kurve aus Beziers erstellen
    #     # start beim ersten Punkt
    #     P = kurven[0][0][0];
    #     #bkt.helpers.message( "%d/%d" % (P[0], P[1]) )
    #     ffb = slide.Shapes.BuildFreeform(1, P[0], P[1])
    #     for kl in kurven:
    #         k = kl[0]
    #         # von den nächsten Beziers immer die nächsten Punkte angeben
    #         ffb.AddNodes(1, 0, k[1][0], k[1][1], k[2][0], k[2][1], k[3][0], k[3][1])
    #         # Parameter: SegmentType, EditingType, X1,Y1, X2,Y2, X3,Y3
    #         # SegmentType: 0=Line, 1=Curve
    #         # EditingType: 0=Auto, 1=Corner (keine Verbindungspunkte), 2=Smooth, 3=Symmetric  --> Zweck?
    #     shp = ffb.ConvertToShape()
    #
    #
    # @staticmethod
    # def box_w_connectors(slide):
    #     #ffb = slide.Shapes.BuildFreeform(1, )
    #     corners = [ [100,100],
    #                 [300,100],
    #                 [300,200],
    #                 [100,200]]
    #
    #     def line_to(ffb, origin, dest, segments):
    #         deltaX = dest[0]-origin[0]
    #         deltaY = dest[1]-origin[1]
    #         for i in range(0,segments+1):
    #             # SegmentType=0 (Line), EditingType=1 (Corner)
    #             #bkt.helpers.message( "%d/%d" % (origin[0]+float(i)/segments*deltaX, origin[1]+float(i)/segments*deltaY) )
    #             # SegmentType=0 (Line), EditingType=1 (Corner)
    #             #ffb.addNodes(0,1, origin[0]+float(i)/segments*deltaX, origin[1]+float(i)/segments*deltaY)
    #
    #             ffb.addNodes(0,0, origin[0]+float(i)/segments*deltaX, origin[1]+float(i)/segments*deltaY)
    #
    #     segments = 3
    #
    #     # EditingType=1 (Corner)
    #     ffb = slide.Shapes.BuildFreeform(1, *corners[0])
    #     last_corner = corners[0]
    #     for c in corners:
    #         #ffb.addNodes(0,1,*c)
    #         line_to(ffb, last_corner, c, segments)
    #         last_corner = c
    #
    #     line_to(ffb, last_corner, corners[0], segments)
    #     # ffb.addNodes(0,1,*corners[0])
    #     shp = ffb.ConvertToShape()
    #
    #     # durch BuildFreeform wird Initialpunkt 4-mal gesetzt. ueberfluessige Punkte entfernen
    #     shp.nodes.delete(4)
    #     shp.nodes.delete(3)
    #     shp.nodes.delete(2)
    #
    #     # alles außer Eckpunkte auf EditingType=Smooth
    #     # segemnts=3 --> Punktliste: [3, 4, 7, 8, 11, 12, 15, 16]
    #     # pro Seite werden segment+1 punkte gezeichnet, der jeweils erste und letzte ist die Ecke --> mod (segments+1) prüfen
    #     # shift um zwei, da VBA-Zählung bei 1 beginnt und um Initialpunkt auszuschließen
    #     # for i in map(lambda x: x+2, filter(lambda x: x%(segemnts+1) != 0 and x%(segemnts+1) != segments, range(4*(segemnts+1)))):
    #     #     shp.nodes.SetEditingType(i, 2)
        















class StateShape(object):
    
    @staticmethod
    def is_state_shape(shape):
        return shape.Type == powerpoint.MsoShapeType['msoGroup']
    
    @staticmethod
    def next_state(shape):
        # ungroup shape, to get list of groups inside grouped items
        ungrouped_shapes = shape.Ungroup()
        shapes = list(iter(ungrouped_shapes))
        shapes.sort(key=lambda s: s.zorderposition)
        for s in shapes:
            s.visible = False
        shapes[-1].zorder(1)
        shapes[-1].visible = True
        ungrouped_shapes.group().select()

    @staticmethod
    def previous_state(shape):
        ungrouped_shapes = shape.Ungroup()
        shapes = list(iter(ungrouped_shapes))
        shapes.sort(key=lambda s: s.zorderposition)
        for s in shapes:
            s.visible = False
        shapes[0].zorder(0)
        shapes[0].visible = True
        ungrouped_shapes.group().select()
        





class EnhancedMetaFile(object):
    
    @staticmethod
    def convert_selection_to_emf(slide, selection):
        # remember position
        left = min([s.left for s in selection.shaperange])
        top  = min([s.top  for s in selection.shaperange])
        # cut shapes
        selection.Cut()
        # paste as ppPasteEnhancedMetafile
        grp = slide.Shapes.PasteSpecial(2)
        # reposition
        grp.left = left
        grp.top = top
        
    

class SolidShadow(object):
    
    @staticmethod
    def solid_white_shadow(shapes):
        for shape in shapes:
            #shape.Shadow.Type = -2
            shape.Shadow.Visible = True
            shape.Shadow.ForeColor.RGB = 16777215 # white
            shape.Shadow.OffsetX = 1
            shape.Shadow.OffsetY = 1
            shape.Shadow.Blur = 0
            shape.Shadow.Size = 100
            shape.Shadow.Transparency = 0



group_diverses = bkt.ribbon.Group(
    label=u"Div.",
    children=[
        bkt.ribbon.Button(label=u"Schwarz/weiß", show_label=True, on_action=bkt.Callback(Diverses.make_black_white)),
        # bkt.ribbon.Button(label='circ w connectors', show_label=True, on_action=bkt.Callback(Diverses.circ_w_connectors)),
        # bkt.ribbon.Button(label='box w connectors', show_label=True, on_action=bkt.Callback(Diverses.box_w_connectors)),
        bkt.ribbon.Button(label=u"Replace by emf-image", show_label=True, on_action=bkt.Callback(EnhancedMetaFile.convert_selection_to_emf)),
        bkt.ribbon.Button(label=u"Weißer Schatten", show_label=True, on_action=bkt.Callback(SolidShadow.solid_white_shadow)),
        
    ]
)



bkt.powerpoint.add_tab(
    bkt.ribbon.Tab(
        label=u'Toolbox RD2',
        id='RubensTab',
        # get_visible defaults to False during async-startup
        get_visible=bkt.Callback(lambda:True),
        children=[
            bkt.ribbon.Group(
                label="Zusatzformen",
                children = [
                    bkt.ribbon.Button(
                        id="rd_underline",
                        label='Unterstreichen', screentip='Shape mit Linie versehen',
                        image_mso='Underline',
                        on_action=bkt.Callback(VariousShapes.create_line)
                    ),
                    bkt.ribbon.Box(children =[
                        bkt.ribbon.Button(
                            id="rd_change_state_prev",
                            label=u'Status wechseln   «',
                            image_mso='GroupSmartArtQuickStyles', #'ScreenNavigatorForward', # GroupSmartArtQuickStyles, GroupShow, GroupDiagramStylesClassic
                            on_action=bkt.Callback(StateShape.previous_state),
                            get_enabled=bkt.Callback(StateShape.is_state_shape),
                        ),
                        bkt.ribbon.Button(
                            id="rd_change_state_next",
                            label=u"»",
                            on_action=bkt.Callback(StateShape.next_state),
                            get_enabled=bkt.Callback(StateShape.is_state_shape),
                        )
                    ])
                ]
            ),
            group_copy_paste_style,
            group_adjustments,
            group_diverses,
            group_shape_points,
        ]
    )
)


