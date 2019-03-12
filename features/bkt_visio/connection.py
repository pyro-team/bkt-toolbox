# -*- coding: utf-8 -*-
'''
Created on 17.03.2017

@author: fstallmann
'''

import bkt
import math
import logging

try:
    from bkt.library import visio
except IOError:
    # System.IO.IOException
    # breaks if visio-interop-library could not be referenced
    visio = None


if visio:
    sec_con_pts = visio.Sections.visSectionConnectionPts

#FIXME: Refactor this whole section
#TODO: create named connection rows, like Connections.BKT_0_0

class ConnectionPointRow(object):
    def __init__(self,shape,row):
        self.shape = shape
        self.row = row
        self._type = int(shape.RowType(sec_con_pts, row.Index))
        self._name = row.NameU
        self._xcell = row.Cell(visio.Cells.visCnnctX).FormulaU
        self._ycell = row.Cell(visio.Cells.visCnnctY).FormulaU
        self._acell = row.Cell(visio.Cells.visCnnctA).FormulaU
        self._bcell = row.Cell(visio.Cells.visCnnctB).FormulaU
        self._ccell = row.Cell(visio.Cells.visCnnctC).FormulaU
        self._dcell = row.Cell(visio.Cells.visCnnctD).FormulaU

    @property
    def type(self):
        return self._type

    @property
    def name(self):
        return self._name
    
    @name.setter
    def name(self, value):
        self.row.NameU = value
        self._name = value

    @property
    def xcell(self):
        return self._xcell
    
    @xcell.setter
    def xcell(self, value):
        self.row.Cell(visio.Cells.visCnnctX).FormulaU = value
        self._xcell = value

    @property
    def ycell(self):
        return self._ycell
    
    @ycell.setter
    def ycell(self, value):
        self.row.Cell(visio.Cells.visCnnctY).FormulaU = value
        self._ycell = value

    @property
    def acell(self):
        return self._acell
    
    @acell.setter
    def acell(self, value):
        self.row.Cell(visio.Cells.visCnnctA).FormulaU = value
        self._acell = value

    @property
    def bcell(self):
        return self._bcell
    
    @bcell.setter
    def bcell(self, value):
        self.row.Cell(visio.Cells.visCnnctB).FormulaU = value
        self._bcell = value

    @property
    def ccell(self):
        return self._ccell
    
    @ccell.setter
    def ccell(self, value):
        self.row.Cell(visio.Cells.visCnnctC).FormulaU = value
        self._ccell = value

    @property
    def dcell(self):
        return self._dcell
    
    @dcell.setter
    def dcell(self, value):
        self.row.Cell(visio.Cells.visCnnctD).FormulaU = value
        self._dcell = value
    
    def copy_to(self,obj):
        if isinstance(obj, ConnectionPointRow):
            # obj.type = self.type
            # obj.name = self.name
            obj.xcell = self.xcell
            obj.ycell = self.ycell
            obj.acell = self.acell
            obj.bcell = self.bcell
            obj.ccell = self.ccell
            obj.dcell = self.dcell

class ConnectionPointSection(object):
    def __init__(self,shape):
        self.shape = shape
        if not shape.SectionExists(sec_con_pts, False):
            shape.AddSection(sec_con_pts)
        self.section = shape.Section(sec_con_pts)
        self.points = []

    def load_points(self):
        for row_i in range(0,int(self.section.Count)):
            self.points.append(ConnectionPointRow(self.shape,self.section.Row(row_i)))

    def clear_points(self):
    	self.points = []
    	self.shape.DeleteSection(sec_con_pts)
        self.shape.AddSection(sec_con_pts)
        self.section = self.shape.Section(sec_con_pts)

    def add_point(self, row):
        # if row.type == 153: #153=visTagCnnctPt
        #     new_row_i = self.shape.AddRow(sec_con_pts,visio.Rows.visRowLast,153)
        # else: #185=visTagCnnctNamed
        #     new_row_i = self.shape.AddNamedRow(sec_con_pts,row.name,185)
        new_row_i = self.shape.AddRow(sec_con_pts,visio.Rows.visRowLast,visio.Cells.visCnnctX)
        new_row = ConnectionPointRow(self.shape, self.section.Row(new_row_i))
        row.copy_to(new_row)
    	self.points.append(new_row)

    def create_point(self, x, y, a=0, b=0, c=0, d=None):
        new_row_i = self.shape.AddRow(sec_con_pts,visio.Rows.visRowLast,visio.Cells.visCnnctX)
        new_row = ConnectionPointRow(self.shape, self.section.Row(new_row_i))
        new_row.xcell = x
    	new_row.ycell = y
    	new_row.acell = a
    	new_row.bcell = b
    	new_row.ccell = c
    	if not d is None:
    		new_row.dcell = d
    	self.points.append(new_row)

class ConnectionPoints(object):
    copied_section = None

    def del_con_pts(self, shapes):
        for shape in shapes:
            sec = ConnectionPointSection(shape.shape)
            sec.clear_points()
    
    def enabled_copy(self, shapes):
        return len(shapes) == 1
        # return (len(shapes) == 1 and shapes[0].shape.SectionExists(sec_con_pts, False) and shapes[0].shape.Section(sec_con_pts).Count > 0)
    
    def enabled_paste(self, shapes):
        return len(shapes) > 0 and self.copied_section != None
        # return len(shapes) > 0 and self.copied_section != None and len(self.copied_section.points) > 0
    
    def copy_con_pts(self, shapes):
        s = shapes[0].shape
        if not s.SectionExists(sec_con_pts, False):
            return
        self.copied_section = ConnectionPointSection(s)
        self.copied_section.load_points()
        # rows = int(s.Section(sec_con_pts).Count)
        # if rows == 0:
        #     return
        # self.copied_points = []
        # for row_i in range(0,rows):
        #     row = s.Section(sec_con_pts).Row(row_i)
        #     self.copied_points.append(ConnectionPointRow(s,row))
    
    def paste_con_pts(self, shapes):
        for shape in shapes:
            sec = ConnectionPointSection(shape.shape)
            for point in self.copied_section.points:
            	sec.add_point(point)
                # if pts.type == 153: #153=visTagCnnctPt
                #     new_row_i = s.AddRow(sec_con_pts,visio.Rows.visRowLast,153)
                # else: #185=visTagCnnctNamed
                #     new_row_i = s.AddNamedRow(sec_con_pts,pts.name,185)
                # new_row = ConnectionPointRow(s, s.Section(sec_con_pts).Row(new_row_i))
                # pts.copy_to(new_row)

    def add_con_pts_edges(self, shapes):
        for shape in shapes:
            sec = ConnectionPointSection(shape.shape)
            sec.create_point("Width*0", "Height*0")
            sec.create_point("Width*1", "Height*0")
            sec.create_point("Width*0", "Height*1")
            sec.create_point("Width*1", "Height*1")

    def add_con_pts_sides1(self, shapes):
        for shape in shapes:
            sec = ConnectionPointSection(shape.shape)
            sec.create_point("Width*0", "Height*0.5")
            sec.create_point("Width*1", "Height*0.5")
            sec.create_point("Width*0.5", "Height*0")
            sec.create_point("Width*0.5", "Height*1")

    def add_con_pts_sides2(self, shapes):
        for shape in shapes:
            sec = ConnectionPointSection(shape.shape)
            sec.create_point("Width*0", "Height*0.333")
            sec.create_point("Width*0", "Height*0.667")
            sec.create_point("Width*1", "Height*0.333")
            sec.create_point("Width*1", "Height*0.667")
            sec.create_point("Width*0.333", "Height*0")
            sec.create_point("Width*0.667", "Height*0")
            sec.create_point("Width*0.333", "Height*1")
            sec.create_point("Width*0.667", "Height*1")

    def add_con_pts_sides3(self, shapes):
        for shape in shapes:
            sec = ConnectionPointSection(shape.shape)
            sec.create_point("Width*0", "Height*0.25")
            sec.create_point("Width*0", "Height*0.5")
            sec.create_point("Width*0", "Height*0.75")
            sec.create_point("Width*1", "Height*0.25")
            sec.create_point("Width*1", "Height*0.5")
            sec.create_point("Width*1", "Height*0.75")
            sec.create_point("Width*0.25", "Height*0")
            sec.create_point("Width*0.5", "Height*0")
            sec.create_point("Width*0.75", "Height*0")
            sec.create_point("Width*0.25", "Height*1")
            sec.create_point("Width*0.5", "Height*1")
            sec.create_point("Width*0.75", "Height*1")

connection_points = ConnectionPoints()

# http://www.visguy.com/2016/07/12/no-glue-to/
class GlueAble(object):

    @staticmethod
    def get_glueable_state(shapes):
        shape = shapes[0]
        return shape.cells["GlueType"].ResultIU == 8 and shape.cells["ObjType"].ResultIU == 4

    @staticmethod
    def make_unglueable(shapes):
        for shape in shapes:
            shape.cells["GlueType"].ResultIU = 8
            shape.cells["ObjType"].ResultIU = 4
            # Only for groups:
            if shape.type == 2:
                shape.cells["IsSnapTarget"].ResultIU = 0
    
    @staticmethod
    def make_glueable(shapes):
        for shape in shapes:
            shape.cells["GlueType"].ResultIU = 2
            shape.cells["ObjType"].ResultIU = 2
            # Only for groups:
            if shape.type == 2:
                shape.cells["IsSnapTarget"].ResultIU = 1
    
    @staticmethod
    def make_glue_reset(shapes):
        for shape in shapes:
            shape.cells["GlueType"].FormulaU = "="
            shape.cells["ObjType"].FormulaU = "="
            # Only for groups:
            if shape.type == 2:
                shape.cells["IsSnapTarget"].FormulaU = "="



verbindungspunkte_gruppe = bkt.ribbon.Group(
    label="Verbindungspunkte",
    image_mso='ConnectionPointTool',
    children=[
        bkt.ribbon.Box(box_style="vertical",
            children=[
                # bkt.ribbon.Button(
                    # id = 'copy_con_pts_from_master',
                    # label="Master-V.pkt. auf alle kopieren",
                    # show_label=True,
                    # image_mso='HappyFace',
                    # screentip="Alle Verbindungspunkte des Master-Shapes (dicker Rahmen) auf andere Shapes kopieren",
                    # on_action=bkt.Callback(connection_points.copy_con_pts_from_master, shapes=True, shapes_min=2),
                    # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                # ),
                bkt.ribbon.Button(
                    id = 'copy_con_pts',
                    label="Punkte kopieren",
                    show_label=True,
                    image_mso='Copy',
                    screentip="Alle Verbindungspunkte des gewählten Shapes kopieren",
                    on_action=bkt.Callback(connection_points.copy_con_pts),
                    get_enabled = bkt.Callback(connection_points.enabled_copy)
                ),
                bkt.ribbon.Button(
                    id = 'paste_con_pts',
                    label="Punkte einfügen",
                    show_label=True,
                    image_mso='Paste',
                    screentip="Kopierte Verbindungspunkte in gewählte Shapes einfügen",
                    on_action=bkt.Callback(connection_points.paste_con_pts),
                    get_enabled = bkt.Callback(connection_points.enabled_paste)
                ),
                bkt.ribbon.Button(
                    id = 'del_con_pts',
                    label="Punkte löschen",
                    show_label=True,
                    image_mso='Delete',
                    screentip="Alle Verbindungspunkte der gewählten Shapes entfernen",
                    on_action=bkt.Callback(connection_points.del_con_pts, shapes=True, shapes_min=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
            ]
        ),
        bkt.ribbon.Box(box_style="vertical",
            children=[
                bkt.ribbon.Menu(
                    label='Neue Punkte anlegen',
                    screentip='Auswahl von Verbindungspunkten',
                    supertip='Vorgefertigte Standard-Auswahl an Verbindungspunkten anlegen ',
                    show_label=True,
                    image_mso='ConnectionPointTool',
                    children = [
                        #bkt.ribbon.MenuSeparator(title="Verbindungspunkte"),
                        bkt.ribbon.Button(
                            id = 'add_con_pts_edges',
                            image = 'conpts_edges',
                            label='In den 4 Ecken',
                            show_label=True,
                            on_action=bkt.Callback(connection_points.add_con_pts_edges),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'add_con_pts_sides1',
                            image = 'conpts_1sd',
                            label='1x je Seite',
                            show_label=True,
                            on_action=bkt.Callback(connection_points.add_con_pts_sides1),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'add_con_pts_sides2',
                            image = 'conpts_2sd',
                            label='2x je Seite',
                            show_label=True,
                            on_action=bkt.Callback(connection_points.add_con_pts_sides2),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'add_con_pts_sides3',
                            image = 'conpts_3sd',
                            label='3x je Seite',
                            show_label=True,
                            on_action=bkt.Callback(connection_points.add_con_pts_sides3),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ]
                ),
                bkt.ribbon.Button(
                    id = 'make_unglueable',
                    label="Shape nicht klebend",
                    show_label=True,
                    image_mso='SnapToggle',
                    screentip="Shape nicht klebend machen",
                    supertip="Macht ein Shape unklebend für Verbinder und andere Shapes über die ShapeSheet-Eigenschaten GlueType, ObjType und IsSnapTarget (nur Gruppen).",
                    on_action=bkt.Callback(GlueAble.make_unglueable, shapes=True, shapes_min=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.SplitButton(
                    children=[
                        bkt.ribbon.Button(
                            id = 'make_glueable',
                            label="Shape klebend",
                            show_label=True,
                            image_mso='GlueToggle',
                            screentip="Shape klebend machen",
                            supertip="Macht ein Shape klebend für Verbinder über die ShapeSheet-Eigenschaten GlueType, ObjType und IsSnapTarget (nur Gruppen).",
                            on_action=bkt.Callback(GlueAble.make_glueable, shapes=True, shapes_min=1),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                        ),
                        bkt.ribbon.Menu(children=[
                            bkt.ribbon.Button(
                                id = 'make_glue_reset',
                                label="Zurücksetzen",
                                show_label=True,
                                #image_mso='GlueToggle',
                                screentip="Shape zurücksetzen",
                                supertip="Setzt die ShapeSheet-Eigenschaten GlueType, ObjType und IsSnapTarget (nur Gruppen) auf Standard zurück.",
                                on_action=bkt.Callback(GlueAble.make_glue_reset, shapes=True, shapes_min=1),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                            )
                        ])
                    ]
                )
            ]
        )
    ]
)