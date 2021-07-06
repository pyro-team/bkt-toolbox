# -*- coding: utf-8 -*-
'''
Created on 02.11.2017

@author: fstallmann
'''

from __future__ import absolute_import, division

import time
import json # required for tags
from collections import namedtuple # required for color class
from contextlib import contextmanager # for locale helper

import System.Array # to create int/str-Arrays

from bkt import settings, dotnet # required to save global locpin setting
from bkt.library import algorithms # required for color helper

# DO NOT REMOVE REFERENCE
# reference is used by other modules
PowerPoint = dotnet.import_powerpoint()

ptToCmFactor = 2.54 / 72
def pt_to_cm(pt):
    return float(pt) * ptToCmFactor
def cm_to_pt(cm):
    return float(cm) / ptToCmFactor

# shape.AutoShapeType
MsoAutoShapeType = {
    'msoShape10pointStar':                        149,   # 10-point star.
    'msoShape12pointStar':                        150,   # 12-point star.
    'msoShape16pointStar':                         94,   # 16-point star.
    'msoShape24pointStar':                         95,   # 24-point star.
    'msoShape32pointStar':                         96,   # 32-point star.
    'msoShape4pointStar':                          91,   # 4-point star.
    'msoShape5pointStar':                          92,   # 5-point star.
    'msoShape6pointStar':                         147,   # 6-point star.
    'msoShape7pointStar':                         148,   # 7-point star.
    'msoShape8pointStar':                          93,   # 8-point star.
    'msoShapeActionButtonBackorPrevious':         129,   # Back or Previous button. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonBeginning':              131,   # Beginning button. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonCustom':                 125,   # Button with no default picture or text. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonDocument':               134,   # Document button. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonEnd':                    132,   # End button. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonForwardorNext':          130,   # Forward or Next button. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonHelp':                   127,   # Help button. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonHome':                   126,   # Home button. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonInformation':            128,   # Information button. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonMovie':                  136,   # Movie button. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonReturn':                 133,   # Return button. Supports mouse-click and mouse-over actions.
    'msoShapeActionButtonSound':                  135,   # Sound button. Supports mouse-click and mouse-over actions.
    'msoShapeArc':                                 25,   # Arc.
    'msoShapeBalloon':                            137,   # Balloon.
    'msoShapeBentArrow':                           41,   # Block arrow that follows a curved 90-degree angle.
    'msoShapeBentUpArrow':                         44,   # Block arrow that follows a sharp 90-degree angle. Points up by default.
    'msoShapeBevel':                               15,   # Bevel.
    'msoShapeBlockArc':                            20,   # Block arc.
    'msoShapeCan':                                 13,   # Can.
    'msoShapeChartPlus':                          182,   # Square divided vertically and horizontally into four quarters.
    'msoShapeChartStar':                          181,   # Square divided six parts along vertical and diagonal lines.
    'msoShapeChartX':                             180,   # Square divided into four parts along diagonal lines.
    'msoShapeChevron':                             52,   # Chevron.
    'msoShapeChord':                              161,   # Circle with a line connecting two points on the perimeter through the interior of the circle; a circle with a chord.
    'msoShapeCircularArrow':                       60,   # Block arrow that follows a curved 180-degree angle.
    'msoShapeCloud':                              179,   # Cloud shape.
    'msoShapeCloudCallout':                       108,   # Cloud callout.
    'msoShapeCorner':                             162,   # Rectangle with rectangular-shaped hole.
    'msoShapeCornerTabs':                         169,   # Four right triangles aligning along a rectangular path; four ‘snipped’ corners.
    'msoShapeCross':                               11,   # Cross.
    'msoShapeCube':                                14,   # Cube.
    'msoShapeCurvedDownArrow':                     48,   # Block arrow that curves down.
    'msoShapeCurvedDownRibbon':                   100,   # Ribbon banner that curves down.
    'msoShapeCurvedLeftArrow':                     46,   # Block arrow that curves left.
    'msoShapeCurvedRightArrow':                    45,   # Block arrow that curves right.
    'msoShapeCurvedUpArrow':                       47,   # Block arrow that curves up.
    'msoShapeCurvedUpRibbon':                      99,   # Ribbon banner that curves up.
    'msoShapeDecagon':                            144,   # Decagon.
    'msoShapeDiagonalStripe':                     141,   # Rectangle with two triangles-shapes removed; a diagonal stripe.
    'msoShapeDiamond':                              4,   # Diamond.
    'msoShapeDodecagon':                          146,   # Dodecagon
    'msoShapeDonut':                               18,   # Donut.
    'msoShapeDoubleBrace':                         27,   # Double brace.
    'msoShapeDoubleBracket':                       26,   # Double bracket.
    'msoShapeDoubleWave':                         104,   # Double wave.
    'msoShapeDownArrow':                           36,   # Block arrow that points down.
    'msoShapeDownArrowCallout':                    56,   # Callout with arrow that points down.
    'msoShapeDownRibbon':                          98,   # Ribbon banner with center area below ribbon ends.
    'msoShapeExplosion1':                          89,   # Explosion.
    'msoShapeExplosion2':                          90,   # Explosion.
    'msoShapeFlowchartAlternateProcess':           62,   # Alternate process flowchart symbol.
    'msoShapeFlowchartCard':                       75,   # Card flowchart symbol.
    'msoShapeFlowchartCollate':                    79,   # Collate flowchart symbol.
    'msoShapeFlowchartConnector':                  73,   # Connector flowchart symbol.
    'msoShapeFlowchartData':                       64,   # Data flowchart symbol.
    'msoShapeFlowchartDecision':                   63,   # Decision flowchart symbol.
    'msoShapeFlowchartDelay':                      84,   # Delay flowchart symbol.
    'msoShapeFlowchartDirectAccessStorage':        87,   # Direct access storage flowchart symbol.
    'msoShapeFlowchartDisplay':                    88,   # Display flowchart symbol.
    'msoShapeFlowchartDocument':                   67,   # Document flowchart symbol.
    'msoShapeFlowchartExtract':                    81,   # Extract flowchart symbol.
    'msoShapeFlowchartInternalStorage':            66,   # Internal storage flowchart symbol.
    'msoShapeFlowchartMagneticDisk':               86,   # Magnetic disk flowchart symbol.
    'msoShapeFlowchartManualInput':                71,   # Manual input flowchart symbol.
    'msoShapeFlowchartManualOperation':            72,   # Manual operation flowchart symbol.
    'msoShapeFlowchartMerge':                      82,   # Merge flowchart symbol.
    'msoShapeFlowchartMultidocument':              68,   # Multi-document flowchart symbol.
    'msoShapeFlowchartOfflineStorage':            139,   # Offline storage flowchart symbol.
    'msoShapeFlowchartOffpageConnector':           74,   # Off-page connector flowchart symbol.
    'msoShapeFlowchartOr':                         78,   # "Or" flowchart symbol.
    'msoShapeFlowchartPredefinedProcess':          65,   # Predefined process flowchart symbol.
    'msoShapeFlowchartPreparation':                70,   # Preparation flowchart symbol.
    'msoShapeFlowchartProcess':                    61,   # Process flowchart symbol.
    'msoShapeFlowchartPunchedTape':                76,   # Punched tape flowchart symbol.
    'msoShapeFlowchartSequentialAccessStorage':    85,   # Sequential access storage flowchart symbol.
    'msoShapeFlowchartSort':                       80,   # Sort flowchart symbol.
    'msoShapeFlowchartStoredData':                 83,   # Stored data flowchart symbol.
    'msoShapeFlowchartSummingJunction':            77,   # Summing junction flowchart symbol.
    'msoShapeFlowchartTerminator':                 69,   # Terminator flowchart symbol.
    'msoShapeFoldedCorner':                        16,   # Folded corner.
    'msoShapeFrame':                              158,   # Rectangular picture frame.
    'msoShapeFunnel':                             174,   # Funnel.
    'msoShapeGear6':                              172,   # Gear with six teeth.
    'msoShapeGear9':                              173,   # Gear with nine teeth
    'msoShapeHalfFrame':                          159,   # Half of a rectangular picture frame.
    'msoShapeHeart':                               21,   # Heart.
    'msoShapeHeptagon':                           145,   # Heptagon.
    'msoShapeHexagon':                             10,   # Hexagon.
    'msoShapeHorizontalScroll':                   102,   # Horizontal scroll.
    'msoShapeIsoscelesTriangle':                    7,   # Isosceles triangle.
    'msoShapeLeftArrow':                           34,   # Block arrow that points left.
    'msoShapeLeftArrowCallout':                    54,   # Callout with arrow that points left.
    'msoShapeLeftBrace':                           31,   # Left brace.
    'msoShapeLeftBracket':                         29,   # Left bracket.
    'msoShapeLeftCircularArrow':                  176,   # Circular arrow pointing counter-clockwise.
    'msoShapeLeftRightArrow':                      37,   # Block arrow with arrowheads that point both left and right.
    'msoShapeLeftRightArrowCallout':               57,   # Callout with arrowheads that point both left and right.
    'msoShapeLeftRightCircularArrow':             177,   # Circular arrow pointing clockwise and counter-clockwise; a curved arrow with points at both ends.
    'msoShapeLeftRightRibbon':                    140,   # Ribbon with an arrow at both ends.
    'msoShapeLeftRightUpArrow':                    40,   # Block arrow with arrowheads that point left, right, and up.
    'msoShapeLeftUpArrow':                         43,   # Block arrow with arrowheads that point left and up.
    'msoShapeLightningBolt':                       22,   # Lightning bolt.
    'msoShapeLineCallout1':                       109,   # Callout with border and horizontal callout line.
    'msoShapeLineCallout1AccentBar':              113,   # Callout with horizontal accent bar.
    'msoShapeLineCallout1BorderandAccentBar':     121,   # Callout with border and horizontal accent bar.
    'msoShapeLineCallout1NoBorder':               117,   # Callout with horizontal line.
    'msoShapeLineCallout2':                       110,   # Callout with diagonal straight line.
    'msoShapeLineCallout2AccentBar':              114,   # Callout with diagonal callout line and accent bar.
    'msoShapeLineCallout2BorderandAccentBar':     122,   # Callout with border, diagonal straight line, and accent bar.
    'msoShapeLineCallout2NoBorder':               118,   # Callout with no border and diagonal callout line.
    'msoShapeLineCallout3':                       111,   # Callout with angled line.
    'msoShapeLineCallout3AccentBar':              115,   # Callout with angled callout line and accent bar.
    'msoShapeLineCallout3BorderandAccentBar':     123,   # Callout with border, angled callout line, and accent bar.
    'msoShapeLineCallout3NoBorder':               119,   # Callout with no border and angled callout line.
    'msoShapeLineCallout4':                       112,   # Callout with callout line segments forming a U-shape.
    'msoShapeLineCallout4AccentBar':              116,   # Callout with accent bar and callout line segments forming a U-shape.
    'msoShapeLineCallout4BorderandAccentBar':     124,   # Callout with border, accent bar, and callout line segments forming a U-shape.
    'msoShapeLineCallout4NoBorder':               120,   # Callout with no border and callout line segments forming a U-shape.
    'msoShapeLineInverse':                        183,   # Line inverse.
    'msoShapeMathDivide':                         166,   # Division symbol ‘÷’.
    'msoShapeMathEqual':                          167,   # Equivalence symbol ‘=’.
    'msoShapeMathMinus':                          164,   # Subtraction symbol ‘-‘.
    'msoShapeMathMultiply':                       165,   # Multiplication symbol ‘x’.
    'msoShapeMathNotEqual':                       168,   # Non-equivalence symbol ‘≠’.
    'msoShapeMathPlus':                           163,   # Addition symbol ‘+’.
    'msoShapeMixed':                               -2,   #  Return value only; indicates a combination of the other states.
    'msoShapeMoon':                                24,   # Moon.
    'msoShapeNonIsoscelesTrapezoid':              143,   # Trapezoid with asymmetrical non-parallel sides.
    'msoShapeNoSymbol':                            19,   # "No" symbol.
    'msoShapeNotchedRightArrow':                   50,   # Notched block arrow that points right.
    'msoShapeNotPrimitive':                       138,   # Not supported.
    'msoShapeOctagon':                              6,   # Octagon.
    'msoShapeOval':                                 9,   # Oval.
    'msoShapeOvalCallout':                        107,   # Oval-shaped callout.
    'msoShapeParallelogram':                        2,   # Parallelogram.
    'msoShapePentagon':                            51,   # Pentagon.
    'msoShapePie':                                142,   # Circle (‘pie’) with a portion missing.
    'msoShapePieWedge':                           175,   # Quarter of a circular shape.
    'msoShapePlaque':                              28,   # Plaque.
    'msoShapePlaqueTabs':                         171,   # Four quarter-circles defining a rectangular shape.
    'msoShapeQuadArrow':                           39,   # Block arrows that point up, down, left, and right.
    'msoShapeQuadArrowCallout':                    59,   # Callout with arrows that point up, down, left, and right.
    'msoShapeRectangle':                            1,   # Rectangle.
    'msoShapeRectangularCallout':                 105,   # Rectangular callout.
    'msoShapeRegularPentagon':                     12,   # Pentagon.
    'msoShapeRightArrow':                          33,   # Block arrow that points right.
    'msoShapeRightArrowCallout':                   53,   # Callout with arrow that points right.
    'msoShapeRightBrace':                          32,   # Right brace.
    'msoShapeRightBracket':                        30,   # Right bracket.
    'msoShapeRightTriangle':                        8,   # Right triangle.
    'msoShapeRound1Rectangle':                    151,   # Rectangle with one rounded corner.
    'msoShapeRound2DiagRectangle':                153,   # Rectangle with two rounded corners, diagonally-opposed.
    'msoShapeRound2SameRectangle':                152,   # Rectangle with two-rounded corners that share a side.
    'msoShapeRoundedRectangle':                     5,   # Rounded rectangle.
    'msoShapeRoundedRectangularCallout':          106,   # Rounded rectangle-shaped callout.
    'msoShapeSmileyFace':                          17,   # Smiley face.
    'msoShapeSnip1Rectangle':                     155,   # Rectangle with one snipped corner.
    'msoShapeSnip2DiagRectangle':                 157,   # Rectangle with two snipped corners, diagonally-opposed.
    'msoShapeSnip2SameRectangle':                 156,   # Rectangle with two snipped corners that share a side.
    'msoShapeSnipRoundRectangle':                 154,   # Rectangle with one snipped corner and one rounded corner.
    'msoShapeSquareTabs':                         170,   # Four small squares that define a rectangular shape.
    'msoShapeStripedRightArrow':                   49,   # Block arrow that points right with stripes at the tail.
    'msoShapeSun':                                 23,   # Sun.
    'msoShapeSwooshArrow':                        178,   # Curved arrow.
    'msoShapeTear':                               160,   # Water droplet.
    'msoShapeTrapezoid':                            3,   # Trapezoid.
    'msoShapeUpArrow':                             35,   # Block arrow that points up.
    'msoShapeUpArrowCallout':                      55,   # Callout with arrow that points up.
    'msoShapeUpDownArrow':                         38,   # Block arrow that points up and down.
    'msoShapeUpDownArrowCallout':                  58,   # Callout with arrows that point up and down.
    'msoShapeUpRibbon':                            97,   # Ribbon banner with center area above ribbon ends.
    'msoShapeUTurnArrow':                          42,   # Block arrow forming a U shape.
    'msoShapeVerticalScroll':                     101,   # Vertical scroll.
    'msoShapeWave':                               103    # Wave.
}



# shape.Type
MsoShapeType = {
    'mso3DModel':          30,
    'msoAutoShape':         1,
    'msoCallout':           2,
    'msoCanvas':           20,
    'msoChart':             3,
    'msoComment':           4,
    'msoContentApp':       27,
    'msoDiagram':          21,
    'msoEmbeddedOLEObject': 7,
    'msoFormControl':       8,
    'msoFreeform':         28,
    'msoGraphic':           5,
    'msoGroup':             6,
    'msoInk':              22,
    'msoInkComment':       23,
    'msoLine':              9,
    'msoLinkedGraphic':    29,
    'msoLinked3DModel':    31,
    'msoLinkedOLEObject':  10,
    'msoLinkedPicture':    11,
    'msoMedia':            16,
    'msoOLEControlObject': 12,
    'msoPicture':          13,
    'msoPlaceholder':      14,
    'msoScriptAnchor':     18,
    'msoShapeTypeMixed':   -2,
    'msoSmartArt':         24,
    'msoTable':            19,
    'msoTextBox':          17,
    'msoTextEffect':       15,
    'msoWebVideo':         26
}


PPColorSchemeIndex = {
    'ppSchemeColorMixed': -2,
    'ppNotSchemeColor':    0,
    'ppBackground':        1,
    'ppForeground':        2,
    'ppShadow':            3,
    'ppTitle':             4,
    'ppFill':              5,
    'ppAccent1':           6,
    'ppAccent2':           7,
    'ppAccent3':           8,
    
}

MsoFillType = {
    'msoFillBackground': 5,   #Fill is the same as the background.
    'msoFillGradient':   3,   #Gradient fill.
    'msoFillMixed':     -2,   #Mixed fill.
    'msoFillPatterned':  2,   #Patterned fill.
    'msoFillPicture':    6,   #Picture fill.
    'msoFillSolid':      1,   #Solid fill.
    'msoFillTextured':   4,   #Textured fill.
} 

MsoColorType = {
    'msoColorTypeMixed': -2,
    'msoColorTypeRGB':    1,
    'msoColorTypeScheme': 2,
    'msoColorTypeCMYK':   3,
    'msoColorTypeCMS':    4,
    'msoColorTypeInk':    5
} 

MsoThemeColorIndex = {
    'msoThemeColorMixed':             -2,
    'msoNotThemeColor':                0,
    'msoThemeColorDark1':              1,
    'msoThemeColorLight1':             2,
    'msoThemeColorDark2':              3,
    'msoThemeColorLight2':             4,
    'msoThemeColorAccent1':            5,
    'msoThemeColorAccent2':            6,
    'msoThemeColorAccent3':            7,
    'msoThemeColorAccent4':            8,
    'msoThemeColorAccent5':            9,
    'msoThemeColorAccent6':           10,
    'msoThemeColorHyperlink':         11,
    'msoThemeColorFollowedHyperlink': 12,
    'msoThemeColorText1':             13,
    'msoThemeColorBackground1':       14,
    'msoThemeColorText2':             15,
    'msoThemeColorBackground2':       16
}


class ShapeDb(object):
    def __init__(self):
        self._db_file = None

    def get_db(self):
        if not self._db_file:
            self._db_file = self._load_db()
        return self._db_file

    def _load_db(self):
        import os.path
        import io
        filename = os.path.realpath(os.path.join(os.path.dirname(__file__), 'shapedb.json'))
        with io.open(filename, 'r', encoding='utf-8') as json_file:
            return json.load(json_file)
    
    def get_by_autoshape_type(self, autoshape_type):
        db = self.get_db()
        return db[str(autoshape_type)]

    def get_by_shape(self, shape):
        return self.get_by_autoshape_type(shape.AutoShapeType)

GlobalShapeDb = ShapeDb()


class LocPin(object):
    '''
    Helper class to storage the "loc pin" of shapes for various powerpoint operations.
    The "loc pin" is the pin location within the shapes that should be fixed when using shape operations (e.g. changing the size).
    '''
    def __init__(self, initial_pin=0, settings_key=None):
        # fix_height = 1 #1=top, 2=middle, 3=bottom
        # fix_width  = 1 #1=left, 2=middle, 3=right
        self.settings_key = settings_key
        if settings_key:
            self.cur_pin = settings.get(settings_key, initial_pin)
        else:
            self.cur_pin = initial_pin #index in locpins list
        self.locpins = [
            (1,1), (1,2), (1,3),
            (2,1), (2,2), (2,3),
            (3,1), (3,2), (3,3),
        ]

    '''
    fixation: The tuple that represents the locpin. (1,1) is the top-left, (3,3) is the bottom-right.
    '''
    @property
    def fixation(self):
        return self.locpins[self.cur_pin]
    @fixation.setter
    def fixation(self, value):
        self.cur_pin = self.locpins.index(value)
        if self.settings_key:
            settings[self.settings_key] = self.cur_pin

    '''
    index: The index value in the list of tuples that represent the locpin. 0 is (1,1) is top-left, 8 is (3,3) is bottom-right.
    '''
    @property
    def index(self):
        return self.cur_pin
    @index.setter
    def index(self, value):
        self.cur_pin = value
        if self.settings_key:
            settings[self.settings_key] = self.cur_pin

    def get_fractions(self):
        '''
        returns tuple (x,y) representing the pin-location within a shape.
        x,y are percentage values between 0 and 1 where
        (0,0) is the top-left pin-location and
        (1,1) is the bottom-right pin-location.
        '''
        return self.fixation[0]*0.5-0.5, self.fixation[1]*0.5-0.5

# The global locpin instance can be used to achieve a consistent behavior across powerpoint operations. E.g. it is used for both BKT size-spinners.
GlobalLocPin = LocPin(settings_key="bkt.global_loc_pin")


# ============================
# = Generic helper functions =
# ============================


def save_paste(obj, *args, **kwargs):
    for _ in range(3):
        try:
            return obj.paste(*args, **kwargs)
        except EnvironmentError:
            #wait some time to avoid EnvironmentError (run ahead bug if clipboard is busy, see https://stackoverflow.com/questions/54028910/vba-copy-paste-issues-pptx-creation)
            time.sleep(.50)
            #FIXME: maybe better way to check if clipboard actually contains "something"
    else:
        raise EnvironmentError("pasting not successfull")


def shape_is_group_child(shape):
    '''
    Test if a shape is part of a group.
    '''
    try:
        return shape.ParentGroup.Id != ""
    except SystemError:
        return False


def shape_indices_on_slide(slide, indices):
    ''' return shape-range in slide by indices '''
    return slide.Shapes.Range(System.Array[int](indices))

def last_n_shapes_on_slide(slide, n):
    ''' return last n shapes in slide as range'''
    return shape_indices_on_slide(slide, range(slide.shapes.Count + 1 -n, slide.shapes.Count + 1))

def shape_names_on_slide(slide, names):
    ''' return shape-range in slide by names '''
    #NOTE: If there are multiple shapes with the same name, only one of them is returned!
    #NOTE: This function is also looking for shapes within groups.
    return slide.Shapes.Range(System.Array[str](names))

def shapes_to_range(shapes):
    '''
    Here is another powerpoint fuckup, it is quite complicated to create a shaperange from a list of shapes.
    -> Slide.Shapes.Range(Array) either requires a list of shape indices or shape names.
    1. My first approach was to use shape names, but they are not unique and if names are replaced in VBA (to make them unique) you cannot
       restore the original name without destroying localization of names. Also, you cannot easily determine if there are multiple shapes
       with the same name as slide.Shapes.Range(Name).Count always return 1, so you have to iterate over all names before.
    2. My new approach is to use shape indices, but the shape does not have an index number, only an ID. In order to get the index number
       you have to iterate over all slide.shapes and compare with the shape your looking for. Luckily, we can leverage pythons dict for that.
    '''

    ###############
    ### Approach 2:
    #shape indices and range-function are different if shapes are within a group
    if shape_is_group_child(shapes[0]):
        all_shapes = shapes[0].ParentGroup.GroupItems
    else:
        all_shapes = shapes[0].Parent.Shapes
    #create mapping dict from all shape ids to shape indices
    shape_id2idx = {s.id: i for i,s in enumerate(all_shapes, start=1)}
    #get indices of shapes
    indices = []
    for s in shapes:
        try:
            indices.append(shape_id2idx[s.id])
        except (KeyError, EnvironmentError):
            pass #just ignore missing shapes
    #return range
    return all_shapes.Range(System.Array[int](indices))

    ###############
    ### Approach 1:
    ### Note: This approach does not properly support shapes within groups
    # import uuid
    # try:
    #     slide = shapes[0].Parent
    #     #set unique names
    #     all_names = [s.name for s in slide.shapes]
    #     orig_names = []
    #     select_names = []
    #     for i,shp in enumerate(shapes):
    #         #only replace original names if not unique as localized names will be destroyd in this step
    #         if all_names.count(shp.name) > 1:
    #             #save original name and replace name with unique one
    #             orig_names.append((i, shp.name))
    #             shp.name = str(uuid.uuid4())
    #         select_names.append(shp.name)
    #     # before return is executed, the finally statement restores original shape names
    #     return shape_names_on_slide(slide, select_names)
    # finally:
    #     #restore names
    #     if orig_names:
    #         for i,name in orig_names:
    #             shapes[i].name = name


def get_shapes_from_selection(selection):
    ''' get list of shapes from selection (considers child shape selection) '''
    # ShapeRange accessible if shape or text selected
    if selection.Type == 2 or selection.Type == 3:
        try:
            if selection.HasChildShapeRange:
                # shape selection inside grouped shapes
                return list(iter(selection.ChildShapeRange))
            else:
                return list(iter(selection.ShapeRange))
        except:
            return []
    else:
        return []

def get_slides_from_selection(selection):
    ''' get list of slides from selection '''
    # SlideRange accessible if slides, shapes or text selected
    try:
        return list(iter(selection.SlideRange))
    except:
        return []

def set_shape_zorder(shape, value=None, delta=None):
    '''
    Sets the shapes Z-Order to a specific value (if value != None) or by a specific delta (if delta != None). Delta can be negative.
    '''
    if not delta and not value:
        raise TypeError("Neither value nor delta are given!")

    if value is None:
        value = shape.ZOrderPosition + delta

    if delta is None:
        delta = value - shape.ZOrderPosition

    if delta < 0:
        direction = 3 #msoSendBackward
    elif delta > 0:
        direction = 2 #msoBringForward
    else:
        return #no change

    factor = delta/abs(delta)
    #simulation of do-while-loop
    while True:
        prev_zorder = shape.ZOrderPosition
        shape.ZOrder(direction)
        if prev_zorder == shape.ZOrderPosition:
            break
            #no change in position
        if factor*shape.ZOrderPosition >= factor*value:
            break
            #zorder reached


def transfer_textrange(from_textrange, to_textrange):
    '''
    This function copy-pastes a textrange into another textrange. The standard textrange.copy() function works fine,
    but the textrange.paste() via code does replace ThemeColors with RGB values (Note: via GUI this works fine).
    So this function manually copies color values after copying the textrange.
    '''
    from_textrange.Copy()
    # to_textrange.Paste()
    save_paste(to_textrange)

    def copy_color(from_obj, to_obj):
        if from_obj.ObjectThemeColor != 0:
            to_obj.ObjectThemeColor = from_obj.ObjectThemeColor
            to_obj.Brightness = from_obj.Brightness

    for i,run in enumerate(from_textrange.Runs(), start=1):
        try:
            to_font = to_textrange.Runs(i).Font
        except ValueError:
            #ValueError: Der Index in der angegebenen Sammlung ist außerhalb des zulässigen Bereichs.
            continue
        copy_color(run.Font.Fill.ForeColor, to_font.Fill.ForeColor)
        copy_color(run.Font.Fill.BackColor, to_font.Fill.BackColor)
        copy_color(run.Font.Line.ForeColor, to_font.Line.ForeColor)
        copy_color(run.Font.Line.BackColor, to_font.Line.BackColor)


def replicate_shape(shape, force_textbox=False):
    '''
    This function replicates a shape, which is similar to shape.Duplicate() but instead a new shape is created.
    The duplicate function throws a ComException if the duplicate is used (e.g. merged, deleted) afterwards due to pending event handling.
    '''
    slide = shape.Parent
    # Note: Placeholder can be table, chart, diagram, smartart, picture, whatever...
    shape_type = shape.Type
    if shape_type == MsoShapeType['msoPlaceholder']:
        shape_type = shape.PlaceholderFormat.ContainedType
    if force_textbox or shape_type == MsoShapeType['msoTextBox']:
        new_shape = slide.shapes.AddTextbox(
            1, #msoTextOrientationHorizontal
            shape.Left, shape.Top, shape.Width, shape.Height)
        new_shape.AutoShapeType = shape.AutoShapeType
    elif shape_type == MsoShapeType["msoAutoShape"]:
        new_shape = slide.shapes.AddShape(
            shape.AutoShapeType,
            shape.Left, shape.Top, shape.Width, shape.Height)
    elif shape_type == MsoShapeType["msoCallout"]:
        new_shape = slide.shapes.AddCallout(
            shape.Callout.Type,
            shape.Left, shape.Top, shape.Width, shape.Height)
    else:
        raise ValueError("replication only possible with autoshapes and textboxes")
    
    #replicate shape properties
    if shape.VerticalFlip != new_shape.VerticalFlip:
        new_shape.Flip(1) #msoFlipVertical
    if shape.HorizontalFlip != new_shape.HorizontalFlip:
        new_shape.Flip(0) #msoFlipHorizontal

    for i in range(1,shape.adjustments.count+1):
        try:
            new_shape.adjustments.item[i] = shape.adjustments.item[i]
        except:
            continue

    new_shape.Rotation = shape.Rotation

    #copy all formatting
    shape.PickUp()
    new_shape.Apply()

    #copy text
    # shape.TextFrame2.TextRange.Copy()
    # new_shape.TextFrame2.TextRange.Paste()
    transfer_textrange(shape.TextFrame2.TextRange, new_shape.TextFrame2.TextRange)

    #ensure correct size and position (size may change due to AutoSize, Flip can change position)
    new_shape.Height = shape.Height
    new_shape.Width  = shape.Width
    new_shape.Top    = shape.Top
    new_shape.Left   = shape.Left
    
    return new_shape


def convert_text_into_shape(shape):
    '''
    This function converts text into a shape. This is very useful for icon fonts. If the shape has a background, the text is cut out of the shape.
    We use the standard merge functions from powerpoint, which are buggy in some situation: If a special shape with adjustments is used, the 
    converted text is not at the exact same position as the original text. This is very annoying for the cut-out function. No workaround found :(

    ### MsoMergeCmd:
    msoMergeCombine     2   Creates a new shape from selected shapes. If the selected shapes overlap, the area where they overlap is cut out, or discarded.
    msoMergeFragment    5   Breaks a shape into smaller parts or create new shapes from intersecting lines or from shapes that overlap.
    msoMergeIntersect   3   Forms a new closed shape from the area where selected shapes overlap, eliminating non-overlapping areas.
    msoMergeSubtract    4   Creates a new shape by subtracting from the primary selection the areas where subsequent selections overlap.
    msoMergeUnion       1   Creates a new shape from the perimeter of two or more overlapping shapes. The new shape is a set of all the points from the original shapes.
    '''
    slide = shape.Parent

    #find shape index
    for index, shp in enumerate(slide.shapes, start=1):
        if shape.id == shp.id:
            shape_index = index
            break
    else:
        #shape not found
        return

    #total shapes
    shape_count = slide.shapes.count
    #original shape name
    orig_name = shape.Name

    #convert actual text into shape
    if shape.Fill.visible == 0:
        #turn off line as it prohibts conversion
        shape.Line.visible = 0

        #add temporary shape
        slide.shapes.AddShape( MsoAutoShapeType['msoShapeRectangle']
            , -10, 0, 10, 10)
        
        #select shape and temporary shape
        shapes = shape_indices_on_slide(slide, [shape_index, shape_count+1])
        shapes.MergeShapes(4, shape)
    
    #cut text out of shape
    elif shape.TextFrame2.HasText:
        # first approach: duplicate shape, remove fill+line, and text from original shape,
        #                 but than MergeShape fails with ComException. It seems that events
        #                 need to be processed before. Workaround: Delay MergeShape in a Thread,
        #                 but than we cannot return the resulting shape.
        # new approach: create new shape and copy all relevant formatting

        #ensure autosize is off
        shape.TextFrame2.AutoSize = 0 #ppAutoSizeNone

        #duplicate shape without using Duplicate()
        text_shape = replicate_shape(shape, True)

        #remove fill and line
        text_shape.Fill.visible=0
        text_shape.Line.visible=0

        #delete text from original shape
        shape.TextFrame2.DeleteText()

        #select shape and text shape
        shapes = shape_indices_on_slide(slide, [shape_index, shape_count+1])
        shapes.MergeShapes(4, shape)
    
    #nothing to do
    else:
        return shape

    new_shape = shape_indices_on_slide(slide, [shape_index])[1]
    new_shape.LockAspectRatio = -1
    new_shape.Name = orig_name
    return new_shape



# ====================
# = Tag helper class =
# ====================

class TagHelper(object):
    '''
    Helper to check if shape has a tag, get all tag values as dict or set tags from dict.
    '''
    @staticmethod
    def get_dict_from_tags(obj_tags):
        '''
        Convert all shape/slide tags to a python dictionary.
        '''
        d = dict()
        for i in range(obj_tags.count):
            d[obj_tags.name(i+1)] = obj_tags.value(i+1)
        return d

    @staticmethod
    def set_tags_from_dict(tags_dict, obj_tags):
        '''
        Set shape tags based on a python dictionary.
        '''
        for k,v in tags_dict.items():
            obj_tags.add(k,v)

    @staticmethod
    def has_tag(obj, tag_name, check_value=None):
        '''
        Test if shape has specified tag (with value)
        '''
        try:
            if check_value is not None:
                return obj.Tags(tag_name) == check_value
            else:
                return obj.Tags(tag_name) != ''
        except: #EnvironmentError
            #Shape.Tags throws COMException for SmartArt child-shapes
            return False

    @staticmethod
    def get_tag(obj, tag_name, default=None, attr_type=None):
        '''
        Get value of tag and try to convert attribute type, otherwise return default
        '''
        try:
            value = obj.Tags(tag_name)
            if value == '':
                return default
            if type(attr_type) == type:
                return attr_type(value)
            return value
        except:
            return default



# ======================
# = Color helper class =
# ======================

class PPTColor(object):
    '''
    This class represents a single color similar to the powerpoint color object.
    Helper methods provided to pickup or apply color from powerpoint color object, 
    and to export color as tuple.
    '''

    COLOR_NONE = 0 #convinience for visible=0
    COLOR_THEME = 1
    COLOR_RGB = 2

    @classmethod
    def from_color_obj(cls, color_obj):
        return cls().pickup_from_color_obj(color_obj)

    @classmethod
    def from_color_tuple(cls, color_tuple):
        if not color_tuple:
            return cls(cls.COLOR_NONE)
        elif color_tuple[0] == cls.COLOR_RGB:
            return cls(cls.COLOR_RGB, color_rgb=color_tuple[1])
        elif color_tuple[0] == cls.COLOR_THEME:
            return cls(cls.COLOR_THEME, color_index=color_tuple[1], brightness=color_tuple[2]/100)

    @classmethod
    def new_rgb(cls, color_rgb):
        return cls(cls.COLOR_RGB, color_rgb=color_rgb)
    
    @classmethod
    def new_theme(cls, color_index, brightness=0, color_rgb=None, shade_index=None):
        return cls(cls.COLOR_THEME, color_index, brightness, color_rgb, shade_index)


    def __init__(self, color_type=COLOR_RGB, color_index=None, brightness=None, color_rgb=None, shade_index=None):
        self.color_type  = color_type
        self.color_index = color_index
        self.brightness  = brightness
        self.color_rgb   = color_rgb
        self.shade_index = shade_index #-1=not yet defined, set index on next update_from_context; None=ignore, only use brightness
    
    def __eq__(self, other):
        if isinstance(other, PPTColor):
            return self.get_color_tuple() == other.get_color_tuple()
        return False
    
    def __ne__(self, other):
        return not self.__eq__(other)
    
    def __nonzero__(self):
        return self.color_type != self.COLOR_NONE
    __bool__ = __nonzero__

    def pickup_from_color_obj(self, color_obj):
        if color_obj.Type == MsoColorType['msoColorTypeScheme']:
            self.color_type  = self.COLOR_THEME
            self.color_index = color_obj.ObjectThemeColor
            # self.color_index = color_obj.SchemeColor
            self.brightness  = color_obj.Brightness
            self.color_rgb   = color_obj.RGB
            self.shade_index = -1 #-1: shade index will be set on first update
        else:
            self.color_type  = self.COLOR_RGB
            self.color_index = None
            self.brightness  = None
            self.color_rgb   = color_obj.RGB
            self.shade_index = None
        return self
    
    def apply_to_color_obj(self, color_obj):
        if self.color_type == self.COLOR_THEME:
            color_obj.ObjectThemeColor = self.color_index
            # color_obj.SchemeColor = self.color_index
            color_obj.Brightness = self.brightness
        elif self.color_type == self.COLOR_RGB:
            color_obj.RGB = self.color_rgb

    def get_color_tuple(self):
        if self.color_type == self.COLOR_THEME:
            return (self.COLOR_THEME, self.color_index, int(100*self.brightness)) #convert brightness to int to avoid floating point comparison problems
        elif self.color_type == self.COLOR_RGB:
            return (self.COLOR_RGB, self.color_rgb)
        else:
            return None


class ColorHelper(object):
    '''
    So, puhhh, how to start, ... colors and color indices are a huge mess in PowerPoint (and Office in general).
    Here is a good article about the mess in Word: http://www.wordarticles.com/Articles/Colours/2007.php
    Here is an article about the shade indices: https://stackoverflow.com/questions/21142732/how-to-get-the-rgb-long-values-from-powerpoint-color-palette
    Basically, a color object has 2 attributes, ObjectThemeColor and SchemeColor.
    ObjectThemeColor goes from index 1 to 16. The default color palette is using 5-10 and 13-16 (11+12 are hyperlink colors).
    SchemeColor goes from 1 to 8, where 7+8 are Hyperlink colors. The ObjectThemeColor indices 13-16 are mappes to 1-4 in SchemeColor internally, not in order, of course.
    In order to get the correct RGB value, you need to use 2 different functions:
      - ColorScheme(index) gets the correct value for indices 1-4 (resp. the mapped values of indices 13-16). But ColorScheme is not defined for values >8.
      - ThemeColorScheme(index) gets the correct value for indices 5-12. ThemeColorScheme is not defined for value >12. For indices 1-4 it will (at least for some themes)
        provide different RGB values than ColorScheme.
    Hint: We could only use the ObjectThemeColor attribute with indices 1-10 and live a happy life, but then the default color palette would not indicate the correct "checked"
    status for the color indices 1-4!

    No coming to theme color shades. The brightness differs depending on HSL-Luminosity of the theme color. So in order to save and restore the same shade across different
    themes, we need to get the index that maps to the brightness. In order to get the RGB value, we need to adjust the theme color by a brightness factor.

    This class provides helper functions to handle this mess.
    '''

    _theme_color_indices = [14,13,16,15, 5,6,7,8,9,10] #powerpoint default color picker is using IDs 5-10 and 13-16
    _theme_color_names = ['Hintergrund 1', 'Text 1', 'Hintergrund 2', 'Text 2', 'Akzent 1', 'Akzent 2', 'Akzent 3', 'Akzent 4', 'Akzent 5', 'Akzent 6']
    _theme_color_shades = [
        # depending on HSL-Luminosity, different brightness-values are used
        # brightness-values = percentage brighter  (darker if negative)
        [[0],           [ 50,   35,  25,  15,   5] ],
        [range(1,20),   [ 90,   75,  50,  25,  10] ],
        [range(20,80),  [ 80,   60,  40, -25, -50] ],
        [range(80,100), [-10,  -25, -50, -75, -90] ],
        [[100],         [ -5,  -15, -25, -35, -50] ]
    ] #using int values to avoid floating point comparison problems

    _color_class = namedtuple("ThemeColor", "rgb brightness shade_index theme_index name")


    ### internal helper methods ###

    @classmethod
    def _theme_color_index_2_color_scheme_index(cls, index):
        mapping = {
            13: 2,
            14: 1,
            15: 4,
            16: 3,
        }
        return mapping[index]
    
    @classmethod
    def _get_color_from_theme_index(cls, context, index): #expect MsoThemeColorSchemeIndex
        if index > 12:
            try:
                ColorScheme = context.slide.ColorScheme
            except: #if presentation has no slides, the above will fail
                ColorScheme = context.presentation.SlideMaster.ColorScheme
            index = cls._theme_color_index_2_color_scheme_index(index)
        else:
            try:
                ColorScheme = context.slide.ThemeColorScheme
            except: #if presentation has no slides, the above will fail
                ColorScheme = context.presentation.SlideMaster.Theme.ThemeColorScheme
        
        return ColorScheme(index)

    @classmethod
    def _get_factors_for_rgb(cls, color_rgb):
        r,g,b = algorithms.get_rgb_from_ole(color_rgb)
        l = round( algorithms.get_brightness_from_rgb(r,g,b) / 255. *100 )
        return [factors[1] for factors in cls._theme_color_shades if l in factors[0]][0]

        ### old method with Drawing:
        # color = Drawing.ColorTranslator.FromOle(color_rgb)
        # l = color.GetBrightness()*255
        # return [factors[1] for factors in cls._theme_color_shades if round(l) in factors[0]][0]
    
    @classmethod
    def _get_color_name(cls, index, shade_index, brightness):
        theme_col_name = cls._theme_color_names[cls._theme_color_indices.index(index)]
        if brightness != 0:
            return "{}, {} {:.0%}".format(theme_col_name, "heller" if brightness > 0 else "dunkler", abs(brightness))
        return theme_col_name


    ### external functions for theme colors and shades ###

    @classmethod
    def adjust_rgb_brightness(cls, color_rgb, brightness):
        if brightness == 0:
            return color_rgb
        
        # load python color transformation library
        import colorsys
        # split rgb color in r,g,b
        r,g,b = algorithms.get_rgb_from_ole(color_rgb)
        # split r,g,b in h,l,s
        h,l,s = colorsys.rgb_to_hls(r/255.,g/255.,b/255.)
        # adjust l value
        if brightness > 0:
            l += (1.-l)*brightness
        else:
            l += l*brightness
        # convert back into r,g,b
        r,g,b = colorsys.hls_to_rgb(h,l,s)
        # return rgb color
        return algorithms.get_ole_from_rgb(round(r*255),round(g*255),round(b*255))

        ### old method with Drawing:
        # # split rgb color in r,g,b
        # color = Drawing.ColorTranslator.FromOle(color_rgb)
        # r,g,b = color.R, color.G, color.B
        # # apply brightness factor
        # if brightness < 0:
        #     r = round(r * (1+brightness))
        #     g = round(g * (1+brightness))
        #     b = round(b * (1+brightness))
        # else:
        #     r = round(r + (255.-r)*brightness)
        #     g = round(g + (255.-g)*brightness)
        #     b = round(b + (255.-b)*brightness)
        # # store color rgb
        # color = Drawing.Color.FromArgb(r, g, b);
        # return Drawing.ColorTranslator.ToOle(color)
    
    @classmethod
    def get_brightness_from_shade_index(cls, color_rgb, shade_index):
        factors = cls._get_factors_for_rgb(color_rgb)
        return factors[shade_index]/100.0
    
    @classmethod
    def get_shade_index_from_brightness(cls, color_rgb, brightness):
        factors = cls._get_factors_for_rgb(color_rgb)
        return factors.index(int(100*brightness))
    
    @classmethod
    def get_theme_index(cls, i):
        return cls._theme_color_indices[i%10]

    @classmethod
    def get_theme_color(cls, context, index, brightness=0, shade_index=None):
        color_rgb = cls._get_color_from_theme_index(context, index).RGB
        if shade_index is not None:
            brightness = cls.get_brightness_from_shade_index(color_rgb, shade_index)
        elif brightness != 0:
            try:
                shade_index = cls.get_shade_index_from_brightness(color_rgb, brightness)
            except ValueError:
                shade_index = None
        
        color_rgb = cls.adjust_rgb_brightness(color_rgb, brightness)
        
        return cls._color_class(color_rgb, brightness, shade_index, index, cls._get_color_name(index, shade_index, brightness))
    
    @classmethod
    def get_theme_colors(cls):
        return zip(cls._theme_color_indices, cls._theme_color_names)


    ### external functions for recent colors ###

    @classmethod
    def get_recent_color(cls, context, index):
        return context.presentation.ExtraColors(index)

    @classmethod
    def get_recent_colors_count(cls, context):
        return context.presentation.ExtraColors.Count



# =========================================
# = Custom BKT tags stored in JSON format =
# =========================================

class BKTTag(object):
    '''
    Use shape tags using with-statement and item-notation. Tag values are stored as json data.
    '''
    TAG_NAME = "BKT"

    def __init__(self, tags):
        self.tags = tags
        self.data = {}
        
    def load(self):
        try:
            tag_data = self.tags.Item(self.TAG_NAME)
            if not tag_data or tag_data == '':
                self.data = {}
            else:
                self.data = json.loads(tag_data)
        except:
            self.data = {}
        
    def save(self):
        try:
            if len(self.data) > 0:
                tag_data = json.dumps(self.data, ensure_ascii=False)
                self.tags.Add(self.TAG_NAME,tag_data)
            else:
                self.tags.Delete(self.TAG_NAME)
        except:
            #import traceback #debugging only
            #traceback.print_exc()
            raise AttributeError("error saving json to tags")
    
    def remove(self):
        self.data = {}

    def get(self, arg, default=None):
        return self.data.get(arg, default)

    def __enter__(self):
        self.load()
        return self

    def __exit__(self, cls, value, traceback):
        self.save()
    
    def __getitem__(self, arg):
        ''' access ribbon-attributes in dict-style, e.g. button['label'] '''
        return self.data[arg]
    
    def __setitem__(self, arg, value):
        ''' access ribbon-attributes in dict-style, e.g. button['label'] = 'foo' '''
        if arg is None or value is None:
            raise ValueError("value cannot be none")
        
        self.data[arg] = value
    
    def __delitem__(self, arg):
        ''' access ribbon-attributes in dict-style, e.g. del button['label'] '''
        del self.data[arg]



# ======================
# = Slide content size =
# ======================

def slidemaster_from_obj(obj):
    ''' get slide master object from any object (presentation, slide, shape, layout, etc.) '''
    #obj.parent:
    # -> if obj.parent.Master exists, then obj was shape
    # -> if obj.parent.CustomLayouts exists, then obj was custom layout or shape on slidemaster
    #obj.parent.parent
    # -> if obj.parent.parent.design.SlideMaster exists, then obj was shape on custom layout
    #obj.parent.parent.parent
    # -> fallback to presentation
    master_obj = obj
    attrs = [None, "parent", "parent", "design", "parent"]
    for attr in attrs:
        if attr:
            master_obj = getattr(master_obj, attr)
        
        if hasattr(master_obj, "CustomLayouts"):
            #obj is slide master
            return master_obj

        try:
            #obj is presentation or design
            return master_obj.SlideMaster
        except AttributeError:
            pass

        try:
            #obj is slide 
            return master_obj.Master
        except AttributeError:
            pass
    
    else:
        raise AttributeError("%s cannot be converted to slidemaster" % obj)

def content_size_from_master(slidemaster):
    ''' get size of content area (i.e. big text field of standard layout) from slide master '''
    try:
        return next([shape.left, shape.top, shape.width, shape.height] for shape in iter(slidemaster.Shapes) if shape.type == 14 and shape.Placeholderformat.type == 2)
    except StopIteration:
        return 0, 0, slidemaster.Width, slidemaster.Height
        # page_setup = slidemaster.Parent.PageSetup
        # return 0, 0, page_setup.SlideWidth, page_setup.SlideHeight

def slide_content_size(any_obj):
    ''' get size of content area (i.e. big text field of standard layout) from any object (slide, presentation, shape, etc.) '''
    return content_size_from_master(slidemaster_from_obj(any_obj))


BKT_CONTENTAREA = "BKT_CONTENTAREA"

class ContentAreaTags(BKTTag):
    TAG_NAME = BKT_CONTENTAREA

    @property
    def is_area_set(self):
        return "contentarea_left" in self.data

    def get_area(self):
        return self.data["contentarea_left"], self.data["contentarea_top"], self.data["contentarea_width"], self.data["contentarea_height"]

class ContentArea(object):
    @staticmethod
    def isset_contentarea(presentation):
        ''' test if custom content area is defined for given presentation '''
        if presentation.Tags.Item(BKT_CONTENTAREA) != '':
            with ContentAreaTags(presentation.Tags) as tags:
                return tags.is_area_set
        else:
            return False

    @staticmethod
    def define_contentarea(presentation, shape):
        ''' define custom content area for given presentation '''
        with ContentAreaTags(presentation.Tags) as tags:
            tags["contentarea_left"]   = float(shape.left)
            tags["contentarea_top"]    = float(shape.top)
            tags["contentarea_width"]  = float(shape.width)
            tags["contentarea_height"] = float(shape.height)
        #shape.Delete()

    @staticmethod
    def reset_contentarea(presentation):
        ''' delete custom content area from given presentation '''
        presentation.tags.Delete(BKT_CONTENTAREA)

    @staticmethod
    def read_contentarea(presentation):
        ''' get custom content area from given presentation '''
        with ContentAreaTags(presentation.Tags) as tags:
            if tags.is_area_set:
                return tags.get_area() #left,top,width,height
            else:
                return slide_content_size(presentation) #left,top,width,height


# =========================================
# = Iterator for "subshapes" & textframes =
# =========================================

class SubShapeIterator(object):
    '''
    Iterate through shapes of different types and return every shapes "subhsapes", e.g. group shapes or table cells
    arg 'from_selection': If shapes are not from a selection (e.g. iterate all shapes of a slide), set this to False to disable selected table cells detection,
                        otherwise not all table cells are iterated at least in the rare case that a table is the only shape on a slide.
    '''

    def __init__(self, shapes, from_selection=True):
        #Ensure list
        if type(shapes) != list:
            shapes = list(iter(shapes))
        
        self.only_selected_table_cells = False

        #If cells within a table are selected, function should only iterate selected cells. If the whole table is selected but no other shape, all cells are selected.
        # if from_selection and len(shapes) == 1 and _get_shp_type(shapes[0]) == MsoShapeType['msoTable']:
        if from_selection and len(shapes) == 1 and shapes[0].HasTable == -1:
            self.only_selected_table_cells = True

        self.shapes = shapes
    
    def __iter__(self):
        for shape in self.shapes:
            shp_type = self._get_shp_type(shape)

            # Iterate each group item
            if shp_type == MsoShapeType['msoGroup']:
                generator = self._iter_group(shape)
            
            # Iterate each smart art node
            elif shp_type == MsoShapeType['msoSmartArt']:
                generator = self._iter_smartart(shape)
            
            # Iterate each chart/diagram shape
            elif shp_type == MsoShapeType['msoChart'] or shp_type == MsoShapeType['msoDiagram']:
                generator = self._iter_chart(shape)
            
            # Iterate each table cell
            elif shp_type == MsoShapeType['msoTable']:
                generator = self._iter_table(shape)
            
            else:
                generator = self._iter_default(shape)
            
            for obj in generator:
                yield obj
    
    def _get_shp_type(self, shape):
        #For table cells Type is not implemented and will throw an error
        try:
            # Note: Placeholder can be table, chart, diagram, smartart, picture, whatever...
            if shape.Type == MsoShapeType['msoPlaceholder']:
                return shape.PlaceholderFormat.ContainedType
            return shape.Type
        except:
            return None
    
    def _iter_group(self, shape):
        for shp in shape.GroupItems:
            yield shp
    
    def _iter_smartart(self, shape):
        for shp in shape.GroupItems:
            yield shp
    
    def _iter_chart(self, shape):
        yield shape
    
    def _iter_default(self, shape):
        yield shape
    
    def _iter_table(self, shape):
        for row in shape.table.rows:
            for cell in row.cells:
                if not self.only_selected_table_cells or cell.Selected:
                    yield cell.Shape


def iterate_shape_subshapes(shapes, from_selection=True, filter_method=lambda shp: True, getter_method=lambda shp: shp):
    ''' Function to create sub shape iterator '''
    return SubShapeIterator(shapes, from_selection)


class TextframeIterator(SubShapeIterator):
    '''
    Iterate through shapes of different types and return every shapes textframe
    '''
    
    def _iter_group(self, shape):
        for shp in shape.GroupItems:
            if shp.HasTextFrame:
                yield shp.TextFrame2
    
    def _iter_default(self, shape):
        if shape.HasTextFrame:
            yield shape.TextFrame2
    
    def _iter_table(self, shape):
        for row in shape.table.rows:
            for cell in row.cells:
                if not self.only_selected_table_cells or cell.Selected:
                    yield cell.Shape.TextFrame2
    
    def _iter_smartart(self, shape):
        # Iterate over nodes instead of GroupItems to get only textframes, but node has no HasTextFrame property!
        for node in shape.SmartArt.AllNodes:
            yield node.TextFrame2
    
    def _iter_chart(self, shape):
        yield shape.Chart.Format.TextFrame2
        yield shape.Chart.ChartArea.Format.TextFrame2
        if shape.Chart.HasTitle:
            yield shape.Chart.ChartTitle.Format.TextFrame2
        if shape.Chart.HasLegend:
            yield shape.Chart.Legend.Format.TextFrame2
        if shape.Chart.HasDataTable:
            yield shape.Chart.DataTable.Format.TextFrame2


def iterate_shape_textframes(shapes, from_selection=True):
    ''' Function to create textframe iterator '''
    return TextframeIterator(shapes, from_selection)



# ===============================
# = Generic class for rectangle =
# ===============================

class BoundingFrame(object):
    '''
    Helper class to simulate a rectangle and create a bounding frame from shape list.
    '''
    def __init__(self, slide=None, contentarea=False):
        self.left=0
        self.top=0
        self.width=100
        self.height=100
        self.rotation=0
        
        if slide != None:
            slidemaster = slidemaster_from_obj(slide)
            if contentarea:
                self.left, self.top, self.width, self.height = content_size_from_master(slidemaster)
            else:
                self.width  = slidemaster.Width
                self.height = slidemaster.Height
                # page_setup  = slide.parent.PageSetup
                # self.width  = page_setup.SlideWidth
                # self.height = page_setup.SlideHeight
    
    @classmethod
    def from_rect(cls, left, top, width, height):
        bf = BoundingFrame()
        bf.left = left
        bf.top = top
        bf.width = width
        bf.height = height
        return bf

    @classmethod
    def from_shape(cls, shape):
        return cls.from_rect(shape.left, shape.top, shape.width, shape.height)

    @classmethod
    def from_shapes(cls, shapes):
        from .wrapper import wrap_shapes
        shapes = wrap_shapes(shapes)

        bf = BoundingFrame()

        shapes.sort(key=lambda shape: shape.visual_x)
        bf.left = shapes[0].visual_x

        shapes.sort(key=lambda shape: shape.visual_y)
        bf.top = shapes[0].visual_y

        shapes.sort(key=lambda shape: shape.visual_x1, reverse=True)
        bf.width = shapes[0].visual_x1 - bf.left

        shapes.sort(key=lambda shape: shape.visual_y1, reverse=True)
        bf.height = shapes[0].visual_y1 - bf.top

        return bf



# ==========================
# = Group helper functions =
# ==========================

class GroupManager(object):
    '''
    This is a helper class to handle more complicated group actions without affecting the groups name, tags and rotation
    '''
    def __init__(self, group, additional_attrs=[]):
        self._group   = group
        self._ungroup = None

        self._name = group.name
        self._tags = TagHelper.get_dict_from_tags(group.tags)
        self._rotation      = group.Rotation
        self._fliph         = group.HorizontalFlip
        self._flipv         = group.VerticalFlip
        self._zorder        = group.ZOrderPosition
        self._aspectratio   = group.LockAspectRatio

        self._attr = {n:getattr(group, n) for n in additional_attrs}

        self._ungroup_prepared = False

    def __getattr__(self, name):
        # provides easy access to shape properties
        return getattr(self._group, name)

    # def __setattr__(self, name, value):
    #     # provides easy access to shape properties
    #     setattr(self._group, name, value)

    @property
    def child_items(self):
        '''
        Get group child items as list, depending if group is already ungrouped or not
        '''
        if self._group:
            return list(iter(self._group.GroupItems))
        else:
            return list(iter(self._ungroup))
    
    @property
    def shape(self):
        '''
        Get group shape. Throws error if already ungrouped
        '''
        if not self._group:
            raise SystemError("not a group")
        return self._group

    def select(self, replace=True):
        '''
        Either select group or all child shapes (if ungrouped).
        Due to random error when selecting, try a second time without replace parameter if first time fails.
        '''
        try:
            if self._group:
                self._group.select(replace=replace)
            else:
                self._ungroup.select(replace=replace)
        except EnvironmentError:
            # Select(replace=False) sometimes throws "Invalid request.  To select a shape, its view must be active.", e.g. right after duplicating the shape
            if self._group:
                self._group.select()
            else:
                self._ungroup.select()

    def refresh(self):
        '''
        Refresh the group, means ungroup and regroup in order to fix corruption,
        e.g. if child shape is duplicated it is not properly added to the group until this method is performed
        '''
        self.ungroup()
        self.regroup()

    def prepare_ungroup(self):
        '''
        Method is executed right before ungroup action in order to set rotation to 0.
        '''
        self._group.rotation = 0
        if self._fliph and self._group.HorizontalFlip: #avoid double flip if function if called twice
            self._group.Flip(0) #msoFlipHorizontal
        if self._flipv and self._group.VerticalFlip:
            self._group.Flip(1) #msoFlipVertical
        self._ungroup_prepared = True

    def post_regroup(self):
        '''
        Method is executed right after regroup action in order to set rotation to original rotation.
        '''
        self._group.rotation = self._rotation
        if self._fliph != self._group.HorizontalFlip:
            self._group.Flip(0) #msoFlipHorizontal
        if self._flipv != self._group.VerticalFlip:
            self._group.Flip(1) #msoFlipVertical
        self._ungroup_prepared = False

    def ungroup(self, prepare=True):
        '''
        Perform ungroup with rotation=0. If prepare=False, prepare-method is not called and rotation is not set to 0.
        '''
        if not self._group:
            raise SystemError("not a group")

        if prepare and not self._ungroup_prepared:
            self.prepare_ungroup()
        self._ungroup = self._group.ungroup()
        self._group = None
        return self
    
    def regroup(self, new_shape_range=None):
        '''
        Perform regroup (actually group) and reset all attributes (name, tags, rotation) to original values.
        If new_shape_range is given, the stored shape-range from ungroup is replaced with the given shape-range.
        '''
        self._ungroup = new_shape_range or self._ungroup
        if not self._ungroup:
            raise SystemError("not ungrouped")

        self._group = self._ungroup.group()
        self._ungroup = None

        #restore name
        self._group.name = self._name
        #restore tags
        TagHelper.set_tags_from_dict(self._tags, self._group.tags)
        #restore additional parameter, e.g. width in process chevrons example
        for k,v in self._attr.items():
            setattr(self._group, k, v)
        #restore zorder
        set_shape_zorder(self._group, value=self._zorder)
        #restore lock aspect ration
        self._group.LockAspectRatio = self._aspectratio
        #call post_regroup to reset rotation
        if self._ungroup_prepared:
            self.post_regroup()
        return self
    
    def add_child_items(self, shapes):
        '''
        Add shape(s) to group without modifying the group.
        '''
        if not self._group:
            raise SystemError("not a group")
        
        #store position of first shape in group
        shape_to_restore_pos = self.shape.GroupItems[1]
        orig_left, orig_top = shape_to_restore_pos.left, shape_to_restore_pos.top
        #add shapes to temporary group
        temp_grp = shapes_to_range([self.shape]+shapes).group()
        #rotate original group to 0
        temp_grp.rotation = - self._rotation
        temp_grp.ungroup()
        #create new group and reset rotation
        self.ungroup()
        self.regroup(new_shape_range=shapes_to_range(self.child_items+shapes))
        #restore position
        self.shape.left -= shape_to_restore_pos.left-orig_left
        self.shape.top  -= shape_to_restore_pos.top-orig_top

        ### Simple method without considering rotation:
        # self.ungroup(prepare=False)
        # self.regroup(new_shape_range=shapes_to_range(self.child_items+shapes))
        return self

    def recursive_ungroup(self):
        '''
        Ungroup the group and all its sub-groups until no more groups exist.
        '''
        if not self._group:
            raise SystemError("not a group")

        def _ungroup(shape_range):
            for s in shape_range:
                if s.Type == MsoShapeType["msoGroup"]:
                    for s2 in _ungroup(s.ungroup()):
                        yield s2
                else:
                    yield s

        self._ungroup = shapes_to_range( list(_ungroup(self._group.ungroup())) )
        self._group = None
        return self



# =================
# = Locale helper =
# =================

language_id_to_locale = {
    1031: 'de_DE', #"Deutsch",
    3079: 'de_AT', #"Deutsch (Österreich)",
    1040: 'it', #"Italienisch",
    1036: 'fr', #"Französisch",
    3082: 'es', #"Spanisch",
    1049: 'ru', #"Russisch",
    1029: 'cz', #"Tschechisch",
    1030: 'dk', #"Dänisch",
    1043: 'nl', #"Holländisch",
    1045: 'pl', #"Polnisch",
    2070: 'pt', #"Portugisisch",
    1053: 'se', #"Schwedisch",
    1055: 'tr', #"Türkisch",
    1033: 'en_US', #"US English",
    2057: 'en_UK', #"UK English",
}

@contextmanager
def override_locale(language_id):
    '''
    Temporarily change the python locale based on msoLanguageId
    '''
    import locale
    category = locale.LC_ALL
    locale_string = language_id_to_locale.get(language_id, 'en')
    prev_locale_string = locale.getlocale(category)[0]
    locale.setlocale(category, locale_string)
    yield
    locale.setlocale(category, prev_locale_string)


