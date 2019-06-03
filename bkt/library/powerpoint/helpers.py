# -*- coding: utf-8 -*-

# DO NOT REMOVE REFERENCE
# reference is used by other modules
import clr
clr.AddReference("Microsoft.Office.Interop.PowerPoint")
import Microsoft.Office.Interop.PowerPoint as PowerPoint

import json # required for tags
from bkt import settings

ptToCmFactor = 2.54 / 72;
def pt_to_cm(pt):
    return float(pt) * ptToCmFactor;
def cm_to_pt(cm):
    return float(cm) / ptToCmFactor;

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
    'msoLinkedOLEObject':  10,
    'msoLinkedPicture':    11,
    'msoMedia':            16,
    'msoOLEControlObject': 12,
    'msoPicture':          13,
    'msoPlaceholder':      14,
    'msoScriptAnchor':     18,
    'msoShapeTypeMixed':    2,
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


'''
Helper class to storage the "loc pin" of shapes for various powerpoint operations.
The "loc pin" is the pin location within the shapes that should be fixed when using shape operations (e.g. changing the size).
'''
class LocPin(object):
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

def shape_indices_on_slide(slide, indices):
    import System.Array # to create int-Arrays
    return slide.Shapes.Range(System.Array[int](indices))

def last_n_shapes_on_slide(slide,n):
    return shape_indices_on_slide(slide, range(slide.shapes.Count + 1 -n, slide.shapes.Count + 1))

def get_shapes_from_selection(selection):
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
    # SlideRange accessible if slides, shapes or text selected
    try:
        return list(iter(selection.SlideRange))
    except:
        return []

def set_shape_zorder(shape, value=None, delta=None):
    if not delta and not value:
        raise ArgumentError("Neither value nor delta are given!")

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

# =========================================
# = Custom BKT tags stored in JSON format =
# =========================================

class BKTTag(object):
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
                tag_data = json.dumps(self.data)
                self.tags.Add(self.TAG_NAME,tag_data)
            else:
                self.tags.Delete(self.TAG_NAME)
        except:
            #import traceback #debugging only
            #traceback.print_exc()
            raise AttributeError
    
    def remove(self):
        self.data = {}

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
            raise ValueError
        
        self.data[arg] = value
    
    def __delitem__(self, arg):
        ''' access ribbon-attributes in dict-style, e.g. del button['label'] '''
        del self.data[arg]



# ======================
# = Slide content size =
# ======================

BKT_CONTENTAREA = "BKT_CONTENTAREA"

class ContentAreaTags(BKTTag):
    TAG_NAME = BKT_CONTENTAREA

    @property
    def is_area_set(self):
        return "contentarea_left" in self.data

    def get_area(self):
        return self.data["contentarea_left"], self.data["contentarea_top"], self.data["contentarea_width"], self.data["contentarea_height"]

def slide_content_size(presentation):
    shapes_sizes = [[shape.left, shape.top, shape.width, shape.height] for shape in iter(presentation.SlideMaster.Shapes) if shape.type == 14 and shape.Placeholderformat.type == 2]
    if len(shapes_sizes) == 0:
        return 0, 0, presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight
    else:
        slide_content_size = shapes_sizes[0]
        return slide_content_size[0], slide_content_size[1], slide_content_size[2], slide_content_size[3]

def isset_contentarea(presentation):
    if presentation.Tags.Item(BKT_CONTENTAREA) != '':
        with ContentAreaTags(presentation.Tags) as tags:
            return tags.is_area_set
    else:
        return False

def define_contentarea(presentation, shape):
    with ContentAreaTags(presentation.Tags) as tags:
        tags["contentarea_left"]   = float(shape.left)
        tags["contentarea_top"]    = float(shape.top)
        tags["contentarea_width"]  = float(shape.width)
        tags["contentarea_height"] = float(shape.height)
    #shape.Delete()

def reset_contentarea(presentation):
    presentation.tags.Delete(BKT_CONTENTAREA)

def read_contentarea(presentation):
    with ContentAreaTags(presentation.Tags) as tags:
        if tags.is_area_set:
            return tags.get_area() #left,top,width,height
        else:
            return slide_content_size(presentation) #left,top,width,height


# =========================================
# = Iterator for "subshapes" & textframes =
# =========================================

#Iterate through shapes of different types and return every shapes "subhsapes", e.g. group shapes or table cells
#arg 'from_selection': If shapes are not from a selection (e.g. iterate all shapes of a slide), set this to False to disable selected table cells detection,
#                      otherwise not all table cells are iterated at least in the rare case that a table is the only shape on a slide.
#arg 'filter_method':  Filter the returned shapes by a function(shape), e.g. to return only shapes that have a textframe
#arg 'getter_method':  Return function(shape) to get certain attributes, e.g. the textframe of a shape
def iterate_shape_subshapes(shapes, from_selection=True, filter_method=lambda shp: True, getter_method=lambda shp: shp):
    only_selected_table_cells = False

    def _get_shp_type(shape):
        #For table cells Type is not implemented and will throw an error
        try:
            return shape.Type
        except:
            return None

    def _iter_all(shape):
        for shape in shapes:
            shp_type = _get_shp_type(shape)
            
            # Note: Placeholder can be table, chart, diagram, smartart, picture, whatever...
            if shp_type == MsoShapeType['msoPlaceholder']:
                shp_type = shape.PlaceholderFormat.ContainedType

            # Iterate each group item
            if shp_type == MsoShapeType['msoGroup'] or shp_type == MsoShapeType['msoSmartArt']:
                for shp in shape.GroupItems:
                    yield shp
            
            # Iterate each chart/diagram shape
            elif shp_type == MsoShapeType['msoChart'] or shp_type == MsoShapeType['msoDiagram']:
                yield shape
                #FIXME: handling of charts can be improved, but it is very tricky!
                #General chart textframe is in shape.Chart.ChartArea.Format.TextFrame2, but there is not "HasTextFrame" property
                #Individual textframes are almost impossible to access
            
            # Iterate each table cell
            elif shp_type == MsoShapeType['msoTable']:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if not only_selected_table_cells or cell.Selected:
                            yield cell.Shape
            
            else:
                yield shape

    #Ensure list
    if type(shapes) != list:
        shapes = list(iter(shapes))

    #If cells within a table are selected, function should only iterate selected cells. If the whole table is selected but no other shape, all cells are selected.
    # if from_selection and len(shapes) == 1 and _get_shp_type(shapes[0]) == MsoShapeType['msoTable']:
    if from_selection and len(shapes) == 1 and shapes[0].HasTable == -1:
        only_selected_table_cells = True


    for shape in _iter_all(shapes):
        if filter_method(shape):
            yield getter_method(shape)


#Iterate through shapes of different types and return every shapes textframe
def iterate_shape_textframes(shapes, from_selection=True):
    return iterate_shape_subshapes(shapes, from_selection,
        filter_method=lambda shp: shp.HasTextFrame == -1,
        getter_method=lambda shp: shp.TextFrame2)



# ===============================
# = Generic class for rectangle =
# ===============================

class BoundingFrame(object):
    def __init__(self, slide=None, contentarea=False):
        self.left=0
        self.top=0
        self.width=100
        self.height=100
        self.rotation=0
        
        if slide != None:
            if contentarea:
                self.left, self.top, self.width, self.height = slide_content_size(slide.parent)
            else:
                self.width  = slide.parent.PageSetup.SlideWidth
                self.height = slide.parent.PageSetup.SlideHeight
    
    @classmethod
    def from_rect(cls, left, top, width, height):
        bf = BoundingFrame()
        bf.left = left
        bf.top = top
        bf.width = width
        bf.height = height
        return bf

    @classmethod
    def from_shapes(cls, shapes):
        from wrapper import wrap_shapes
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