# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''

from __future__ import absolute_import

from collections import namedtuple # required for color class

from bkt import config, message
from bkt.library.excel import constants
from bkt.library import algorithms # required for color helper

# from System.Runtime.InteropServices.Marshal import ReleaseComObject

# def release_object(object):
#     ReleaseComObject(object)

application = None

def set_application(app):
    global application
    application = app


def confirm_no_undo(text="Dies kann nicht rückgängig gemacht werden. Ausführen?"):
	if config.excel_ignore_warnings:
		return True
	else:
		return message.confirmation(text)


restore_screen_updating = True
restore_display_alerts = True
restore_interactive = True
restore_calculation = constants.XlCalculation["xlCalculationAutomatic"]

def freeze_app(disable_screen_updating=True, disable_display_alerts=False, disable_interactive=False, disable_calculation=False):
    global restore_screen_updating, restore_display_alerts, restore_interactive, restore_calculation
    restore_screen_updating = application.ScreenUpdating
    restore_display_alerts = application.DisplayAlerts
    restore_interactive = application.Interactive
    restore_calculation = application.Calculation
    
    if disable_screen_updating:
        application.ScreenUpdating = False
    if disable_display_alerts:
        application.DisplayAlerts = False
    if disable_interactive:
        application.Interactive = False
    if disable_calculation:
        application.Calculation = constants.XlCalculation["xlCalculationManual"]

def unfreeze_app(force=False):
    application.ScreenUpdating = force or restore_screen_updating
    application.DisplayAlerts = force or restore_display_alerts
    application.Interactive = force or restore_interactive
    application.Calculation = constants.XlCalculation["xlCalculationAutomatic"] if force else restore_calculation
    application.CutCopyMode = False


def get_unused_ranges(sheet):
    left_border = sheet.UsedRange.Column-1
    right_border = left_border + sheet.UsedRange.Columns.Count +1
    top_border = sheet.UsedRange.Row-1
    bottom_border = top_border + sheet.UsedRange.Rows.Count +1

    selection = []

    if left_border > 0:
        selection.append( sheet.Range(sheet.Columns(1), sheet.Columns(left_border)) )
    if top_border > 0:
        selection.append( sheet.Range(sheet.Rows(1), sheet.Rows(top_border)) )

    if right_border < sheet.Columns.Count:
        selection.append( sheet.Range(sheet.Columns(right_border), sheet.Columns(sheet.Columns.Count)) )
    if bottom_border < sheet.Rows.Count:
        selection.append( sheet.Range(sheet.Rows(bottom_border), sheet.Rows(sheet.Rows.Count)) )

    return selection

def create_temp_sheet():
    temporary_sheet = application.ActiveWorkbook.Sheets.Add()
    temporary_sheet.Visible = 0 #xlSheetHidden
    return temporary_sheet

def formula_int2local(formula):
	save_alert_state = application.DisplayAlerts
	application.DisplayAlerts = False
	temporary_sheet = create_temp_sheet()
	temporary_sheet.Range("A1").Formula = formula
	return_formula = temporary_sheet.Range("A1").FormulaLocal
	temporary_sheet.Delete()
	application.DisplayAlerts = save_alert_state
	return return_formula

def rename_sheet(sheet, name):
    try:
        sheet.Name = name[:31]
    except:
        i=2
        while True:
            try:
                li = 31 - len(str(i)) - 1
                sheet.Name = name[:li] + " " + str(i)
                return
            except:
                i += 1

def get_address_external(rng, row_abs=False, col_abs=False):
    import re
    address = rng.AddressLocal(row_abs, col_abs, External=True)
    return re.sub("\[.*\]", "", address)


def resize_areas(areas, rows=None, cols=None, rows_delta=None, cols_delta=None):
    #iterate over copy if areas list to be able to remove area from list
    return_areas = []
    for area in areas:
        if rows_delta is not None:
            rows = area.Rows.Count + rows_delta
            if rows < 0:
                continue
        
        if cols_delta is not None:
            cols = area.Columns.Count + cols_delta
            if cols < 0:
                continue
        
        if rows is not None and cols is not None:
            return_areas.append( area.Resize(rows, cols) )
        elif rows is not None:
            return_areas.append( area.Resize(RowSize=rows) )
        elif cols is not None:
            return_areas.append( area.Resize(ColumnSize=cols) )

    return return_areas

def resize_range(selection, rows=None, cols=None, rows_delta=None, cols_delta=None):
    new_range = None
    areas = resize_areas( list(iter(selection.Areas)), rows, cols, rows_delta, cols_delta )
    #for area in areas:
    #    new_range = range_union(new_range, area)
    if len(areas) > 0:
        sep = application.International(5) #xlListSeparator FIXME: Does this always return the correct separator for range???
        new_range = application.Range(sep.join([area.AddressLocal(False, False) for area in areas]))
    return new_range
 

def xls_evaluate(input_text, dec_sep=None, numberformat=None):
    dec_sep = dec_sep or application.International(constants.XlApplicationInternational["xlDecimalSeparator"])
    
    def _conv_dec(dec):
        #use unicode instead str as str has problems with unicode characters
        if dec is None:
            return ''
        if numberformat is not None:
            try:
                return application.WorksheetFunction.Text(dec, numberformat)
            except:
                pass
        return unicode(dec).replace('.', dec_sep)

    res = application.Evaluate(input_text)
    #Test if result is iterable
    try:
        res = list(iter(res))
    except:
        pass
    type_res = type(res)

    #Error values are int
    if type_res == int and res in constants.errValues:
        return "FEHLER: " + constants.errValues[res]
    
    #Boolean statements
    elif type_res == bool:
        return "WAHR" if res else "FALSCH"
    
    #Result of calculation, default case
    elif type_res == float:
        return _conv_dec(res)
    
    #Any iterable object, can be range of cells or System.Array
    elif type_res == list:
        try:
            ret = [_conv_dec(cell.Value2) for cell in res]
        except:
            ret = [_conv_dec(value) for value in res]
        if len(ret) > 1:
            return "{" + ";".join(ret) + "}"
        else:
            return ret[0]
    
    #Fallback
    else:
        return unicode(res)
        #return "{0!r}".format(res).replace('.', dec_sep)


directions = {
    'bottom': (1,0),
    'top': (-1,0),
    'left': (0,-1),
    'right': (0,1)
}

def get_next_cell(cell, direction='bottom'):
    if not direction in directions:
        raise KeyError('Invalid direction: ' + str(direction))
    
    if direction == 'top' and cell.Row == 1:
            raise IndexError('Reached top border of sheet')
    
    if direction == 'bottom' and cell.Row == cell.Worksheet.Rows.Count:
            raise IndexError('Reached bottom border of sheet')

    if direction == 'left' and cell.Column == 1:
            raise IndexError('Reached left border of sheet')
    
    if direction == 'right' and cell.Column == cell.Worksheet.Columns.Count:
            raise IndexError('Reached right border of sheet')
    
    return cell.Offset(*directions[direction])

def get_next_visible_cell(cell, direction='bottom'):
    if not direction in directions:
        raise KeyError('Invalid direction: ' + str(direction))
    
    cur_cell = cell
    while True:
        cur_cell = get_next_cell(cur_cell, direction)
        if direction in ['bottom', 'top'] and not cur_cell.EntireRow.Hidden:
            break
        if direction in ['left', 'right'] and not cur_cell.EntireColumn.Hidden:
            break

        # if not cur_cell.EntireRow.Hidden and not cur_cell.EntireColumn.Hidden:
        #     break
    return cur_cell


def range_union(range1, range2):
    if not range1:
        return range2
    if not range2:
        return range1

    return application.Union(range1, range2)


# Original function: http://dailydoseofexcel.com/archives/2007/08/17/two-new-range-functions-union-and-subtract/
def range_substract(main_range, diff_range):
    #application = main_range.Application

    def _range_subtract_one_area(main_area, diff_area):
        if main_area.Areas.Count > 1 or diff_area.Areas.Count > 1:
            raise ValueError("Range consists of more than one area")

        int_area = application.Intersect(main_area, diff_area)
        if not int_area:
            # print "only main area: " + main_area.Address()
            return main_area

        sheet = main_area.Parent
        int_selected = None

        if int_area.Row > main_area.Row:
            int_selected = sheet.Range(main_area.Rows(1), main_area.Rows(int_area.Row - main_area.Row))

        if int_area.Row + int_area.Rows.Count < main_area.Row + main_area.Rows.Count:
            int_selected = range_union(int_selected, sheet.Range(main_area.Rows(int_area.Row - main_area.Row + int_area.Rows.Count + 1), main_area.Rows(main_area.Rows.Count)) )

        if int_area.Column > main_area.Column:
            int_selected = range_union(int_selected, sheet.Range(sheet.Cells(int_area.Row, main_area.Column), sheet.Cells(int_area.Row + int_area.Rows.Count - 1, int_area.Column - 1)) )

        if int_area.Column + int_area.Columns.Count < main_area.Column + main_area.Columns.Count:
            int_selected = range_union(int_selected, sheet.Range(sheet.Cells(int_area.Row, int_area.Column + int_area.Columns.Count), sheet.Cells(int_area.Row + int_area.Rows.Count - 1, main_area.Column + main_area.Columns.Count - 1)) )

        # print "return subtract one area: " + int_selected.Address()
        return int_selected

    #Begin of main function
    if not application.Intersect(main_range, diff_range):
        # print "only main range: " + main_range.Address()
        return main_range

    cells_selected = None

    for m_area in iter(main_range.Areas):
        # print "main area interation: " + m_area.Address()
        diff_areas = iter(diff_range.Areas)
        area_selected = _range_subtract_one_area(m_area, next(diff_areas)) #First area
        for d_area in diff_areas:
            # print "diff area iteration: " + d_area.Address()
            area_selected = application.Intersect(area_selected, _range_subtract_one_area(m_area, d_area))

        cells_selected = range_union(cells_selected, area_selected)

    #release_object(application)

    # print "return subtract: " + cells_selected.Address()
    return cells_selected


class ColorHelper(object):
    '''
    For description refer to PowerPoint ColorHelper class
    '''
    _theme_color_indices = [1,2,3,4, 5,6,7,8,9,10] #powerpoint default color picker is using IDs 5-10 and 13-16
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
            1: 2,
            2: 1,
            3: 4,
            4: 3,
        }
        return mapping[index]
    
    @classmethod
    def _get_color_from_theme_index(cls, context, index): #expect MsoThemeColorSchemeIndex
        if index < 5:
            index = cls._theme_color_index_2_color_scheme_index(index)
        return context.app.ActiveWorkbook.Theme.ThemeColorScheme(index)

    @classmethod
    def _get_factors_for_rgb(cls, color_rgb):
        r,g,b = algorithms.get_rgb_from_ole(color_rgb)
        l = round( algorithms.get_brightness_from_rgb(r,g,b) / 255. *100 )
        return [factors[1] for factors in cls._theme_color_shades if l in factors[0]][0]
    
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
        #excel does not provide recent colors in VBA
        return 0

    @classmethod
    def get_recent_colors_count(cls, context):
        return 0