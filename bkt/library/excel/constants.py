# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''



XlPasteType = {
    'xlPasteAll':                       -4104,
    'xlPasteColumnWidths':                  8,
    'xlPasteFormats':                   -4122,
    'xlPasteFormulas':                  -4123,
    'xlPasteFormulasAndNumberFormats':     11,
    'xlPasteValidation':                    6,
    'xlPasteValues':                    -4163,
    'xlPasteValuesAndNumberFormats':       12,
}

XlSheetVisibility = {
    'xlSheetHidden':       0,   #Hides the worksheet which the user can unhide via menu.
    'xlSheetVeryHidden':   2,   #Hides the object so that the only way for you to make it visible again is by setting this property to True (the user cannot make the object visible).
    'xlSheetVisible':     -1,   #Displays the sheet.
}

XlSheetType = {
    'xlChart':                 -4109,   #Chart
    'xlDialogSheet':           -4116,   #Dialog sheet
    'xlExcel4IntlMacroSheet':      4,   #Excel version 4 international macro sheet
    'xlExcel4MacroSheet':          3,   #Excel version 4 macro sheet
    'xlWorksheet':             -4167,   #Worksheet
}

XlCellType = {
    'xlCellTypeAllFormatConditions':    -4172,   #Cells of any format.
    'xlCellTypeAllValidation':          -4174,   #Cells having validation criteria.
    'xlCellTypeBlanks':                     4,   #Empty cells.
    'xlCellTypeComments':               -4144,   #Cells containing notes.
    'xlCellTypeConstants':                  2,   #Cells #containing constants.
    'xlCellTypeFormulas':               -4123,   #Cells containing formulas.
    'xlCellTypeLastCell':                  11,   #The last cell in the used range.
    'xlCellTypeSameFormatConditions':   -4173,   #Cells having the same format.
    'xlCellTypeSameValidation':         -4175,   #Cells having the same validation criteria.
    'xlCellTypeVisible':                   12,   #All visible cells.
}

XlApplicationInternational = {
    'xlGeneralFormatName':  26,
    'xlDecimalSeparator':    3,
    'xlListSeparator':       5,
}

XlCalculation = {
    'xlCalculationAutomatic':   -4105,
    'xlCalculationManual':      -4135,
}

XlRangeValueDataType = {
    'xlRangeValueDefault':  10,
}

XlFormatConditionType = {
    "xlAboveAverageCondition":      12,      #Above average condition
    "xlBlanksCondition":            10,      #Blanks condition
    "xlCellValue":                  1,       #Cell value
    "xlColorScale":                 3,       #Color scale
    "xlDatabar":                    4,       #Databar
    "xlErrorsCondition":            16,      #Errors condition
    "xlExpression":                 2,       #Expression
    "XlIconSet":                    6,       #Icon set
    "xlNoBlanksCondition":          13,      #No blanks condition
    "xlNoErrorsCondition":          17,      #No errors condition
    "xlTextString":                 9,       #Text string
    "xlTimePeriod":                 11,      #Time period
    "xlTop10":                      5,       #Top 10 values
    "xlUniqueValues":               8,       #Unique values
}

XlFormatConditionOperator = {
    "xlBetween":        1,      #Between. Can be used only if two formulas are provided.
    "xlEqual":          3,      #Equal.
    "xlGreater":        5,      #Greater than.
    "xlGreaterEqual":   7,      #Greater than or equal to.
    "xlLess":           6,      #Less than.
    "xlLessEqual":      8,      #Less than or equal to.
    "xlNotBetween":     2,      #Not between. Can be used only if two formulas are provided.
    "xlNotEqual":       4,      #Not equal.
}

XlDVType = {
    "xlValidateCustom":         7,  #Data is validated using an arbitrary formula.
    "xlValidateDate":           4,  #Date values.
    "xlValidateDecimal":        2,  #Numeric values.
    "xlValidateInputOnly":      0,  #Validate only when user changes the value.
    "xlValidateList":           3,  #Value must be present in a specified list.
    "xlValidateTextLength":     6,  #Length of text.
    "xlValidateTime":           5,  #Time values.
    "xlValidateWholeNumber":    1,  #Whole numeric values.
}

XlDVAlertStyle = {
    "xlValidAlertInformation":  3,  #Information icon.
    "xlValidAlertStop":         1,  #Stop icon.
    "xlValidAlertWarning":      2,  #Warning icon.
}

subtotalFunction = {
    "AVERAGE":  101,
    "MAX":  104,
    "MIN":  105,
    "SUM":  109,
}

errValues = {
    -2146826281: "#DIV/0!",    #"#Div0!",
    -2146826246: "#NV",        #"#N/A",
    -2146826259: "#NAME?",     #"#Name",
    -2146826288: "#NULL!",     #"#Null!",
    -2146826252: "#ZAHL!",     #"#Num!",
    -2146826265: "#BEZUG!",    #"#Ref!",
    -2146826273: "#WERT!",     #"#Value!",
}