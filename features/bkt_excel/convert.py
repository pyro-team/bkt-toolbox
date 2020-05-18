# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''

convert_units = {
	"Weight and mass | Gram":											"g",
	"Weight and mass | Slug":											"sg",
	"Weight and mass | Pound mass (avoirdupois)":						"lbm",
	"Weight and mass | U (atomic mass unit)":							"u",
	"Weight and mass | Ounce mass (avoirdupois)":						"ozm",
	"Weight and mass | Grain":											"grain",
	"Weight and mass | U.S. (short) hundredweight":						"cwt",
	"Weight and mass | Imperial hundredweight":							"uk_cwt",
	"Weight and mass | Stone":											"stone",
	"Weight and mass | Ton":											"ton",
	"Weight and mass | Imperial ton":									"uk_ton",
	"Distance | Meter":													"m",
	"Distance | Statute mile":											"mi",
	"Distance | Nautical mile":											"Nmi",
	"Distance | Inch":													"in",
	"Distance | Foot":													"ft",
	"Distance | Yard":													"yd",
	"Distance | Angstrom":												"ang",
	"Distance | Ell":													"ell",
	"Distance | Light-year":											"ly",
	"Distance | Parsec":												"parsec",
	"Distance | Pica (1/72 inch)":										"Picapt",
	"Distance | Pica (1/6 inch)":										"pica",
	"Distance | U.S survey mile (statute mile)":						"survey_mi",
	"Time | Year":														"yr",
	"Time | Day":														"day",
	"Time | Hour":														"hr",
	"Time | Minute":													"mn",
	"Time | Second":													"sec",
	"Pressure | Pascal":												"Pa",
	"Pressure | Atmosphere":											"atm",
	"Pressure | mm of Mercury":											"mmHg",
	"Pressure | PSI":													"psi",
	"Pressure | Torr":													"Torr",
	"Force | Newton":													"N",
	"Force | Dyne":														"dyn",
	"Force | Pound force":												"lbf",
	"Force | Pond":														"pond",
	"Energy | Joule":													"J",
	"Energy | Erg":														"e",
	"Energy | Thermodynamic calorie":									"c",
	"Energy | IT calorie":												"cal",
	"Energy | Electron volt":											"eV",
	"Energy | Horsepower-hour":											"HPh",
	"Energy | Watt-hour":												"Wh",
	"Energy | Foot-pound":												"flb",
	"Energy | BTU":														"BTU",
	"Power | Horsepower":												"HP",
	"Power | Pferdestärke":												"PS",
	"Power | Watt":														"W",
	"Magnetism | Tesla":												"T",
	"Magnetism | Gauss":												"ga",
	"Temperature | Degree Celsius":										"C",
	"Temperature | Degree Fahrenheit":									"F",
	"Temperature | Kelvin":												"K",
	"Temperature | Degrees Rankine":									"Rank",
	"Temperature | Degrees Réaumur":									"Reau",
	"Volume (or liquid measure ) | Teaspoon":							"tsp",
	"Volume (or liquid measure ) | Modern teaspoon":					"tspm",
	"Volume (or liquid measure ) | Tablespoon":							"tbs",
	"Volume (or liquid measure ) | Fluid ounce":						"oz",
	"Volume (or liquid measure ) | Cup":								"cup",
	"Volume (or liquid measure ) | U.S. pint":							"pt",
	"Volume (or liquid measure ) | U.K. pint":							"uk_pt",
	"Volume (or liquid measure ) | Quart":								"qt",
	"Volume (or liquid measure ) | Imperial quart (U.K.)":				"uk_qt",
	"Volume (or liquid measure ) | Gallon":								"gal",
	"Volume (or liquid measure ) | Imperial gallon (U.K.)":				"uk_gal",
	"Volume (or liquid measure ) | Liter":								"l",
	"Volume (or liquid measure ) | Cubic angstrom":						"ang3",
	"Volume (or liquid measure ) | U.S. oil barrel":					"barrel",
	"Volume (or liquid measure ) | U.S. bushel":						"bushel",
	"Volume (or liquid measure ) | Cubic feet":							"ft3",
	"Volume (or liquid measure ) | Cubic inch":							"in3",
	"Volume (or liquid measure ) | Cubic light-year":					"ly3",
	"Volume (or liquid measure ) | Cubic meter":						"m3",
	"Volume (or liquid measure ) | Cubic Mile":							"mi3",
	"Volume (or liquid measure ) | Cubic yard":							"yd3",
	"Volume (or liquid measure ) | Cubic nautical mile":				"Nmi3",
	"Volume (or liquid measure ) | Cubic Pica":							"Picapt3",
	"Volume (or liquid measure ) | Gross Registered Ton":				"GRT",
	"Volume (or liquid measure ) | Measurement ton (freight ton)":		"MTON",
	"Area | International acre":										"uk_acre",
	"Area | U.S. survey/statute acre":									"us_acre",
	"Area | Square angstrom":											"ang2",
	"Area | Are":														"ar",
	"Area | Square feet":												"ft2",
	"Area | Hectare":													"ha",
	"Area | Square inches":												"in2",
	"Area | Square light-year":											"ly2",
	"Area | Square meters":												"m2",
	"Area | Morgen":													"Morgen",
	"Area | Square miles":												"mi2",
	"Area | Square nautical miles":										"Nmi2",
	"Area | Square Pica":												"Picapt2",
	"Area | Square yards":												"yd2",
	"Information | Bit":												"bit",
	"Information | Byte":												"byte",
	"Speed | Admiralty knot":											"admkn",
	"Speed | Knot":														"kn",
	"Speed | Meters per hour":											"m/h",
	"Speed | Meters per second":										"m/s",
	"Speed | Miles per hour":											"mph",
}

convert_prefixes = {
	"yotta":		"Y",			#1,00E+24
	"zetta":		"Z",			#1,00E+21
	"exa":			"E",			#1,00E+18
	"peta":			"P",			#1,00E+15
	"tera":			"T",			#1,00E+12
	"giga":			"G",			#1,00E+09
	"mega":			"M",			#1,00E+06
	"kilo":			"k",			#1,00E+03
	"hecto":		"h",			#1,00E+02
	"dekao":		"da", 		#1,00E+01
	"deci":			"d",			#1,00E-01
	"centi":		"c",			#1,00E-02
	"milli":		"m",			#1,00E-03
	"micro":		"u",			#1,00E-06
	"nano":			"n",			#1,00E-09
	"pico":			"p",			#1,00E-12
	"femto":		"f",			#1,00E-15
	"atto":			"a",			#1,00E-18
	"zepto":		"z",			#1,00E-21
	"yocto":		"y",			#1,00E-24
}