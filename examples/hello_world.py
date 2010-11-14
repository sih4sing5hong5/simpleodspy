#!/usr/bin/env python
# -*- coding: utf-8 -*-

# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.

# Copyright (C) 2010 Yaacov Zamir (2010) <kzamir@walla.co.il>
# Author: Yaacov Zamir (2010) <kzamir@walla.co.il>

from simpleodspy.sodsspreadsheet import SodsSpreadSheet

t = SodsSpreadSheet(10, 10)
	
print "Test spreadsheet naming:"
print "-----------------------"

# setting values and formulas
t.setValue("A1", "Hello world !")

t.setValue("B2:C2", 123.4)
t.setValue("B3", "=B2+3")
t.setValue("B4", "=sum(B2:B3)")

# cell styles
t.setStyle("A1", color = "#0000ff")
t.setStyle("A1:A5", background_color = "#00ff00")
t.setStyle("B1:B5", border_left = "1pt solid #ff0000", background_color = "#ffff00")

# conditional styles
t.setStyle("B1:B5", condition = "cell-content()<=125")
t.setStyle("B1:B5", condition_color = "#ff0000")

# export
from simpleodspy.sodsxml import SodsXml
tw = SodsXml(t)
tw.save("test.xml")

# import
t2 = SodsSpreadSheet(10, 10)
tw = SodsXml(t2)
tw.load("test.xml")
tw.save("test2.xml")

from simpleodspy.sodshtml import SodsHtml

def myCallback(args):
	val = 0
	for arg in t.rangeIterator(args):
		val += t.evaluateFormula(arg) * 2
		
	return str(val)

t.registerFunction('MY', myCallback)
t.setValue("B5", "=MY(B2:B3)")

tw = SodsHtml(t)
tw.save("test.html")


