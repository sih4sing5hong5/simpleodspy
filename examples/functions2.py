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
from simpleodspy.sodshtml import SodsHtml
 
t = SodsSpreadSheet()
 
def mul3Callback(arg):
	val = t.evaluateFormula(arg) * 3 
	return str(val)
 
t.registerFunction('MUL3', mul3Callback)
 
t.setValue("D2", 123.5)
t.setValue("D3", "=3.0/5.0")
t.setValue("D4", "=MUL3(D2+D3)")
 
tw = SodsHtml(t)
tw.save("test.html")

