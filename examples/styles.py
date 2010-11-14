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
 
t.setValue("A1:A3", "green")
t.setStyle("A1:A3", background_color = "#00ff00")
t.setValue("B1:B3", "blue")
t.setStyle("B1:B3", background_color = "#0000ff")
t.setStyle("B3", border_top = "1pt solid #ff0000", border_bottom = "1pt solid #ff0000")
 
tw = SodsHtml(t)
tw.save("test.html")

