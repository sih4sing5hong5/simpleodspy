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

from xml.sax.saxutils import escape

class SodsCell:
	def __init__(self):
		''' init and set default values for cell elements '''
		
		# TextProperties
		self.font_size = "12pt"
		self.font_family = "Arial"
		self.color = "#000000"
		
		# TableCellProperties
		self.background_color = "default"
		self.border_top = "none"
		self.border_bottom = "none"
		self.border_left = "none"
		self.border_right = "none"
		
		# TableCell
		self.text = ""
		self.format = ""
		self.value_type = "string"
		self.value = None
		self.formula = None
		self.date_value = None
		
		# Map
		self.condition = None
		self.condition_state = False
		self.condition_color = "#000000"
		self.condition_background_color = "#ffffff"
	
if __name__ == "__main__":
	c = SodsCell()

	print "Test html export:"
	print "-----------------"
	
	c.text = "hello world"
	c.condition_state = True
	c.text = "hello world"
	c.value = 123.3
	c.condition_state = False

