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

class SodsCell:
	def __init__(self):
		''' init and set default values for cell elements '''
		
		# TextProperties
		self.font_size = "12pt"
		self.font_family = "sans-serif"
		self.color = "#000000"
		
		# TableCellProperties
		self.background_color = "#ffffff"
		self.border_top = "none"
		self.border_bottom = "none"
		self.border_left = "none"
		self.border_right = "none"
		
		# TableCell
		self.text = ""
		self.value_type = "string"
		self.value = None
		self.formula = None
		self.date_value = None
		
		# Map
		self.condition = None
		self.condition_state = False
		self.condition_color = "#000000"
		self.condition_background_color = "#ffffff"
		
	def exportHtml(self, i = 0, j = 0):
		''' export cell data as html table cell '''
		
		# if condition state is true use condition style
		# we assume condition_state is up to date
		color = [self.color, self.condition_color][self.condition_state]
		background_color = [self.background_color, self.condition_background_color][self.condition_state]
		
		# adjust values for html
		font_size = self.font_size.replace('pt', 'px')
		border_top = self.border_top.replace('pt', 'px')
		border_bottom = self.border_bottom.replace('pt', 'px')
		border_left = self.border_left.replace('pt', 'px')
		border_right = self.border_right.replace('pt', 'px')
		
		# create cell string
		# we assume text is up to date
		out = '''
<td style="color:{0}; font-family:'{1}'; font-size:{2}; 
		background-color:{3}; 
		border-top:{4}; border-bottom:{5}; 
		border-left:{6}; border-right:{7}; ">
	{8}
</td>'''.format(color, self.font_family, font_size,
				background_color, 
				border_top, border_bottom, border_left, border_right,
				self.text + "&nbsp;")
		
		return out
	
	def exportXml(self, i = 0, j = 0):
		''' export cell data as xml table cell '''
		
		# create cell string
		out = '''
<cell>
	<i>{0}</i>
	<j>{1}</j>
	
	<color>{2}</color>
	<font_family>{3}</font_family>
	<font_size>{4}</font_size>
	
	<background_color>{5}</background_color>
	<border_top>{6}</border_top>
	<border_bottom>{7}</border_bottom>
	<border_left>{8}</border_left>
	<border_right>{9}</border_right>
	
	<text>{10}</text>
	<value_type>{11}</value_type>
	<value>{12}</value>
	<formula>{13}</formula>
	<date_value>{14}</date_value>
	
	<condition>{15}</condition>
	<condition_state>{16}</condition_state>
	<condition_color>{17}</condition_color>
	<condition_background_color>{18}</condition_background_color>
</cell>'''.format(i, j,
				self.color, self.font_family, self.font_size,
				self.background_color, 
				self.border_top, self.border_bottom, 
				self.border_left, self.border_right,
				self.text, self.value_type, self.value, 
				self.formula, self.date_value, 
				self.condition, self.condition_state, 
				self.condition_color, self.condition_background_color)
		
		return out

if __name__ == "__main__":
	c = SodsCell()

	print "Test html export:"
	print "-----------------"
	
	print c.exportHtml()
	c.text = "hello world"
	print c.exportHtml()
	c.condition_state = True
	print c.exportHtml()

	print "Test xml export:"
	print "-----------------"
	
	print c.exportXml()
	c.text = "hello world"
	c.value = 123.3
	print c.exportXml()
	c.condition_state = False
	print c.exportXml()

