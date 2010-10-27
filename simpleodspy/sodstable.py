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

from copy import deepcopy
from xml.etree import ElementTree

from sodscell import SodsCell

class SodsTable:
	def __init__(self):
		''' init and set default values for table elements '''
		
		# the data table
		self.rows = {}
		
		# dtd for the xml import/export
		# TODO: add dtd or scheme
		self.dtd= ''' '''
		
		# takes table
		self.html_format = '''<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head><body>
%s
</body></html>'''
		
		# takes dtd, cells
		self.xml_format = '''<?xml version="1.0" encoding="UTF-8"?>
%s
<table>
%s</table>'''
		
	def getCellAt(self, i, j):
		''' get the cell object in i,j '''
		
		# try to get the row i
		row = self.rows.get(i, {})
		
		# try to get the cell j
		return row.get(j, SodsCell())
	
	def setCellAt(self, i, j, new_cell):
		''' set the cell object in i, j '''
		
		# try to get the row i, if row is empty create new one
		row = self.rows.get(i, {})
		
		# if the value is Null, delete cell
		if new_cell:
			row[j] = new_cell
			self.rows[i] = row
		else:
			del row[j]
			# if we deleted the last cell in the row delete it
			if len(row) < 1:
				del self.rows[i]
		
		return
	
	def setAt(self, i_range, j_range, 
			font_size = None, font_family = None, color = None, 
			background_color = None, border_top = None,
			border_bottom = None, border_left = None, border_right = None,
			text = None, value_type = None,
			value = None, formula = None,
			date_value = None,
			condition = None, condition_state = None,
			condition_color = None, condition_background_color = None):
		''' set values of cell object range in i, j ranges '''
		
		# we want ranges
		if type(i_range) != type(list()):
			i_range = [i_range]
		
		if type(j_range) != type(list()):
			j_range = [j_range]
		
		# loop on cell range
		for i in i_range:
			for j in j_range:
				self.setAtOneCell(i, j, 
					font_size, font_family, color, 
					background_color, border_top,
					border_bottom, border_left, border_right,
					text, value_type, value, formula, date_value,
					condition, condition_state,
					condition_color, condition_background_color)
				
		return
		
	def setAtOneCell(self, i, j, 
			font_size = None, font_family = None, color = None, 
			background_color = None, border_top = None,
			border_bottom = None, border_left = None, border_right = None,
			text = None, value_type = None,
			value = None, formula = None,
			date_value = None,
			condition = None, condition_state = None,
			condition_color = None, condition_background_color = None):
		''' set values of one cell object in i, j '''
		
		# get the cell in i,j
		c = self.getCellAt(i, j)
				
		if font_size:
			c.font_size = font_size
		if font_family:
			c.font_family = font_family
		if color:
			c.color = color
	
		if background_color:
			c.background_color = background_color
		if border_top:
			c.border_top = border_top
		if border_bottom:
			c.border_bottom = border_bottom
		if border_left:
			c.border_left = border_left
		if border_right:
			c.border_right = border_right
	
		if text:
			c.text = text
		if value_type:
			c.value_type = value_type
		if value:
			c.value = value
		if formula:
			c.formula = formula
		if date_value:
			c.date_value = date_value
	
		if condition:
			c.condition = condition
		if condition_state:
			c.condition_state = condition_state
		if condition_color:
			c.condition_color = condition_color
		if condition_background_color:
			c.condition_background_color = condition_background_color
		
		# return cell to table
		self.setCellAt(i, j, c)
		
		return
	
	def exportHtml(self, i_max, j_max):
		''' export table in html format '''
		
		# create the table element of the html page
		out = "<table>\n"
		
		for i in range(1, i_max):
			out += "<tr>\n"
			for j in range(1, j_max):
				out += self.getCellAt(i,j).exportHtml()
			out += "</tr>\n"
		out += "</table>"
		
		return self.html_format % out
		
	def exportXml(self, i_max, j_max):
		''' export table in xml format '''
		
		out = ""
		
		for i in range(1, i_max):
			for j in range(1, j_max):
				out += self.getCellAt(i,j).exportXml(i,j)
		
		return self.xml_format % (self.dtd, out)
		
	def loadXml(self, xml_text):
		''' load cells from text in xml format '''
		
		# get the cells elements from our xml file
		xml_table = ElementTree.XML(xml_text)
		
		# loop on all the cells in xml file
		for xml_cell in xml_table:
			# FIXME: we assume that all the cell element are in the right
			# order
			
			# get i, j
			i, j = int(xml_cell[0].text), int(xml_cell[1].text)
			
			# get cell
			c = SodsCell()
			
			c.color = xml_cell[2].text
			c.font_family = xml_cell[3].text
			c.font_size = xml_cell[4].text
			
			c.background_color = xml_cell[5].text
			c.border_top = xml_cell[6].text
			c.border_bottom = xml_cell[7].text
			c.border_left = xml_cell[8].text
			c.border_right = xml_cell[9].text
	
			if xml_cell[10].text:
				c.text = xml_cell[10].text
			c.value_type = xml_cell[11].text
			c.value = [None, xml_cell[12].text][xml_cell[12].text == 'None']
			c.formula = [None, xml_cell[13].text][xml_cell[13].text == 'None']
			c.date_value = [None, xml_cell[14].text][xml_cell[14].text == 'None']
	
			c.condition = [None, xml_cell[15].text][xml_cell[15].text == 'None']
			c.condition_state = eval(xml_cell[16].text)
			c.condition_color = xml_cell[17].text
			c.condition_background_color = xml_cell[18].text
			
			# insert cell to table
			self.setCellAt(i, j, c)
			
	def copy(self):
		''' return a copy of the table '''
		
		return deepcopy(self)

if __name__ == "__main__":
	
	t = SodsTable()
	
	t.setAt(1,1, text = "hello world")
	t.setAt(1,range(1,3), background_color = "#00ff00")
	
	print "Test table export:"
	print "------------------------------"
	
	file("test.xml","w").write(t.exportXml(6,6))
	file("test.html","w").write(t.exportHtml(6,6))
	
	print "Test table xml load from file:"
	print "------------------------------"
	
	t2 = SodsTable()
	t2.loadXml(file("test.xml").read())
	file("test2.xml","w").write(t2.exportXml(6,6))
	
