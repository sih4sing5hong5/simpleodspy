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

import sys
from xml.sax.saxutils import unescape

from odf.opendocument import OpenDocumentSpreadsheet
from odf.opendocument import load
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.style import Style, TextProperties, TableCellProperties, TableColumnProperties, Map
from odf.number import NumberStyle, DateStyle, CurrencyStyle, TextStyle, Number, Text, Day, Month, Year, Era
from odf.text import P, Date

from sodscell import SodsCell

class SodsOds():
	def __init__(self, table, i_max = 30, j_max = 30):
		''' init and set default values for spreadsheet elements '''
		
		self.table = table
		self.styles = {}
	
	def getStyle(self, c, cell, datastylename, style_id, odfdoc):
		''' get a style_name by style_id '''
		
		if not style_id in self.styles.keys():
			cs = Style(name = cell, family = 'table-cell', datastylename=datastylename)
			cs.addElement(TextProperties(color = c.color, 
				fontsize =c.font_size, fontfamily = c.font_family))
			
			if c.background_color != "default":
				cs.addElement(TableCellProperties(backgroundcolor = c.background_color))
			if c.border_top != "none":
				cs.addElement(TableCellProperties(bordertop = c.border_top))
			if c.border_bottom != "none":
				cs.addElement(TableCellProperties(borderbottom = c.border_bottom))
			if c.border_left != "none":
				cs.addElement(TableCellProperties(borderleft = c.border_left))
			if c.border_right != "none":
				cs.addElement(TableCellProperties(borderright = c.border_right))
			
			# set ods conditional style
			if (c.condition):
				cns = Style(name = "cns"+cell, family = 'table-cell')
				cns.addElement(TextProperties(color = c.condition_color))
				cns.addElement(TableCellProperties(backgroundcolor = c.condition_background_color))
				odfdoc.styles.addElement(cns)
				
				cs.addElement(Map(condition = c.condition, applystylename = "cns"+cell))
				
			odfdoc.automaticstyles.addElement(cs)
			
			self.styles[style_id] = cell
		
		return self.styles[style_id]
		
	def load(self, filename):
		''' load a table in ods format '''
		
		# read the file
		doc = load(filename)
		
		# loop on all rows and columns
		i = 0
		for row in doc.getElementsByType(TableRow):
			j = 1
			i += 1
			for cell in row.getElementsByType(TableCell):
				
				# creat a new cell
				c = SodsCell()
				
				try:
					numbercolumnsrepeated = int(cell.getAttribute('numbercolumnsrepeated'))
				except:
					numbercolumnsrepeated = 1
				
				# get cell data
				try:
					c.text = unescape(str(cell.firstChild))
				except:
					c.text = unescape(str(cell.firstChild.firstChild))
					
				c.value_type = cell.getAttribute('valuetype')
				c.formula = cell.getAttribute('formula')
				if c.formula:
					c.formula = unescape(c.formula)
				c.date_value = cell.getAttribute('datevalue')
				c.value = cell.getAttribute('value')
				
				stylename = cell.getAttribute('stylename')
				if stylename:
					style = doc.getStyleByName(stylename)
					datastyle = style.getAttribute('datastylename')
				
					for p in style.getElementsByType(TextProperties):
						c.font_family = p.getAttribute('fontfamily')
						c.font_size = p.getAttribute('fontsize')
						c.color = p.getAttribute('color')
				
					for p in style.getElementsByType(TableCellProperties):
						if stylename == 'ce3':
							print "  -- ce3", doc.getStyleByName(stylename)
							print "  -- ce3 x ",  p.getAttribute('bordertop')
							
						c.background_color = p.getAttribute('backgroundcolor')
						c.border_top = p.getAttribute('bordertop')
						if not c.border_top:
							c.border_top = "none"
						c.border_bottom = p.getAttribute('borderbottom')
						if not c.border_bottom:
							c.border_bottom = "none"
						c.border_left = p.getAttribute('borderleft')
						if not c.border_left:
							c.border_left = "none"
						c.border_right = p.getAttribute('borderright')
						if not c.border_right:
							c.border_right = "none"
						print "  -- ce3 w ",  c.border_top
						
					for p in style.getElementsByType(Map):
						c.condition = p.getAttribute('condition')
						if c.condition:
							c.condition = unescape(c.condition)
					
						applystylename = p.getAttribute('applystylename')
						applystyle = doc.getStyleByName(applystylename)
					
						for cp in style.getElementsByType(TextProperties):
							c.condition_color = cp.getAttribute('color')
					
						for cp in style.getElementsByType(TableCellProperties):
							c.condition_background_color = cp.getAttribute('backgroundcolor')
				
				# check for sodsOds formulas (starting with !)
				if len(c.text) > 0 and c.text[0] == '!':
					c.formula = c.text
					c.value_type = "float"
				
				# insert cell to table
				while numbercolumnsrepeated > 0:
					self.table.setCellAt(i, j, c)
					j += 1
					numbercolumnsrepeated -= 1
				
	def save(self, filename, i_max = None, j_max = None):
		''' save table in ods format '''
		
		if not i_max: i_max = self.table.i_max
		if not j_max: j_max = self.table.j_max
		
		# create new odf spreadsheet
		odfdoc = OpenDocumentSpreadsheet()
		table = Table()
		
		# default style
		ts = Style(name = "ts", family = "table-cell")
		ts.addElement(TextProperties(fontfamily = SodsCell().font_family, fontsize = SodsCell().font_size))
		odfdoc.styles.addElement(ts)
		
		cs = Style(name = "cs", family = "table-column")
		cs.addElement(TableColumnProperties(columnwidth = "2.8cm", breakbefore = "auto"))
		odfdoc.automaticstyles.addElement(cs)

		# create columns
		for j in range(1, j_max):
			table.addElement(TableColumn(stylename = "cs", defaultcellstylename = "ts"))
			
		# make sure values are up to date
		# loop and update the cells value
		for i in range(1, i_max):
			# create new ods row
			tr = TableRow()
			table.addElement(tr)
			
			# create default data styles for dates and numbers
			ncs = NumberStyle(name="ncs")
			ncs.addElement(Number(decimalplaces="2", minintegerdigits="1", grouping="true"))
			odfdoc.styles.addElement(ncs)
			
			dcs = DateStyle(name="dcs")
			dcs.addElement(Year(style='long'))
			dcs.addElement(Text(text = u'-'))
			dcs.addElement(Month(style='long'))
			dcs.addElement(Text(text = u'-'))
			dcs.addElement(Day(style='long'))
			odfdoc.styles.addElement(dcs)
			
			for j in range(1, j_max):
				# update the cell text and condition
				cell = self.table.encodeColName(j) + str(i)
				self.table.updateOneCell(cell)
				c = self.table.getCellAt(i, j)
				
				# chose datastylename
				if c.value_type == 'date':
					datastylename = "dcs"
				else:
					datastylename = "ncs"
				
				# get cell style id
				if (c.condition):
					style_id = (datastylename + c.color + c.font_size + c.font_family + 
						c.background_color + c.border_top + c.border_bottom + 
						c.border_left + c.border_right + 
						c.condition_color + c.condition_background_color)
				else:
					style_id = (datastylename + c.color + c.font_size + c.font_family + 
						c.background_color + c.border_top + c.border_bottom + 
						c.border_left + c.border_right)
				
				# set ods style
				style_name = self.getStyle(c, cell, datastylename, style_id, odfdoc)
				
				# create new ods cell
				if (c.formula):
					tc = TableCell(valuetype = c.value_type, 
						formula = c.formula, value = c.value, stylename = style_name)
				elif (c.value_type == 'date'):
					tc = TableCell(valuetype = c.value_type, 
						datevalue = c.date_value, stylename = style_name)
				elif (c.value_type == 'float'):
					tc = TableCell(valuetype = c.value_type, 
						value = c.value, stylename = style_name)
				else:
					tc = TableCell(valuetype = c.value_type, stylename = style_name)
				
				# set ods text
				tc.addElement(P(text = unicode(c.text, 'utf-8')))
				
				tr.addElement(tc)

		odfdoc.spreadsheet.addElement(table)
		odfdoc.save(filename)
		
if __name__ == "__main__":
	
	from sodsspreadsheet import SodsSpreadSheet
	
	t = SodsSpreadSheet(12, 12)
	
	print "Test spreadsheet naming:"
	print "-----------------------"
	
	t.setStyle("A1", text = "Hello world")
	t.setStyle("A1:G2", background_color = "#00ff00")
	t.setStyle("A3:G5", background_color = "#ffff00")
	
	t.setValue("A2", 123.4)
	t.setValue("B2", "2010-01-01")
	t.setValue("C2", "0.6")
	
	t.setValue("C5", 0.6)
	t.setValue("C6", 0.6)
	t.setValue("C7", 0.8)
	t.setValue("C8", 0.8)
	t.setValue("C9", "=AVERAGE(C5:C8)")
	t.setValue("C10", "=SUM(C5:C8)")
	
	t.setValue("D2", "= SIN(PI()/2)")
	t.setValue("D10", "=IF(A2>3;C7;C9)")
	
	t.setStyle("A3:D3", border_top = "1pt solid #ff0000")
	t.setValue("C3", "Sum of cells:")
	t.setValue("D3", "=SUM($A$2:D2)")
	
	t.setStyle("D2:D3", condition = "cell-content()<=100")
	t.setStyle("D2:D3", condition_background_color = "#ff0000")
	
	tw = SodsOds(t)
	#tw.save("test.ods")
	
	print "Test load:"
	print "----------"
	
	t2 = SodsSpreadSheet(12, 12)
	tw = SodsOds(t2)
	tw.load("test.ods")
	t2.saveXml('text.xml')
