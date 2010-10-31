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

from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.style import Style, TextProperties, TableCellProperties, TableColumnProperties, Map
from odf.number import NumberStyle, CurrencyStyle, TextStyle, Number,  Text
from odf.text import P

from sodscell import SodsCell

class SodsOds():
	def __init__(self, table, i_max = 30, j_max = 30):
		''' init and set default values for spreadsheet elements '''
		
		self.table = table
	
	def saveOds(self, filename, i_max = None, j_max = None):
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
		cs.addElement(TableColumnProperties(columnwidth = "2.4cm", breakbefore = "auto"))
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

			for j in range(1, j_max):
				# update the cell text and condition
				cell = self.table.encodeColName(j) + str(i)
				self.table.updateOneCell(cell)
				c = self.table.getCellAt(i, j)
				
				# set ods style
				cs = Style(name = cell, family = 'table-cell')
				cs.addElement(TextProperties(color = c.color, 
					fontsize =c.font_size, fontfamily = c.font_family))
				cs.addElement(TableCellProperties(backgroundcolor = c.background_color,
					bordertop = c.border_top,
					borderbottom = c.border_bottom,
					borderleft = c.border_left,
					borderright = c.border_right))
					
				# set ods conditional style
				if (c.condition):
					cns = Style(name = "cns"+cell, family = 'table-cell')
					cns.addElement(TextProperties(color = c.condition_color))
					cns.addElement(TableCellProperties(backgroundcolor = c.condition_background_color))
					odfdoc.styles.addElement(cns)
					
					cs.addElement(Map(condition = c.condition, applystylename = "cns"+cell))
				
				odfdoc.automaticstyles.addElement(cs)
				
				# create new ods cell
				if (c.formula):
					tc = TableCell(valuetype = c.value_type, 
						formula = c.formula, value = c.value, stylename = cell)
				elif (c.value_type == 'date'):
					tc = TableCell(valuetype = c.value_type, 
						datevalue = c.date_value, stylename = cell)
				elif (c.value_type == 'float'):
					tc = TableCell(valuetype = c.value_type, 
						value = c.value, stylename = cell)
				else:
					tc = TableCell(valuetype = c.value_type, stylename = cell)
				
				# set ods text
				if (tc and c.value_type == 'string'):
					tc.addElement(P(text = c.text))
				
				tr.addElement(tc)

		odfdoc.spreadsheet.addElement(table)
		odfdoc.save(filename)
		
if __name__ == "__main__":
	
	from sodsspreadsheet import SodsSpreadSheet

	t = SodsSpreadSheet()
	
	print "Test spreadsheet naming:"
	print "-----------------------"
	
	t.setStyle("A1", text = "Simple ods python")
	t.setStyle("A1:G2", background_color = "#00ff00")
	t.setStyle("A3:G5", background_color = "#ffff00")
	
	t.setValue("A2", 123.4)
	t.setValue("B2", "2010-01-01")
	t.setValue("C2", "=0.6")
	t.setValue("D2", "= A2 + 3")
	
	t.setStyle("A3:D3", border_top = "1pt solid #ff0000")
	t.setValue("C3", "Sum of cells:")
	t.setValue("D3", "=sum(A2:D2)")
	
	t.setStyle("D2:D3", condition = "cell-content()<=200")
	t.setStyle("D2:D3", condition_color = "#ff0000")
	
	t.saveHtml("test.html", 16,16)
	
	tw = SodsOds(t)
	tw.saveOds("test.ods")
	
