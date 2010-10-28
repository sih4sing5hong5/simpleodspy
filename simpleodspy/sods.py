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
from odf.table import Table, TableRow, TableCell
from odf.style import Style, TextProperties, TableCellProperties, Map
from odf.text import P

from sodsspreadsheet import SodsSpreadSheet

class Sods(SodsSpreadSheet):
	def __init__(self):
		''' init and set default values for spreadsheet elements '''
		
		SodsSpreadSheet.__init__(self)
	
	def saveOds(self, filename, i_range, j_range):
		''' save table in ods format '''
		
		odfdoc = OpenDocumentSpreadsheet()
		table = Table()
		
		# regular style
		rs = Style(name="content", family="table-cell")
		rs.addElement(TextProperties(fontfamily="sans-serif", fontsize="12pt"))
		odfdoc.styles.addElement(rs)

		# make sure values are up to date
		# loop and update the cells value
		for i in range(1, i_range):
			# create new ods row
			tr = TableRow()
			table.addElement(tr)

			for j in range(1, j_range):
				# update the cell text and condition
				cell = self.encodeColName(j) + str(i)
				self.updateOneCell(cell)
				c = self.getCellAt(i, j)
				
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
					ccs_name = "c_" + cell
					ccs = Style(name = ccs_name, family = 'table-cell')
					ccs.addElement(TextProperties(color = c.condition_color))
					ccs.addElement(TableCellProperties(backgroundcolor = c.condition_background_color))
					odfdoc.automaticstyles.addElement(ccs)
					
					cs.addElement(Map(condition = c.condition, applystylename = ccs_name))
					
				# add style to document
				odfdoc.automaticstyles.addElement(cs)
				
				# create new ods cell
				if (c.formula):
					tc = TableCell(valuetype = c.value_type, 
						formula = c.formula, datevalue = c.date_value, stylename = cell)
				else:
					tc = TableCell(valuetype = c.value_type, 
						value = c.value, datevalue = c.date_value, stylename = cell)
				
				# set ods text
				if (c.value_type == 'string'):
					tc.addElement(P(text = c.text))
				
				tr.addElement(tc)

		odfdoc.spreadsheet.addElement(table)
		odfdoc.save(filename)
		
if __name__ == "__main__":
	
	t = Sods()
	
	print "Test spreadsheet naming:"
	print "-----------------------"
	
	t.setCell("A1", text = "Simple ods python")
	t.setCell("A1:G2", background_color = "#00ff00")
	t.setCell("A3:G5", background_color = "#ffff00")
	
	t.setValue("A2", 123.4)
	t.setValue("B2", "2010-01-01")
	t.setValue("C2", "=0.6")
	t.setValue("D2", "= A2 + 3")
	
	t.setCell("A3:D3", border_top = "1pt solid #ff0000")
	t.setValue("C3", "Sum of cells:")
	t.setValue("D3", "=sum(A2:D2)")
	
	t.setCell("D2:D3", condition = "value()<=200")
	t.setCell("D2:D3", condition_background_color = "#ff0000")
	
	t.saveHtml("test.html", 16,16)
	t.saveOds("test.ods", 16,16)
	
