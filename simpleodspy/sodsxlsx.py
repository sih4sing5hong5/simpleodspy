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

import re
from datetime import datetime

from openpyxl.workbook import Workbook
from openpyxl.style import Color, Fill
from openpyxl.writer.excel import ExcelWriter

class SodsXlsx():
	def __init__(self, table, i_max = 30, j_max = 30):
		''' init and set default values for spreadsheet elements '''
		
		self.table = table
	
	def saveXlsx(self, filename, i_max = None, j_max = None):
		''' save table in ods format '''
		
		if not i_max: i_max = self.table.i_max
		if not j_max: j_max = self.table.j_max
		
		# create new xlsx spreadsheet
		wb = Workbook()
		ew = ExcelWriter(workbook = wb)
		dest_filename = filename
		ws = wb.worksheets[0]
		ws.title = filename

		# make sure values are up to date
		# loop and update the cells value
		for i in range(1, i_max):
			for j in range(1, j_max):
				# update the cell text and condition
				cell = self.table.encodeColName(j) + str(i)
				self.table.updateOneCell(cell)
				c = self.table.getCellAt(i, j)
				
				# set xls text
				if (c.formula):
					ws.cell(cell).value = c.formula
				elif c.value_type == 'float':
					ws.cell(cell).value = c.value
				elif c.value_type == 'date':
					ws.cell(cell).value = datetime.strptime(c.date_value, "%Y-%m-%d")
				else:
					ws.cell(cell).value = c.text
		
		ew.save(filename = filename)
		
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
	
	tw = SodsXlsx(t)
	tw.saveXlsx("test.xlsx")
	
