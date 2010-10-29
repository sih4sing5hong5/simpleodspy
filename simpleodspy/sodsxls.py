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

from xlwt import *

from sods import Sods

class SodsXls(Sods):
	def __init__(self, i_max = 30, j_max = 30):
		''' init and set default values for spreadsheet elements '''
		
		Sods.__init__(self, i_max, j_max)
	
	def convertXlsBorderWidth(self, border):
		return [Borders.NO_LINE, Borders.THIN][border != 'none']
	
	def convertXlsBorderColor(self, border):
		return 0
		
	def convertXlsColor(self, color):
		return 1
		
	def saveXls(self, filename, i_max = None, j_max = None):
		''' save table in ods format '''
		
		if not i_max: i_max = self.i_max
		if not j_max: j_max = self.j_max
		
		# create new xls spreadsheet
		w = Workbook(encoding='utf-8')
		ws = w.add_sheet("sheet 1")
		
		# make sure values are up to date
		# loop and update the cells value
		for i in range(1, i_max):
			for j in range(1, j_max):
				# update the cell text and condition
				cell = self.encodeColName(j) + str(i)
				self.updateOneCell(cell)
				c = self.getCellAt(i, j)
				
				# set xls style
				fnt = Font()
				fnt.name = c.font_family
				fnt.height = 18 * int(c.font_size.replace('pt',''))
				fnt.colour_index = 0x0000 #self.convertXlsColor(c.color)
				
				borders = Borders()
				borders.left = self.convertXlsBorderWidth(c.border_left)
				borders.left_colour = self.convertXlsBorderColor(c.border_left)
				borders.right = self.convertXlsBorderWidth(c.border_right)
				borders.right_colour = self.convertXlsBorderColor(c.border_right)
				borders.top = self.convertXlsBorderWidth(c.border_top)
				borders.top_colour = self.convertXlsBorderColor(c.border_top)
				borders.bottom = self.convertXlsBorderWidth(c.border_bottom)
				borders.bottom_colour = self.convertXlsBorderColor(c.border_bottom)
				
				pattern = Pattern()
				pattern.pattern = Pattern.SOLID_PATTERN
				pattern.pattern_fore_colour = 0x0001 #self.convertXlsColor(c.background_color)
				
				style = XFStyle()
				style.font = fnt
				style.borders = borders
				style.pattern = pattern
				
				# set xls text
				if (c.formula):
					ws.write(i - 1, j - 1, Formula(c.formula[1:]), style)
				else:
					ws.write(i - 1, j - 1, c.text, style)
		
		w.save(filename)
		
if __name__ == "__main__":
	
	t = SodsXls()
	
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
	
	t.setCell("D2:D3", condition = "cell-content()<=200")
	t.setCell("D2:D3", condition_color = "#ff0000")
	
	t.saveHtml("test.html", 16,16)
	t.saveXls("test.xls")
	
