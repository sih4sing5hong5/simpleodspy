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

from sodsspreadsheet import SodsSpreadSheet

class Sods(SodsSpreadSheet):
	def __init__(self):
		''' init and set default values for spreadsheet elements '''
		
		SodsSpreadSheet.__init__(self)
	
	def saveXml(self, filename, i_range, j_range):
		''' save table in xml format '''
		
		# make sure values are up to date
		# loop and update the cells value
		for i in range(1, i_range):
			for j in range(1, j_range):
				cell = self.encodeColName(j) + str(i)
				self.updateOneCell(cell)
		
		# if filename is - print to stdout
		if filename == '-':
			print t.exportXml(i_range, j_range)
		else:
			file(filename,"w").write(t.exportXml(i_range, j_range))
	
	def loadXml(self, filename):
		''' load a table from file '''
		
		self.loadXml(file(filename).read())
		
	def saveHtml(self, filename, i_range, j_range):
		''' save table in xml format '''
		
		# make sure values are up to date
		# loop and update the cells value
		for i in range(1, i_range):
			for j in range(1, j_range):
				cell = self.encodeColName(j) + str(i)
				self.updateOneCell(cell)
		
		# if filename is - print to stdout
		if filename == '-':
			print t.exportHtml(i_range, j_range)
		else:
			file(filename,"w").write(t.exportHtml(i_range, j_range))
	
	def saveOds(self, filename, i_range, j_range):
		''' save table in ods format '''
		
		# make sure values are up to date
		# loop and update the cells value
		for i in range(1, i_range):
			for j in range(1, j_range):
				cell = self.encodeColName(j) + str(i)
				self.updateOneCell(cell)
		
		
if __name__ == "__main__":
	
	t = Sods()
	
	print "Test spreadsheet naming:"
	print "-----------------------"
	
	t.setCell("A1", text = "Simple ods python")
	t.setCell("A1:G2", background_color = "#00ff00")
	t.setCell("A3:G5", background_color = "#ffff00")
	
	t.setValue("A2", 123.4)
	t.setValue("B2", "2010-01-01")
	t.setValue("C2", "0.6")
	t.setValue("D2", "= A2 + 3")
	
	t.setCell("A3:D3", border_top = "1pt solid #ff0000")
	t.setValue("C3", "Sum of cells:")
	t.setValue("D3", "=sum(A2:D2)")
	
	t.setCell("D2:D3", condition = "value()<=200")
	
	t.saveHtml("test.html", 16,16)
	
