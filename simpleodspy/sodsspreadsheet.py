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
import math
from datetime import datetime

from sodstable import SodsTable

class SodsSpreadSheet(SodsTable):
	def __init__(self, i_max = 30, j_max = 30):
		''' init and set default values for spreadsheet elements '''
		
		SodsTable.__init__(self, i_max, j_max)
		
	def parseColName(self, name):
		''' parse a col name "A" or "BC" to the col number 1.. '''
		
		if name == '':
			return 0
		
		return (ord(name[-1:]) - ord('A') + 1) + 26 * self.parseColName(name[:-1])
	
	def encodeColName(self, n):
		''' parse a col number 1.. to a name "A" or "BC" '''
		
		n1 = (n - 1) % 26
		n2 = int((n - 1) / 26)
		
		if n2 == 0:
			return chr(ord('A') + n1)
		
		return chr(ord('A') + n2 - 1) + chr(ord('A') + n1)
		
	def parseOneCellName(self, name):
		''' parse a cell name "A2" or "BC43" to the row and col numbers (2,3) '''
		
		# A..ZZ part is columns numbers (j) and the number part is rows (i)
		p = re.search('([A-Z]+)([0-9]+)', name)
		if p:
			return (int(p.group(2)), self.parseColName(p.group(1)))
			
		return None
	
	def parseCellName(self, name):
		''' parse a cell name or a range name
		
		parse "B3", "A2:A4" or "A1:BC43" to the row and col ranges ([2,3,4],[3]) 
		'''
		names = name.split(':')
		
		# if this is one cell return the cell
		if len(names) == 1:
			return self.parseOneCellName(names[0])
		
		# if this is a range, get the two cells
		if len(names) == 2:
			a1 = self.parseOneCellName(names[0])
			a2 = self.parseOneCellName(names[1])
			
			if a1 and a2:
				return (range(a1[0], a2[0] + 1), range(a1[1], a2[1] + 1))
			
		return None
	
	def parseCellRangeToCells(self, name):
		''' parse a range name "A2:A4" or "A1:BC43" to array of cells '''
		
		cell_list = []
		
		# get the range part of the formula
		i_range, j_range = self.parseCellName(name)
		
		# we want ranges
		if type(i_range) != type(list()):
			i_range = [i_range]
		
		if type(j_range) != type(list()):
			j_range = [j_range]
		
		# loop and create the cells sum
		for i in i_range:
			for j in j_range:
				cell = self.encodeColName(j) + str(i)
				cell_list.append(cell)
		
		return cell_list
	
	def isFloat(self, x):
		''' return true if x represent float '''
		
		try:
			y = float(x)
			return True
		except:
			return False
		
		return False
	
	def isDate(self, x):
		''' return true if x represent date '''
		
		try:
			y = datetime.strptime(x, "%Y-%m-%d")
			return True
		except:
			return False
		
		return False
		
	def setValue(self, name, value):
		''' set cell/s value and type automatically '''
		
		# parse i,j from cell name
		i_range, j_range = self.parseCellName(name)
		
		# we want ranges
		if type(i_range) != type(list()):
			i_range = [i_range]
		
		if type(j_range) != type(list()):
			j_range = [j_range]
		
		# loop on cell range
		for i in i_range:
			for j in j_range:
				# get the cell
				c = self.getCellAt(i, j)
		
				# delete old value
				c.text = ""
				c.value_type = "string"
				c.value = None
				c.formula = None
				c.date_value = None
			
				# check cell type
				if type(value) == type(str()) and len(value) > 0 and value[0] == '=':
					c.value_type = "float"
					c.formula = value
				elif self.isFloat(value):
					c.value_type = "float"
					c.value = float(value)
					c.text = str(value)
				elif self.isDate(value):
					c.value_type = "date"
					c.date_value = str(value)
					c.text = str(value)
				else:
					c.value_type = "string"
					c.text = str(value)
			
				# set the cell
				self.setCellAt(i, j, c)
	
	def setCell(self, name, 
			font_size = None, font_family = None, color = None, 
			background_color = None, border_top = None,
			border_bottom = None, border_left = None, border_right = None,
			text = None, value_type = None,
			value = None, formula = None,
			date_value = None,
			condition = None, condition_state = None,
			condition_color = None, condition_background_color = None):
		''' set values of cell object range in i, j ranges '''
		
		# parse i,j from cell name
		i, j = self.parseCellName(name)
		
		# set at can handle i,j ranges
		self.setAt(i, j, 
			font_size, font_family, color, 
			background_color, border_top,
			border_bottom, border_left, border_right,
			text, value_type,
			value, formula,
			date_value,
			condition, condition_state,
			condition_color, condition_background_color)
		
	def getRangeString(self, cells):
		''' create a string of cells from range '''
		
		# we work with a string, 
		# get the re.group string
		cell_range = cells.group(0)
			
		# get the cell list
		cell_list = self.parseCellRangeToCells(cell_range)
		
		out = "(" + ",".join(cell_list) + ")"
			
		# return the sum as string
		return out
	
	def getOneCellValueRe(self, name):
		''' return the updated float value of a cell input name is re group '''
		
		# we work with a string, 
		# get the re.group string
		name = name.group(0)
		
		return self.getOneCellValue(name)
	
	def getOneCellValue(self, name):
		''' return the updated float value of a cell '''
		
		# parse i,j from cell name
		i, j = self.parseCellName(name)
		
		c = self.getCellAt(i, j)
		
		# if this is not a float cell return 0
		if (c.value_type != 'float'):
			return "0.0"
			
		# if this is just a value, return it
		if (not c.formula):
			return str(c.value)
		
		# this is a formula, we need to update it's value
		self.updateOneCell(name)
		
		# re-get the cell
		c = self.getCellAt(i, j)
		
		return str(c.value)
		
	def updateOneCell(self, name):
		''' update one cell text '''
		
		# parse i,j from cell name
		i, j = self.parseCellName(name)
		
		c = self.getCellAt(i, j)
		
		# check if the cell has formula
		if c.formula:
			# remove the '='
			formula = c.formula[1:]
			
			# remove white spaces
			formula = re.sub(r'\s', '', formula)
			formula = formula.replace('!', '')
			formula = formula.replace(';', '')
			formula = formula.replace('$', '')
			formula = formula.replace('import', '')
			formula = formula.replace('print', '')
			
			# replce spreadsheet function names to python function
			formula = formula.replace('SUM(', 'sum(')
			formula = formula.replace('MIN(', 'min(')
			formula = formula.replace('MAX(', 'max(')
			formula = formula.replace('ABS(', 'abs(')
			
			formula = formula.replace('POWER(', 'math.pow(')
			formula = formula.replace('SQRT(', 'math.sqrt(')
			
			formula = formula.replace('PI', 'math.pi')
			formula = formula.replace('SIN(', 'math.sin(')
			formula = formula.replace('COS(', 'math.cos(')
			formula = formula.replace('TAN(', 'math.cos(')
			formula = formula.replace('ASIN(', 'math.asin(')
			formula = formula.replace('ACOS(', 'math.acos(')
			formula = formula.replace('ATAN(', 'math.atan(')
			
			# check for ranges e.g. 'A2:G3' and replace them with (A2,A3 ... G3) tupple
			formula = re.sub('[A-Z]+[0-9]+:[A-Z]+[0-9]+', 
				self.getRangeString, formula)
		
			# get all the cell names in this formula and replace them with values
			value = eval(re.sub('[A-Z]+[0-9]+', self.getOneCellValueRe, formula))
			
			# update cell value and text string
			c.value = value
			c.text = str(value)
			
		# check if the cell has condition
		if c.condition:
			# replace the cell-content() function with the cells value
			if (c.value):
				value = c.value
			else:
				value = self.getOneCellValue(name)
			formula = c.condition.replace("cell-content()", str(value))
		
			# check for ranges e.g. 'A2:G3' and replace them with (A2,A3 ... G3) tupple
			formula = re.sub('[A-Z]+[0-9]+:[A-Z]+[0-9]+', 
				self.getRangeString, formula)
		
			# get all the cell names in this formula and replace them with values
			value = eval(re.sub('[A-Z]+[0-9]+', self.getOneCellValueRe, formula))
		
			# update condition state
			c.condition_state = value
		
		self.setCellAt(i, j, c)
		
	def updateCell(self, name):
		''' update cell text value '''
		
		# parse i,j from cell name
		i_range, j_range = self.parseCellName(name)
		
		# we want ranges
		if type(i_range) != type(list()):
			i_range = [i_range]
		
		if type(j_range) != type(list()):
			j_range = [j_range]
		
		# loop and update the cells value
		for i in i_range:
			for j in j_range:
				cell = self.encodeColName(j) + str(i)
				self.updateOneCell(cell)
		
	def saveXml(self, filename, i_max = None, j_max = None):
		''' save table in xml format '''
		
		if not i_max: i_max = self.i_max
		if not j_max: j_max = self.j_max
		
		# make sure values are up to date
		# loop and update the cells value
		for i in range(1, i_max):
			for j in range(1, j_max):
				cell = self.encodeColName(j) + str(i)
				self.updateOneCell(cell)
		
		# if filename is - print to stdout
		if filename == '-':
			print self.exportXml(i_max, j_max)
		else:
			file(filename,"w").write(self.exportXml(i_max, j_max))
	
	def loadXmlfile(self, filename):
		''' load a table from file '''
		
		self.loadXml(file(filename).read())
		
	def saveHtml(self, filename, i_max = None, j_max = None):
		''' save table in xml format '''
		
		if not i_max: i_max = self.i_max
		if not j_max: j_max = self.j_max
		
		# make sure values are up to date
		# loop and update the cells value
		for i in range(1, i_max):
			for j in range(1, j_max):
				cell = self.encodeColName(j) + str(i)
				self.updateOneCell(cell)
		
		# if filename is - print to stdout
		if filename == '-':
			print self.exportHtml(i_max, j_max)
		else:
			file(filename,"w").write(self.exportHtml(i_max, j_max))
		
if __name__ == "__main__":
	
	t = SodsSpreadSheet()
	
	print "Test spreadsheet naming:"
	print "-----------------------"
	
	t.setCell("A1", text = "Hello world")
	t.setCell("A1:G2", background_color = "#00ff00")
	t.setCell("A3:G5", background_color = "#ffff00")
	
	t.setValue("A2", 123.4)
	t.setValue("B2", "2010-01-01")
	t.setValue("C2", "0.6")
	t.setValue("D2", "= SIN(PI/2)")
	
	t.setCell("A3:D3", border_top = "1pt solid #ff0000")
	t.setValue("C3", "Sum of cells:")
	t.setValue("D3", "=SUM($A$2:D2)")
	
	t.setCell("D2:D3", condition = "cell-content()<=200")
	t.setCell("D2:D3", condition_background_color = "#ff0000")
	
	t.updateCell("A1:G3")
	
	t.saveHtml("test.html") 
	
