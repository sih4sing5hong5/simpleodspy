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

class SodsHtml():
	def __init__(self, table, i_max = 30, j_max = 30):
		''' init and set default values for spreadsheet elements '''
		
		self.table = table
		
		# takes table
		self.html_format = '''<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
%s
</head><body>
%s
</body></html>'''
	
	def fancyNumber(self, f, sep = ",", neg = "-", pos = ""):
		''' format a fancy string for a number '''
		
		# split number
		n_parts = str(f).split(".")
		
		if len(n_parts) == 2:
			n, m = n_parts
		else:
			n, m = n_parts[0], "00"
		
		sign = False
		if n[0] == "-":
			n = n[1:]
			sign = True
		
		# call the positive int formater
		return [pos, neg][sign] + self.fancyIntNumber(n, sep) + "." + m[:2]
		
	def fancyIntNumber(self, n, sep = ","):
		''' format a fancy string for int number '''
		
		# we have string, continue
		if len(n) < 4: return n
		return self.fancyIntNumber(n[:-3]) + sep + n[-3:]
	
	def exportCellCss (self, c, i = 0, j = 0):
		''' export cell data as html table cell '''
		
		# if condition state is true use condition style
		# we assume condition_state is up to date
		color = [c.color, c.condition_color][c.condition_state]
		background_color = [c.background_color, c.condition_background_color][c.condition_state]
		
		# check for default backround color
		if background_color == "default":
			background_color = "#ffffff"
		
		# adjust values for html
		cell_name = self.table.encodeCellName(i, j)
		font_size = c.font_size.replace('pt', 'px')
		border_top = c.border_top.replace('pt', 'px')
		border_bottom = c.border_bottom.replace('pt', 'px')
		border_left = c.border_left.replace('pt', 'px')
		border_right = c.border_right.replace('pt', 'px')
		text_align = c.text_align.replace('start', 'right').replace('end', 'left')
		
		if self.table.direction == 'rtl':
			border = border_left
			border_left = border_right
			border_right = border
			
		# get cell text
		if c.value_type == 'float':
			text = self.fancyNumber(c.value) 
		else:
			text = escape(c.text)
			
		# create cell string
		# we assume text is up to date
		out = '''#%s {
	padding: 2px 10px; 
	color:%s; 
	font-family:'%s'; 
	font-size:%s; 
	background-color:%s; 
	border-top:%s; 
	border-bottom:%s; 
	border-left:%s; 
	border-right:%s; 
	text-align:%s;
}

''' % (cell_name, color, c.font_family, font_size,
				background_color, 
				border_top, border_bottom, border_left, border_right,
				text_align)
		
		return out
	
	def exportTableCss (self, i_max = None, j_max = None):
		''' export table in html format '''
		
		if not i_max: i_max = self.table.i_max
		if not j_max: j_max = self.table.j_max
		
		# update cells text
		self.table.updateTable(i_max, j_max)
		
		# create the table element of the html page
		out = "<style type='text/css'>\n"

		#self.table.direction
		# table
		out += '''
body {direction: %s}
table { border-collapse:collapse; }
''' % self.table.direction
		
		# cols
		for j in range(1, j_max):
			# adjust values for html
			col_name = self.table.encodeColName(j)
			c = self.table.getCellAt(0, j)
				
			# printout
			out += ".%s {width:%s;}\n" % (col_name, c.column_width.replace('pt','px'))
					
		# rows
		for i in range(1, i_max):
			for j in range(1, j_max):
				# adjust values for html
				cell_name = self.table.encodeCellName(i, j)
				c = self.table.getCellAt(i, j)
					
				# printout
				out += self.exportCellCss(c, i, j)
		out += "</style>"
		
		return out
	
	def exportTableHtml(self, i_max = None, j_max = None):
		''' export table in html format '''
		
		if not i_max: i_max = self.table.i_max
		if not j_max: j_max = self.table.j_max
		
		# update cells text
		self.table.updateTable(i_max, j_max)
		
		# create the table element of the html page
		out = "<table>\n"
		
		# columns
		for j in range(1, j_max):
			width = self.table.getCellAt(0, j).column_width.replace('pt', 'px')
			out += "<col class='%s' />\n" % self.table.encodeColName(j)
		
		# rows
		for i in range(1, i_max):
			out += "<tr>\n"
			for j in range(1, j_max):
				# adjust values for html
				cell_name = self.table.encodeCellName(i, j)
				c = self.table.getCell(cell_name)
				
				# get cell text
				if c.value_type == 'float':
					text = self.fancyNumber(c.value) 
				else:
					text = escape(c.text)
					
				# printout
				out += "<td id = '%s'>%s &nbsp;</td>\n" % (cell_name, text)
			out += "</tr>\n"
		out += "</table>"
		
		return out
	
	def exportHtml(self, i_max = None, j_max = None):
		''' export table in html format '''
			
		return self.html_format % (self.exportTableCss(i_max, j_max),
			self.exportTableHtml(i_max, j_max))
	
	def save(self, filename, i_max = None, j_max = None):
		''' save table in xml format '''
		
		# if filename is - print to stdout
		if filename == '-':
			print self.exportHtml(i_max, j_max)
		else:
			file(filename,"w").write(self.exportHtml(i_max, j_max))
		
if __name__ == "__main__":
	
	from sodsspreadsheet import SodsSpreadSheet
	
	t = SodsSpreadSheet(20,6)
	
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
	
	tw = SodsHtml(t)
	tw.save("test.html")
	
