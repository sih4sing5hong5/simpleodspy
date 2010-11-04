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
</head><body>
%s
</body></html>'''
	
		
	def exportCellHtml(self, c, i = 0, j = 0):
		''' export cell data as html table cell '''
		
		# if condition state is true use condition style
		# we assume condition_state is up to date
		color = [c.color, c.condition_color][c.condition_state]
		background_color = [c.background_color, c.condition_background_color][c.condition_state]
		
		# check for default backround color
		if background_color == "default":
			background_color = "#ffffff"
		
		# adjust values for html
		font_size = c.font_size.replace('pt', 'px')
		border_top = c.border_top.replace('pt', 'px')
		border_bottom = c.border_bottom.replace('pt', 'px')
		border_left = c.border_left.replace('pt', 'px')
		border_right = c.border_right.replace('pt', 'px')
		
		# get cell text
		if c.value_type == 'float':
			text = "%0.2f" % c.value
		else:
			text = escape(c.text) + "&nbsp;"
		
		# create cell string
		# we assume text is up to date
		out = '''
<td style="color:{0}; font-family:'{1}'; font-size:{2}; 
		background-color:{3}; 
		border-top:{4}; border-bottom:{5}; 
		border-left:{6}; border-right:{7}; ">
	{8}
</td>'''.format(color, c.font_family, font_size,
				background_color, 
				border_top, border_bottom, border_left, border_right,
				text)
		
		return out
	
	def exportHtml(self, i_max = None, j_max = None, delimiter = ",", txt_delimiter = '"'):
		''' export table in html format '''
		
		if not i_max: i_max = self.table.i_max
		if not j_max: j_max = self.table.j_max
		
		# create the table element of the html page
		out = "<table>\n"
		
		for i in range(1, i_max):
			out += "<tr>\n"
			for j in range(1, j_max):
				out += self.exportCellHtml(self.table.getCellAt(i,j), i, j)
			out += "</tr>\n"
		out += "</table>"
		
		return self.html_format % out
		
	def save(self, filename, i_max = None, j_max = None, delimiter = ",", txt_delimiter = '"'):
		''' save table in xml format '''
		
		# update cells text
		self.table.updateTable(i_max, j_max)
		
		# if filename is - print to stdout
		if filename == '-':
			print self.exportHtml(i_max, j_max)
		else:
			file(filename,"w").write(self.exportHtml(i_max, j_max))
		
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
	
	tw = SodsHtml(t)
	tw.save("test.html")
	
