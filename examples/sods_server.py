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
import urllib
import getopt
from urlparse import unquote
from BaseHTTPServer import BaseHTTPRequestHandler, HTTPServer

from simpleodspy.sodsspreadsheet import SodsSpreadSheet
from simpleodspy.sodshtml import SodsHtml

t = SodsSpreadSheet(16,8)
	
class HttpHandler(BaseHTTPRequestHandler):
	
	def log_message(self, format, *args):
		pass
	
	def parse_Path(self):
		full_path = self.path.strip('/').split('?')
		self.clean_path = full_path[0]
		
		self.parameters = {}
		
		if len(full_path) == 2:
			for arg in full_path[1].split('&'):
				arg_t = arg.split('=')
				if len(arg_t) == 2:
					self.parameters[arg_t[0]] = unquote(arg_t[1])
				else:
					self.parameters[arg] = ""
			
	def do_GET(self):
		# get path
		self.parse_Path()
		
		# page headers
		self.send_response(200)
		self.send_header('Content-type', 'text/html')
		self.end_headers()
		
		# do table operations
		if 'cell' in self.parameters.keys() and 'value' in self.parameters.keys():
			cell = self.parameters['cell']
			value = self.parameters['value']
			if cell != "":
				t.setValue(self.parameters['cell'], self.parameters['value'])
		
		if 'cell' in self.parameters.keys() and 'bgcolor' in self.parameters.keys():
			cell = self.parameters['cell']
			value = self.parameters['value']
			if cell != "":
				t.setStyle(self.parameters['cell'], background_color = self.parameters['bgcolor'])
		
		# set form
		form_html = '''
		<form>
		<table>
		<tr><td>Name</td><td><input type='text' name = 'cell' /></td></tr>
		<tr><td>Value</td><td><input type='text' name = 'value' /></td></tr>
		<tr><td>Bckground color</td><td><input type='text' name = 'bgcolor' /></td></tr>
		<tr><td></td><td><input type='submit' /></td></tr>
		</table>
		<form>
		'''
		
		# print out page data
		tw = SodsHtml(t)
		print >> self.wfile, "<html><head>"
		print >> self.wfile, tw.exportTableCss()
		print >> self.wfile, "</head><body>"
		print >> self.wfile, form_html
		print >> self.wfile, tw.exportTableHtml(headers = True)
		print >> self.wfile, "</body></html>"
		
		# finish page
		self.wfile.flush()

# run web server and rates reader
def main(argv=None):
	if argv is None:
		argv = sys.argv
		
	try:
		opts, args = getopt.getopt(argv[1:], "hp:", ["help","port"])
	except getopt.error, msg:
		print >>sys.stderr, msg
		return 2
	
	# default port
	port = 8080
	server = None
	
	# get user options
	for o, a in opts:
		if o in ("-h", "--help"):
			usage()
			sys.exit()
		elif o in ("-p", "--port"):
			port = int(a)
		else:
			usage()
			sys.exit()
	
	# run the reader and a basic web server
	try:
		# check if a server already running
		try:
			f = urllib.urlopen("http://127.0.0.1:%s/" % port, proxies={})
			
			print """
			A web server already using port %s
			Quiting.
			""" % (port)
			return
		except:
			pass
		
		print """
		Web access using port %s
		Press CTRL-C to stop the server.
		""" % (port)
		
		# create server and reader objects
		server = HTTPServer(('', port), HttpHandler)
			
		# start the web server loop
		server.serve_forever()
	
	except KeyboardInterrupt:
		print """
		Quiting.
		"""
	
	finally:
		if server:	
			server.socket.close()
			
if __name__ == '__main__':
	sys.exit(main())
