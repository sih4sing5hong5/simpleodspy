#!/usr/bin/env python
# -*- coding: utf-8 -*-

from distutils.core import setup

setup(name='simpleodspy',
	version='1.4.1',
	description='Simple spreadsheet class with ods export',
	author='Yaacov Zamir',
	author_email='kzamir@walla.co.il',
	url='http://simple-odspy.sourceforge.net/',
	packages=['simpleodspy'],
	
	classifiers=[
		'Programming Language :: Python :: 3.3'
	],
	install_requires=['xlwt-future>=0.8.0'],
	)
     
