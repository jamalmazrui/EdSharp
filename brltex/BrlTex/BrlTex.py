#!/usr/bin/python
#BrlTex
#Copyright (C) 2006 Michael Whapples, All Rights Reserved.
#Original Author: Michael Whapples, mwhapples@users.sourceforge.net
#Unless linsor has explicitly granted you any other licensing terms, the contents of this file is subject to the RPL 
1.1, 
#available at the BrlTex website (http://brltex.sourceforge.net).
#All software distributed under the Licenses is provided strictly on an "AS IS" basis,
#WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, AND the licensor
#HEREBY DISCLAIMS ALL SUCH WARRANTIES, INCLUDING WITHOUT LIMITATION, ANY WARRANTIES
#OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, QUIET ENJOYMENT, OR NON-INFRINGEMENT.
#Licensor reserves the right to change the Venue stated in the license to the UK.

from plasTeX.TeX import TeX
import getopt, sys
def usagemsg():
	print('''BrlTex by Michael Whapples\n
Usage:\tBrlTex (opts) inputfile
Options:
-c{code}\tUse {code} for braille output''')
try:
	opts, args = getopt.getopt(sys.argv[1:], 'c:', 'code=')
except getopt.GetoptError:
	usagemsg
	sys.exit(2)
if len(args) == 1:
	infilename = args[0]
else:
	usagemsg()
	sys.exit(2)
brlcode = 'bauk'
for o, a in opts:
	if o == "-c":
		brlcode = a
exec('from %s import Renderer' % brlcode)
if infilename[-4:].lower().find('.tex') < 0:
	Print('This does not seem to be a .tex file')
	sys.exit()
import os
if not os.path.exists(infilename):
	print('Cannot find %s' % infilename)
	sys.exit()
f = open(infilename)
texstr = f.read()
f.close()
outfilename = infilename[:-4] + '.brf'
tex = TeX()
tex.ownerDocument.config['files']['filename'] = outfilename
tex.ownerDocument.config['files']['split-level'] = -100
tex.input(texstr)
document = tex.parse()
Renderer().render(document)
