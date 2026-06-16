#!/usr/bin/python
#This will download all pre-requesites for BrlTex (excluding python).
#First make the directories to put things in
import sys
import os
import urllib2
olddir = os.getcwd()
try:
	if not os.path.exists('temp'):
		os.makedirs('temp')
	if not os.path.exists(os.path.join('temp', 'nfbtrans')):
		os.makedirs(os.path.join('temp', 'nfbtrans'))
	#Plastex will make its own when extracted
	if not os.path.exists(os.path.join('temp', 'nfbtrans.zip')):
		#Now download nfbtrans
		print('Downloading NFBTRANS')
		urlfile = urllib2.urlopen("http://www.nfbnet.org/files/nfbtrans/NFBTR770.ZIP")
		s = urlfile.read()
		urlfile.close()
		f = open(os.path.join('temp', 'nfbtrans.zip'), 'wb')
		f.write(s)
		f.close()
	#Now try to install it
	print('Now installing nfbtrans')
	import zipfile
	z = zipfile.ZipFile(os.path.join('temp', 'nfbtrans.zip'), 'r')
	for x in z.namelist():
		s = z.read(x)
		f = open(os.path.join('temp', 'nfbtrans', x), 'wb')
		f.write(s)
		f.close()
	z.close()
	#Need to check OS
	import platform
	if platform.system() == 'Linux':
		os.chdir(os.path.join('temp', 'nfbtrans'))
		os.rename('MAKEFILE', 'Makefile')
		os.system('make lowercase')
		os.system('make linux')
		if os.path.exists(os.path.join(olddir, 'nfbtrans')):
			os.remove(os.path.join(olddir, 'nfbtrans'))
		os.rename('nfbtrans', os.path.join(olddir, 'nfbtrans'))
	elif platform.system() == 'Windows':
		if os.path.exists('nfbtrans.exe'):
			os.remove('nfbtrans.exe')
		os.rename(os.path.join('temp', 'nfbtrans', 'nfbtrans.exe'), 'nfbtrans.exe')
	else:
		print('This script doesn ot support your OS, please follow the manual install instructions in INSTALL')
		sys.exit()
	os.chdir(olddir)
	if not os.path.exists(os.path.join('temp', 'plastex.tar.gz')):
		#Now get plastex
		print('Downloading plasTeX')
		urlfile = urllib2.urlopen("http://kent.dl.sourceforge.net/sourceforge/plastex/plastex-0.6.1.tgz")
		#The following url could be used instead
		#http://umn.dlsourceforge.net/sourceforge/plastex/plastex-0.6.1.tgz
		s = urlfile.read()
		urlfile.close()
		f = open(os.path.join('temp', 'plastex.tar.gz'), 'wb')
		f.write(s)
		f.close()
	#Now the install
	print('Installing plasTeX')
	os.chdir('temp')
	import tarfile
	plastexfile = tarfile.open('plastex.tar.gz')
	for tarinfo in plastexfile:
		plastexfile.extract(tarinfo)
	plastexfile.close()
	print('patching plasTeX')
	os.chdir('plastex-0.6.1')
	f = open(os.path.join('plasTeX', 'Base', 'LaTeX', 'Entities.py'), 'r')
	entstr = f.read()
	f.close()
	entstr = entstr.replace('''e.parse(os.path.join(os.path.dirname(__file__),'ent.xml'))''', '''#e.parse(os.path.join(os.path.dirname(__file__),'ent.xml'))''')
	f = open(os.path.join('plasTeX', 'Base', 'LaTeX', 'Entities.py'), 'w')
	f.write(entstr)
	f.close()
	if platform.system() == 'Linux':
		os.system('python setup.py install')
	elif platform.system() == 'Windows':
		os.system('setup.py install')
	#Check that plastex installed OK
	import plasTeX
	os.chdir(olddir)
	#now clean up
	print('Cleaning up')
	def dirempty(fn):
		for file in os.listdir(fn):
			if os.path.isdir(os.path.join(fn, file)):
				dirempty(os.path.join(fn, file))
				os.rmdir(os.path.join(fn, file))
			else:
				if fn != 'temp':
					os.remove(os.path.join(fn, file))
	dirempty('temp')
	#os.rmdir('temp')
except:
	Print('''There was a problem installing. This script needs to access the internet if you haven't already got the 
required files.''')
print('Now you should be able to run the BrlTex.py script.')

