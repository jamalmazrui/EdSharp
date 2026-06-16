#Copyright (C) 2006 Michael Whapples, All Rights Reserved.
#Original Author: Michael Whapples, mwhapples@users.sourceforge.net
#Unless linsor has explicitly granted you any other licensing terms, the contents of this file is subject to the RPL 
#1.1, 
#available at the BrlTex website (http://brltex.sourceforge.net).
#All software distributed under the Licenses is provided strictly on an "AS IS" basis,
#WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, AND the licensor
#HEREBY DISCLAIMS ALL SUCH WARRANTIES, INCLUDING WITHOUT LIMITATION, ANY WARRANTIES
#OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, QUIET ENJOYMENT, OR NON-INFRINGEMENT.
#Licensor reserves the right to change the Venue stated in the license to the UK.

#Some stuff for later on
#A dictionary for lowered numbers, such as those used in denominators
#This is not really needed for BAUK, but this will aid if other codes don't use the what is defined below.
#Not the best variable name, some with a better imagination, suggestion?
lownum = {'1': '1', '2': '2', '3': '3', '4': '4', '5': '5', '6': '6', '7': '7', '8': '8', '9': '9', '0': '0'}
#need the normal number chars
highnum = {'1': 'a', '2': 'b', '3': 'c', '4': 'd', '5': 'e', '6': 'f', '7': 'g', '8': 'h', '9': 'i', '0': 'j'}
#also need a number prefix
numprefix = '#'
#a function to do number conversion.
def numtrans(numint):
        numstr = str(numint)
        x = 0
        outnum = numprefix
        while x < len(numstr):
                outnum = outnum + highnum[numstr[x]]
                x = x + 1
        return outnum

#A function to convert lower numbers (added by Alasstair Irving)
def lownumtrans(numint):
        numstr = str(numint)
        x=0
        outnum=""
        while x < len(numstr):
                outnum += lownum[numstr[x]]
                x+=1
        return outnum

#This is a new version of translate.py, it will attempt to use nfbtrans for the translator.
#This should be seen as a quick fix, as there are too many calls to the executable to make this effecient.
import os
import string
import sys
def grade1(inputtext):
	#need to change the working directory to where this file is, it will lead to better multiple code support
	#as different codes can have different nfbtrans settings files
	olddir = os.getcwd()
	os.chdir(os.path.dirname(__file__))
	leadingspace = ''
	trailingspace = ''
	#see to spaces at either end of the string, nfbtrans will do formatting if there is.
	while (len(inputtext) > 0) and (inputtext[0] == ' '):
		leadingspace = leadingspace + ' '
		inputtext = inputtext[1:]
	while (len(inputtext) > 0) and (inputtext[-1] == ' '):
		trailingspace = trailingspace + ' '
		inputtext = inputtext[:-1]
	f = open('brltemp.txt', 'w')
	f.write(inputtext)
	f.close()
	if os.path.exists(os.path.join(os.pardir, 'nfbtrans')) or os.path.exists(os.path.join(os.path.pardir, 'nfbtrans.exe')):
		 os.system('%s brltemp.txt >brlout.brf' % os.path.join(os.path.pardir, 'nfbtrans'))
	else:
		nfbsuccess = os.system('nfbtrans brltemp.txt >brlout.brf')
		if nfbsuccess != 0:
			print('\nnfbtrans cannot be found.')
			sys.exit()
	outputtext = ''
	f = open('brlout.brf')
	for line in f:
		outputtext = outputtext + line[:-1] + ' '
	f.close()
	#an extra space is being inserted by the line joiner
	outputtext = outputtext[:-1]
	#remove the page break characters
	outputtext = outputtext.replace('\x0c', '')
	#clean up
	os.remove('brlout.brf')
	os.remove('brltemp.txt')
	os.chdir(olddir)
	#put the spaces and the leading dot 6 back
	outputtext = '%s%s%s' % (leadingspace, outputtext[:-1], trailingspace)
	return '%s' % outputtext
#The text translator
def text(inputtext):
	if string.punctuation.find(inputtext[0]) != -1:
		return ',%s' % grade1(inputtext)
	else:
		return grade1(inputtext)
#the math text translator
def math(inputtext, compact=False):
	#until I have code to sort this out it will just go through the standard translator
	return math2brl(inputtext, compact)
#We will try and do a translator for the text of math mode
#This should be simple, as there are no contractions, and about everything is a simple substitution.
#It is currently being done as a seperate function to math as this will allow testing without affecting the current 
#situation
#We will need some chars that should not have a space after
mathnospaf = '''([{+=-/| '''
#Now for before
mathnospbf = ''')]'}/| '''
mathpreferspbf = '=+-'
mathsimpletrans = {'+': ';6', '=': ';7', '-': ';-', "'": '9', '(': '<', ')': '>', '<': '[', '>': 'O', '[': '(', ']': 
')', '|': '_', '!': '6'}
def math2brl(inputtext, compact=False):
	outputtext = ''
	prevchar = ''
	prevprevchar = ''
	nextchar = ''
	for x, character in enumerate(inputtext):
		#Now get prevprevchar, prevchar and nextchar
		if (x -2) >= 0:
			prevprevchar = inputtext[x-2]
		if (x - 1) >= 0:
			prevchar = inputtext[x-1]
		if (x + 1) < len(inputtext):
			nextchar = inputtext[x+1]
		#Find unnecessary spaces
		if character.isspace() and (mathnospaf.find(prevchar) >=0) and (prevchar != ''):
			continue
		elif character.isspace() and (mathnospbf.find(nextchar) >= 0):
			continue
		#Now for character translation
		elif mathsimpletrans.has_key(character):
			if (mathpreferspbf.find(character) >= 0) and (not compact) and (mathnospaf.find(prevchar) < 0):
				outputtext += ' %s' % mathsimpletrans[character]
			else:
				outputtext += mathsimpletrans[character]
		elif character.isupper() and nextchar.isupper() and (not prevchar.isupper()):
			outputtext += ',,%s' % character
		elif (not prevchar.isupper()) and character.isupper() and (not nextchar.isupper()):
			outputtext += ',%s' % character
		elif ((not prevchar.isdigit()) and (prevchar != '.')) and character.isdigit():
			outputtext += numprefix + highnum[character]
		elif ((prevchar == '.') or prevchar.isdigit()) and character.isdigit():
			outputtext += highnum[character]
		elif prevchar.isdigit() and (character == '.') and nextchar.isdigit():
			outputtext += '1'
		elif prevchar.isdigit() and (character.islower() and character.isalpha):
			outputtext += ';%s' % character
		else:
			outputtext += character
	return outputtext
