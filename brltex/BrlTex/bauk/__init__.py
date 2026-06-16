#!/usr/bin/python
#Copyright (C) 2006 Michael Whapples, All Rights Reserved.
#Original Author: Michael Whapples, mwhapples@users.sourceforge.net
#Contributors: Alastair Irving,
#Unless linsor has explicitly granted you any other licensing terms, the contents of this file is subject to the RPL 
1.1, 
#available at the BrlTex website (http://brltex.sourceforge.net).
#All software distributed under the Licenses is provided strictly on an "AS IS" basis,
#WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, AND the licensor
#HEREBY DISCLAIMS ALL SUCH WARRANTIES, INCLUDING WITHOUT LIMITATION, ANY WARRANTIES
#OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, QUIET ENJOYMENT, OR NON-INFRINGEMENT.
#Licensor reserves the right to change the Venue stated in the license to the UK.

''''Renderer for plasTeX to give braille output'''
#
#This module will initially allow BAUK output
#Hopefully after a definite lisencing decision will allow modification for other codes
#Lets start with some variables that could be useful
#These are things like a math mode variable
class general:
	debug = False
	unknown = 0
class transdata:
	linelen = 40
	pagelen = 25
#def _eqnnum():
#	#hopefully we won't run out of numbers
#	for x in range(1, 10000):
#		yield x
class _eqnnum:
	val = 0
#a cache for translations, should speed things up
class cache:
	translatetext = {}
	translatemath = {}
	translatemathcomp = {}
class transmode:
	compact = False
	mathmode = False
#Initially define all the conversions
#the Renderer interfacing has to come at the end due to
#python requiring functions to be defined earlier than calls


#Functions for Greek letters
def alpha_handler(node):
	return unicode('.A')
def cap_alpha_handler(node):
	return unicode('_A')
def beta_handler(node):
	return unicode('.B')
def cap_beta_handler(node):
	return unicode('_B')
def gamma_handler(node):
	return unicode('.g')
def cap_gamma_handler(node):
	return unicode('_G')
def delta_handler(node):
	return unicode('.D')
def cap_delta_handler(node):
	return unicode('_D')
def epsilon_handler(node):
	return unicode('.E')
def cap_epsilon_handler(node):
	return unicode('_E')
def varepsilon_handler(node):
	return unicode('.E')
def phi_handler(node):
	return unicode('.F')
def cap_phi_handler(node):
	return unicode('_F')
def varphi_handler(node):
	return unicode('.F')
def eta_handler(node):
	return unicode('.:')
def cap_eta_handler(node):
	return unicode('_:')
def theta_handler(node):
	return unicode('.?')
def cap_theta_handler(node):
	return unicode('_?')
def vartheta_handler(node):
	return unicode('.?')
def iota_handler(node):
	return unicode('.I')
def cap_iota_handler(node):
	return unicode('_I')
def kappa_handler(node):
	return unicode('.K')
def cap_kappa_handler(node):
	return unicode('_K')
def lambda_handler(node):
	return unicode('.L')
def cap_lambda_handler(node):
	return unicode('_L')
def mu_handler(node):
	return unicode('.M')
def cap_mu_handler(node):
	return unicode('_M')
def nu_handler(node):
	return unicode('.N')
def cap_nu_handler(node):
	return unicode('_N')
def omicron_handler(node):
	return unicode('.O')
def cap_omicron_handler(node):
	return unicode('_O')
def pi_handler(node):
	return unicode('.P')
def cap_pi_handler(node):
	return unicode('_P')
def varpi_handler(node):
	return unicode('.P')
def rho_handler(node):
	return unicode('.R')
def cap_rho_handler(node):
	return unicode('_R')
def varrho_handler(node):
	return unicode('.R')
def sigma_handler(node):
	return unicode('.S')
def cap_sigma_handler(node):
	return unicode('_S')
def varsigma_handler(node):
	return unicode('.S')
def tau_handler(node):
	return unicode('.T')
def cap_tau_handler(node):
	return unicode('_T')
def upsilon_handler(node):
	return unicode('.U')
def cap_upsilon_handler(node):
	return unicode('_U')
def omega_handler(node):
	return unicode('.W')
def cap_omega_handler(node):
	return unicode('_W')
def xi_handler(node):
	return unicode('.x')
def cap_xi_handler(node):
	return unicode('_X')
def psi_handler(node):
	return unicode('.Y')
def cap_psi_handler(node):
	return unicode('_Y')
def zeta_handler(node):
	return unicode('.Z')
def cap_zeta_handler(node):
	return unicode('_Z')
def chi_handler(node):
	#Think this is correct, check it though
	return unicode('.[')
#Now add these to the renderer dictionary
#This dictionary is not the actual renderer part,
#a seperate one is assigned to first, so that all the data can be added in one go.
#This way it is also possible for different dictionaries to be defined,
#and then the one the user wants can be used.
#A possible use might be for other braille codes.
BAUKdict = {'alpha': alpha_handler, 'Alpha': cap_alpha_handler, 
'beta': beta_handler, 'Beta': cap_beta_handler, 
'gamma': gamma_handler, 'Gamma': cap_gamma_handler, 
'delta': delta_handler, 'Delta': cap_delta_handler, 
'epsilon': epsilon_handler, 'Epsilon': cap_epsilon_handler, 'varepsilon': varepsilon_handler, 
'phi': phi_handler, 'Phi': cap_phi_handler, 'varphi': varphi_handler, 
'eta': eta_handler, 'Eta': cap_eta_handler, 
'theta': theta_handler, 'Theta': cap_theta_handler, 'vartheta': vartheta_handler, 
'iota': iota_handler, 'Iota': cap_iota_handler, 
'kappa': kappa_handler, 'Kappa': cap_kappa_handler, 
'lambda': lambda_handler, 'Lambda': cap_lambda_handler, 
'mu': mu_handler, 'Mu': cap_mu_handler, 
'nu': nu_handler, 'Nu': cap_nu_handler, 
'omicron': omicron_handler, 'Omicron': cap_omicron_handler, 
'pi': pi_handler, 'Pi': cap_pi_handler, 'varpi': varpi_handler, 
'rho': rho_handler, 'Rho': cap_rho_handler, 'varrho': varrho_handler, 
'sigma': sigma_handler, 'Sigma': cap_sigma_handler, 'varsigma': varsigma_handler, 
'tau': tau_handler, 'Tau': cap_tau_handler, 
'upsilon': upsilon_handler, 'Upsilon': cap_upsilon_handler, 
'omega': omega_handler, 'Omega': cap_omega_handler, 
'xi': xi_handler, 'Xi': cap_xi_handler, 
'psi': psi_handler, 'Psi': cap_psi_handler, 
'zeta': zeta_handler, 'Zeta': cap_zeta_handler, 'chi': chi_handler}

#Now some operators
#Note: some of these won't appear as nodes due to them being simple text
#An example of these text symbols are +, -, /, etc
def frac_handler(node):
	#Changing this to cope with fracs such int/int properly
	#Some of this was tidied up by Alastair Irving
	correctlen = False
	if (len(node.attributes['numer']) == 1) and (len(node.attributes['denom']) == 1):
		correctlen = True
	if unicode(node.attributes['numer'].textContent).isdigit() and unicode(node.attributes['denom'].textContent).isdigit() and correctlen:
		outstr = ""
		outstr += translate.numtrans(node.attributes["numer"].textContent)
		outstr += translate.lownumtrans(node.attributes["denom"].textContent)
		return unicode(outstr)
	else:
		return unicode('(%s/%s)' % (unicode(node.attributes['numer']), unicode(node.attributes['denom'])))
#The next bit may seem a bit out of place here, but these are used for the \over command
#In future when I have the time to really tweak things to be perfect I may do it differently 
#and not need this here
def bgroup_handler(node):
	return unicode('(%s)' % unicode(node))
def over_handler(node):
	return unicode('/')
def div_handler(node):
	return unicode('/')
def cdot_handler(node):
	return unicode("'")
def times_handler(node):
	return unicode(';8')
def int_handler(node):
	return unicode('!')
def oint_handler(node):
	return unicode('@!')
def sum_handler(node):
	return unicode('_S')
def prod_handler(node):
	return unicode('_P')
def pm_handler(node):
	return unicode(';6-')
def ll_handler(node):
	return unicode('[[')
def le_handler(node):
	return unicode('[7')
def ge_handler(node):
	return unicode('7O')
def gg_handler(node):
	return unicode('OO')
def sim_handler(node):
	return unicode('_7')
def simeq_handler(node):
	return unicode('_77')
def approx_handler(node):
	return unicode('.7')
def ne_handler(node):
	return unicode('"7')
def propto_handler(node):
	return unicode('37')
def leftarrow_handler(node):
	return unicode('[3')
def rightarrow_handler(node):
	return unicode('3o')
def langle_handler(node):
	return unicode('.(')
def rangle_handler(node):
	return unicode('.)')
def lbrace_handler(node):
	return unicode('[')
def rbrace_handler(node):
	return unicode('O')
def left_handler(node):
	return u'<'
def right_handler(node):
	return u'>'
#I don't have the braille to hand to confirm the translations for the dot commands that follow in the symbols file.
def re_handler(node):
	return unicode('$re%s' % unicode(node))
def im_handler(node):
	return unicode('$im%s' % unicode(node))
#leave these slightly rarer symbols for later.
def infty_handler(node):
	return unicode('=')
def nabla_handler(node):
	return unicode('_0')
def partial_handler(node):
	#I think this should include the d
	return unicode('@D')
def prime_handler(node):
	return unicode('9')
#Not sure about surd, need to check
def percent_handler(node):
	return u'3p'
#The following may need to be tweaked, or other code may need a tweak
#to make it properly contract where possible to read correctly
#Hopefully this will be enough to be readable
def sqrt_handler(node):
	return unicode('%')
#Again add this to the dictionary
BAUKdict.update({'pm': pm_handler, 'frac': frac_handler, 'div': div_handler, 'bgroup': bgroup_handler, 'over': 
over_handler, 'cdot': 
cdot_handler, 'times': times_handler, 
'int': int_handler, 'oint': oint_handler, 'sum': sum_handler, 'prod': prod_handler,
'll': ll_handler, 'le': le_handler, 'ge': ge_handler, 'gg': gg_handler, 
'sim': sim_handler, 'simeq': simeq_handler, 'approx': approx_handler, 'ne': ne_handler, 'propto': propto_handler, 
'leftarrow': leftarrow_handler, 'rightarrow': rightarrow_handler,
'langle': langle_handler, 'rangle': rangle_handler, '{': lbrace_handler, '}': rbrace_handler,
'left': left_handler, 'right': right_handler,
'Re': re_handler, 'Im': im_handler,
'infty': infty_handler, 'nabla': nabla_handler, 'prime': prime_handler, 'partial': partial_handler, 
'%': percent_handler, 'sqrt': sqrt_handler})
#Continuing in the symbols file at trig functions
def sin_handler(node):
	return unicode('$S')
def cos_handler(node):
	return unicode('$C')
def tan_handler(node):
	return unicode('$T')
def arcsin_handler(node):
	return unicode('$@S')
def arccos_handler(node):
	return unicode('$@C')
def arctan_handler(node):
	return unicode('$@T')
def csc_handler(node):
	return unicode('$<')
def sec_handler(node):
	return unicode('$-')
def cot_handler(node):
	return unicode('$|')
def exp_handler(node):
	return unicode('$EXP')
def sinh_handler(node):
	return unicode('$HS')
def cosh_handler(node):
	return unicode('$HC')
def tanh_handler(node):
	return unicode('$HT')
def coth_handler(node):
	return unicode('$H|')

#Now the dictionary again
BAUKdict.update({'sin': sin_handler, 'cos': cos_handler, 'tan': tan_handler, 
'arcsin': arcsin_handler, 'arccos': arccos_handler, 'arctan': arctan_handler,
'csc': csc_handler, 'sec': sec_handler, 'cot': cot_handler,
'sinh': sinh_handler, 'cosh': cosh_handler, 'tanh': tanh_handler, 
'coth': coth_handler, 'exp': exp_handler})
#moving on,
#Not sure if this is correct for arg, but seems sensible
def arg_handler(node):
	return unicode('$arg')
def log_handler(node):
	#again this is probably not best, but readable,
	#should improve on this in the future
	return unicode('$L')
def ln_handler(node):
	#Note that there is lg and ln commands in LaTex that both mean this
	#this means that this should be applied to both
	#This can be done in the dictionary, hence there should not be two functions for this
	return unicode('$ln')
#in the symbols file there is an example for log to base,
#this uses the log command, so should not repeat, the subscript will handle the base
#May as well do this sub/superscript bit here
def subscript_handler(node):
	#There are special cases where the general rule is broken, 
	#Now should be corrected
	#Again tidied by Alastair Irving
	correctlen = False
	outstr = u'*'
	transmode.compact = True
	if len(node.attributes['arg']) == 1:
		correctlen = True
	if node.attributes['arg'].textContent.isdigit() and correctlen:
		outstr+=translate.lownumtrans(node.attributes['arg'].textContent)
	else:
		outstr = unicode('*%s]' % unicode(node.attributes['arg']))
	transmode.compact = False
	return outstr
#read comments about subscript, same applies to superscript
def superscript_handler(node):
	#Again tidied by Alastair Irving
	correctlen = False
	outstr = u'+'
	transmode.compact = True
	if len(node.attributes['arg']) == 1:
		correctlen = True
	if node.attributes['arg'].textContent.isdigit() and correctlen:
		outstr+=translate.lownumtrans(node.attributes['arg'].textContent)
	else:
		outstr = unicode('+%s]' % unicode(node.attributes['arg']))
	transmode.compact = False
	return outstr
def deg_handler(node):
	return unicode('0')
def min_handler(node):
	return unicode('$MIN')
def max_handler(node):
	return unicode('$MAX')
def det_handler(node):
	return unicode('$det')
def hbar_handler(node):
	#Officially there are cases this may be confusing, never found one
	#may consider later
	return unicode('H:')
#again the dictionary
BAUKdict.update({'arg': arg_handler, 'log': log_handler, 'ln': ln_handler, 'lg': ln_handler, 
'active::_': subscript_handler, 'active::^': superscript_handler,
'deg': deg_handler, 'min': min_handler, 'max': max_handler,
'det': det_handler, 'hbar': hbar_handler})
#Moving on to the accent part of the file
def hat_handler(node):
	#I think this is the right notation, will check though
	return unicode('%s?' % unicode(node))
def dot_handler(node):
	return unicode('%s\'' % unicode(node))
def ddot_handler(node):
	return unicode('%s-' % unicode(node))
def bar_handler(node):
	#This probably should check if it is one character
	return unicode('%s:' % unicode(node))
def vec_handler(node):
	return u'%s3O' % unicode(node)
def mathbf_handler(node):
	return unicode('@%s' % unicode(node))
def cal_handler(node):
	return unicode(node)
def mathrm_handler(node):
	return unicode(node)
BAUKdict.update({'bar': bar_handler, 'hat': hat_handler, 
'dot': dot_handler, 'ddot': ddot_handler,
'vec': vec_handler, 'mathbf': mathbf_handler, 'cal': cal_handler, 'mathrm': mathrm_handler})
#Moving on again,
#there are some things in this that might be math dependent for braille output, 
#in these cases use the variable _math_mode
def it_handler(node):
	#I think italics are done like this,
	#in fact I think all emphasised text should be done like this
	#so at the moment all bold, italic and underline will be pointed here
	if len(unicode(node)) > 1:
		return unicode('..%s' % unicode(node).replace(' ',' ..'))
	elif len(unicode(node)) == 1:
		return unicode('.%s' % unicode(node))
	else:
		return u''
BAUKdict.update({'bf': it_handler, 'it': it_handler,
'textit': it_handler, 'textbf': it_handler})
#The maths environments
def math_handler(node):
	transmode.mathmode = True
	outputtext = unicode(node)
	transmode.mathmode = False
	return outputtext
def displaymath_handler(node):
	transmode.mathmode = True
	outputtext = unicode(node)
	transmode.mathmode = False
	return outputtext
def equation_handler(node):
	#_eqnnum.val += 1
	transmode.mathmode = True
	outputtext = unicode(node)
	transmode.mathmode = False
	return outputtext
BAUKdict.update({'math': math_handler, 'displaymath': displaymath_handler, 'equation': equation_handler})
#Lets see to the paragraphs and sections
def par_handler(node):
	#There are certain cases where we don't want indent
	outputtext = ''
	if node.hasChildNodes():
		if node.childNodes[0].nodeName == 'array':
			return unicode('\n%s' % unicode(node))
		else:
			outputtext = unicode('\n  %s' % unicode(node))
	else:
		outputtext = unicode('\n  %s' % unicode(node))
	#This corrects if spaces have got in front of the beginning
	while outputtext[3].isspace():
		outputtext = '\n  ' + outputtext[4:]
	return outputtext
def section_handler(node):
	outtext = ''
	if unicode(node).find('\n\n') == 0:
		outtext = unicode('\n\n   %s%s' % (node.attributes['title'], unicode(node)))
	else:
		outtext = unicode('\n\n   %s\n%s' % (node.attributes['title'], unicode(node)))
	nextnode = unicode(node.nextSibling)
	if (len(nextnode) > 0) and (nextnode[0] == '\n'):
		outtext = outtext[:-1]
	#This corrects starting spacing
	while outtext[5].isspace():
		outtext = '\n\n   ' + outtext[6:]
	return outtext
#The document one,
#this should tidy up the ends
def document_handler(node):
	outtext = unicode(node)
	while outtext[0].isspace():
		outtext = outtext[1:]
	return outtext
#I think the default behaviour for tex is not to show anything on the \newcommand
def newcommand_handler(node):
	return u''

BAUKdict.update({'par': par_handler, 'section': section_handler, 'subsection': section_handler, 'subsubsection': section_handler,
'document': document_handler, 'newcommand': newcommand_handler})
#Now some lists
def enumerate_handler(node):
	outtext = ''
	itemcount = 1
	for item in node:
		outtext = outtext + translate.numtrans(itemcount) + ',4 '
		for par in item:
			outtext = outtext + unicode(par)
		outtext = outtext + '\n  '
		itemcount = itemcount + 1
	return outtext[:-3]
def itemize_handler(node):
	outtext = ''
	for item in node:
		for par in item:
			outtext = outtext + unicode(par)
		outtext = outtext + '\n  '
	return outtext[:-3]
BAUKdict.update({'enumerate': enumerate_handler, 'itemize': itemize_handler})
def tabular_handler(node):
	#We will need to get hold of the contents of each cell before we can output.
	#This will aid with processing,
	#to check things like the table will fit to the page,
	#If it won't fit, it should then try and do it in paragraph form
	#This will hold all cells
	cells = []
	#this will hold column lengths
	collen = [0]
	#Now get the cells
	for row in node:
		rows = []
		for cell in row:
			outtext = ''
			for par in cell:
				outtext = outtext + unicode(par)
			rows.append(outtext)
		#Check if the column lengths have changed
		if len(collen) != len(rows):
			collen = []
			for x in rows:
				collen.append(len(x))
		else:
			for x, y in enumerate(rows):
				if collen[x] < len(y):
					collen[x] = len(y)
		cells.append(rows)
	#Now find if it fits pagewidth
	tablewidth = 0
	for x in collen:
		tablewidth = tablewidth + x
	outtext = ''
	if tablewidth < _linelen:
		#Lets try and layout this correctly
		for x in cells:
			for colnum, y in enumerate(x):
				outtext = outtext + y
				#Now pad it to the next column
				padlen = collen[colnum] - len(y) + 1
				for padcount in range(0,padlen -1):
					outtext = outtext + ' '
			outtext = outtext + '\n'
	else:
		print('\nWARNING: Table too big for page, putting in paragraph form')
		outtext = translate.text('The following table is done in paragraph form. Each row is a paragraph, and each cell of the row is seperated from the next by a ;')
		outtext = outtext + '\n'
		for x in cells:
			for y in x:
				outtext = outtext + y + '; '
			outtext = outtext + '\n'
	return unicode(outtext)
BAUKdict.update({'array': tabular_handler, 'tabular': tabular_handler})
#Now make the renderer part
import platform
from plasTeX.Renderers import Renderer
import translate
class brlrenderer(Renderer):
	def default(self, node):
		#This will handle all that has not been defined in the dictionary
		#so something useful needs to be done with this
		#at the moment I will just get it to dump the plasTeX names in the output
		if general.debug == False:
			general.unknown += 1
			return u'%s' % node.source
		s = []
		s.append('<%s>' % node.nodeName)
		if node.hasAttributes():
			for key, value in node.attributes.items():
				if key == 'self':
					continue
				s.append(' %s=%s\n' % (key, unicode(value)))
		s.append(unicode(node))
		s.append('</%s>' % node.nodeName)
		return u'\n'.join(s)
	def textDefault(self, node):
		#This handles all the text
		#At the moment some simple translation
		#in the future this should be completed
		#this may use external translators
		nodetext = unicode(node.nodeValue)
		#We need to replace some characters that can't go to ascii
		nodetext = nodetext.replace(u'\u2019', "'")
		if transmode.mathmode == True:
			if (not transmode.compact) and (node.previousSibling != None) and (translate.mathpreferspbf.find(nodetext[0]) >= 0):
				nodetext = ' %s' % nodetext
			if cache.translatemath.has_key(nodetext) and (not transmode.compact):
				translated = cache.translatemath[nodetext]
			elif cache.translatemathcomp.has_key(nodetext) and transmode.compact:
				translated = cache.translatemathcomp[nodetext]
			else:
				translated = translate.math(nodetext, transmode.compact)
				if transmode.compact:
					cache.translatemathcomp.update({nodetext: translated})
				else:
					cache.translatemath.update({nodetext: translated})
			return translated
		else:
			if cache.translatetext.has_key(nodetext):
				translated = cache.translatetext[nodetext]
			else:
				translated = translate.text(nodetext)
				cache.translatetext.update({nodetext: translated})
			return translated
	def processFileContent(self, document, content):
		#Firstly get the line lengths correct
		s = content.split('\n')
		outputtext = ''
		linenum = 1
		pagenum = 1
		for x in s:
			words = x.split(' ')
			outputline = ''
			for y in words:
				#check the line length, it is different for line1
				if linenum != 1:
					linelength = transdata.linelen
				else:
					linelength = transdata.linelen - 8
				if (len(outputline) + len(y)) < transdata.linelen:
					outputline += y + ' '
				elif (len(outputline) + len(y)) == transdata.linelen:
					outputline += y
				else:
					if linenum != 1:
						outputtext += outputline + '\n'
						linenum += 1
					else:
						pagenumstr = translate.numtrans(str(pagenum))
						while len(outputline) + len(pagenumstr) < transdata.linelen:
							outputline += ' '
						outputline += pagenumstr
						outputtext += outputline + '\n'
					linenum += 1
					if linenum > transdata.pagelen:
						linenum = 1
						pagenum += 1
						outputtext += '\f'
					if y.isspace() or (y == ''):
						outputline = ''
					else:
						outputline = y + ' '
			if linenum != 1:
				outputtext += outputline + '\n'
			else:
				pagenumstr = translate.numtrans(str(pagenum))
				while len(outputline) + len(pagenumstr) < transdata.linelen:
					outputline += ' '
				outputtext += outputline + pagenumstr + '\n'
			linenum += 1
			if linenum > transdata.pagelen:
				linenum = 1
				pagenum += 1
				outputtext += '\f'
		#Now give those windows users the line endings for Windows
		if platform.system() == 'Windows':
			outputtext = outputtext.replace('\n', '\r\n')
		return outputtext
def Renderer(brldict=BAUKdict):
	return brlrenderer(brldict)

