/* liblouisxml Braille Transcription Library

   This file may contain code borrowed from the Linux screenreader
   BRLTTY, copyright (C) 1999-2006 by
   the BRLTTY Team

   Copyright (C) 2004, 2005, 2006
   ViewPlus Technologies, Inc. www.viewplus.com
   and
   JJB Software, Inc. www.jjb-software.com
   All rights reserved

   This file is free software; you can redistribute it and/or modify it
   under the terms of the Lesser or Library GNU General Public License 
   as published by the
   Free Software Foundation; either version 3, or (at your option) any
   later version.

   This file is distributed in the hope that it will be useful, but
   WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the 
   Library GNU General Public License for more details.

   You should have received a copy of the Library GNU General Public 
   License along with this program; see the file COPYING.  If not, write to
   the Free Software Foundation, 51 Franklin Street, Fifth Floor,
   Boston, MA 02110-1301, USA.

   Maintained by John J. Boyer john.boyer@jjb-software.com
   */

#ifndef LIBLOUISXML_H_
#define LIBLOUISXML_H_
#include <liblouis.h>

#ifdef __cplusplus
extern "C"
{
#endif				/* __cplusplus */

/* Function prototypes are documented briefly below. For more extensive 
documentation see liblouisxml.html. */

char * EXPORT_CALL lbx_version (void);
/* Returns the version of liblouisxml. */

  void * EXPORT_CALL lbx_initialize (const char *const configFileName, 
const char
			const *logFileName, const char const *settingsString);

/* This function initializes the libxml2 library, runs canonical.cfg and
processes the configuration file given in configFileName, sets up a log
file if logFileName is not NULL and processes the settings in
settingsString if this is not null. It returns a pointer to the UserData
structure. This pointer is void and must be cast to (UserData *) in the
calling program. To access the information in this structure you must
include louisxml.h */

  typedef enum
  {
    dontInit = 1,
    htmlDoc = 2
  } processingModes;

  int EXPORT_CALL lbx_translateString
    (const char *const configFileName,
     char *inbuf, widechar * outbuf, int *outlen, unsigned int mode);

/* This function takes a well-formed xml expression in inbuf and
translates it into a string of 16- or 32-bit braille characters in
outbuf.  The xml expression must be immediately followed by a zero or
null byte. If it does not begin with an xul header, one is added. The
header is specified by the xmlHeader line in the configuration file. If
no such line is present, a default header specifying UTF-8 encoding is
used. The configFileName parameter points to a configuration file.
canonical.cfg is processed before this file. Note that the *outlen
parameter is a pointer to an integer. When the function is called, this
integer contains the maximum output length. When it returns, it is set
to the actual length used.  The mode parameter is used to pass options
to liblouisxml which are applied before any configuration fpe is
processed. For now, if mode is 0, a full initialization is cone. If it
is 1 only a few things are reset to prepare for a new document. The
function returns 1 if no errors were encountered and a negative number
if a conplete translation could not be done.  */


  int EXPORT_CALL lbx_translateFile (char *configFileName, char 
*inputFileName,
			 char *outputFileName, unsigned int mode);

  int EXPORT_CALL lbx_translateTextFile (char *configFileName, char 
*inputFileName,
			     char *outputFileName, unsigned int mode);
  int EXPORT_CALL lbx_backTranslateFile (char *configFileName, char
			     *inputFileName,
			     char *outputFileName, unsigned int mode);

  void EXPORT_CALL lbx_free (void);

/* This function should be called at the end of the application to free
all memory allocated by liblouisxml or liblouis. */

#ifdef __cplusplus
}
#endif				/* __cplusplus */


#endif				/*LIBLOUISXML_H_ */
