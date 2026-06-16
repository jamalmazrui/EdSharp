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

#ifndef liblouisxml_h
#define liblouisxml_h
#include <libxml/parser.h>
#include "liblouisxml.h"
#include "sem_enum.h"
#include "transcriber.h"

#ifdef _WIN32
#define strcasecmp _strnicmp
#define snprintf _snprintf
#define vsnprintf _vsnprintf
#endif

#define CHARSIZE sizeof (widechar)
#define BUFSIZE 8192
#define MAX_LENGTH BUFSIZE - 4
#define MAX_TRANS_LENGTH 2 * BUFSIZE - 4
#define MAXNAMELEN 256
#define MAXNUMLEN 32
#define STACKSIZE 100

typedef enum
{
  utf8 = 0,
  utf16,
  utf32,
  ascii8
} Encoding;

typedef enum
{
  plain = 0,
  html
} TextFormat;

typedef enum
{
textDevice = 0,
browser
} FormatFor;

typedef struct
{				/*user data */
  FILE *inFile;
  FILE *outFile;
  int text_length;
  int translated_length;
  int paragraph_interrupted;
  int interline;
  int has_math;
  int has_comp_code;
  int has_chem;
  int has_graphics;
  int has_music;
  int has_cdata;
  Encoding input_encoding;
  Encoding output_encoding;
  Encoding input_text_encoding;
  TextFormat back_text;
  int back_line_length;
  FormatFor format_for;
  int contents;
  int has_contentsheader;
  int cells_per_line;
  int lines_per_page;
  int beginning_braille_page_number;
  int interpoint;
  int print_page_number_at;
  int braille_page_number_at;
  int hyphenate;
  int internet_access;
  int new_entries;
  char *mainBrailleTable;
  char *inbuf;
  int inlen;
  widechar *outbuf;
  int outlen;
  int outlen_so_far;
  int lines_on_page;
  int braille_page_number;
  int prelim_pages;
  int paragraphs;
  int braille_pages;
  int print_pages;
  char path_list[4 * MAXNAMELEN];
  char *writeable_path;
  char string_escape;
  char file_separator;
  char line_fill;
  char lit_hyphen[5];
  char comp_hyphen[5];
  widechar running_head[MAXNAMELEN / 2];
  widechar footer[MAXNAMELEN / 2];
  int running_head_length;
  int footer_length;
  char contracted_table_name[MAXNAMELEN];
  char uncontracted_table_name[MAXNAMELEN];
  char compbrl_table_name[MAXNAMELEN];
  char mathtext_table_name[MAXNAMELEN];
  char mathexpr_table_name[MAXNAMELEN];
  char edit_table_name[MAXNAMELEN];
  char interline_back_table_name[MAXNAMELEN];
  char semantic_files[MAXNAMELEN];
  widechar print_page_number[MAXNUMLEN];
  widechar braille_page_string[MAXNUMLEN];
  char lineEnd[8];
  char pageEnd[8];
  char fileEnd[8];
  /*stylesheet */
  StyleType document_style;
  StyleType para_style;
  StyleType heading1_style;
  StyleType heading2_style;
  StyleType heading3_style;
  StyleType heading4_style;
  StyleType section_style;
  StyleType subsection_style;
  StyleType table_style;
  StyleType volume_style;
  StyleType titlepage_style;
  StyleType contentsheader_style;
  StyleType contents1_style;
  StyleType contents2_style;
  StyleType contents3_style;
  StyleType contents4_style;
  StyleType code_style;
  StyleType quotation_style;
  StyleType attribution_style;
  StyleType indexx_style;
  StyleType glossary_style;
  StyleType biblio_style;
  StyleType list_style;
  StyleType caption_style;
  StyleType exercise1_style;
  StyleType exercise2_style;
  StyleType exercise3_style;
  StyleType directions_style;
  StyleType stanza_style;
  StyleType line_style;
  StyleType spatial_style;
  StyleType arith_style;
  StyleType note_style;
  StyleType trnote_style;
  StyleType dispmath_style;
  StyleType disptext_style;
  StyleType matrix_style;
  StyleType music_style;
  StyleType graph_style;
  StyleType graphlabel_style;
  StyleType dedication_style;
  StyleType blanklinebefore_style;
  StyleType style1_style;
  StyleType style2_style;
  StyleType style3_style;
  StyleType style4_style;
  StyleType style5_style;
  StyleType scratch_style;
/*end of stylesheet*/
  int top;
  sem_act stack[STACKSIZE];
  char xml_header[BUFSIZE];
  widechar text_buffer[2 * BUFSIZE];
  widechar translated_buffer[2 * BUFSIZE];
  unsigned char typeform[2 * BUFSIZE];
} UserData;
extern UserData *ud;

/* Function definitions */
StyleType *style_cases (sem_act action);
sem_act find_semantic_number (const char *semName);
int file_exists (const char *completePath);
int find_file (const char *fileName, char *filePath);
int set_paths (const char *configPath);
int read_configuration_file (const char *configFileName,
			     const char *logFileName, const char 
*settingsString, unsigned int mode);
int config_compileSettings (const char *fileName);
int examine_document (xmlNode * node);
int transcribe_document (xmlNode * node);
int transcribe_math (xmlNode * node, int action);
int transcribe_computerCode (xmlNode * node, int action);
int transcribe_cdataSection (xmlNode *node);
int transcribe_paragraph (xmlNode * node, int action);
int transcribe_chemistry (xmlNode * node, int action);
int transcribe_graphic (xmlNode * node, int action);
int transcribe_music (xmlNode * node, int action);
int compile_semantic_table (xmlNode * rootElement);
sem_act set_sem_attr (xmlNode * node);
sem_act get_sem_attr (xmlNode * node);
sem_act push_sem_stack (xmlNode * node);
sem_act pop_sem_stack ();
void destroy_semantic_table (void);
void append_new_entries (void);
int insert_code (xmlNode *node, int which);
xmlChar *get_attr_value (xmlNode * node);
int change_table (xmlNode * node);
int initialize_contents (void);
int start_heading (sem_act action, widechar *translatedBuffer, int 
translatedLength);
int finish_heading (sem_act action);
int make_contents (void);
void do_reverse (xmlNode *node);
int do_boxline (xmlNode *node);
int do_newpage (void);
int do_blankline (void);
int do_softreturn (void);
int do_righthandpage (void);
int do_configstring (xmlNode *node);
#endif /*louisxml_h*/
