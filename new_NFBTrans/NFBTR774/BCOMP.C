#define LINT_ARGS
#include	<stdio.h>
#include <io.h>
#include <stdlib.h>
#include	<string.h>
#include <ctype.h>
#include <fcntl.h>
#include <conio.h>
#include <malloc.h>
#include <memory.h>
#include <process.h>
#include <time.h>

void trim_trail(char *);

void trim_trail(char *string)
{
  int l = strlen(string) - 1;
  do
  {
    if (string[l] == ' ' || string[l] == '\n')
      string[l] = 0;
    else
      break;
    l--;
  }
  while (l >= 0);
}		   /* trim_trail */

void main(int, char **);

void main(int argc, char **argv)
{
  char line1[256], line2[256];
  int file_arg = 1, report_all_errors = 0, i,ignore_case = 0;
  long line_num = 0, start_line = 1, end_line = 99999l;
  FILE *f1, *f2;
  if (argc > 1 && argv[1][0] == '-')
  {		   /* - */
    file_arg++;
    for(i=1; argv[1][i]; i++)
    switch (argv[1][i])
    {
	case 'a':
      report_all_errors = 1;
      break;
      case 'i':
      ignore_case = 1;
      break;
}/*switch*/
  }		   /* - */
  if (argc < file_arg + 2)
  {		   /* error */
    fprintf(stderr, "BCOMP Version 1.02 September 29, 1997\n");
    fprintf(stderr, "Usage: BCOMP [-ai] FILE1 FILE2 [StartLine] [EndLine]\n");
    exit(1);
  }
  if ((f1 = fopen(argv[file_arg], "rt")) == NULL)
  {		   /* error */
error:
    fprintf(stderr, "%s not found\n", argv[file_arg]);
    exit(1);
  }		   /* error */
  file_arg++;
  if ((f2 = fopen(argv[file_arg], "rt")) == NULL)
    goto error;
  if (argc > file_arg + 1)
    start_line = atol(argv[file_arg + 1]);
  if (argc > file_arg + 2)
    end_line = atol(argv[file_arg + 2]);
  for (;;)
  {
    if ((fgets(line1, 250, f1)) == NULL)
      break;
    fgets(line2, 250, f2);
    line_num++;
    if (line_num < start_line)
      continue;
    if (line_num > end_line)
      break;
    trim_trail(line1);
    trim_trail(line2);
    if(ignore_case)
    {
    strlwr(line1);
    strlwr(line2);
}
    if (strcmp(line1, line2))
    {		   /* lines don't compare */
	i = 0;
	while(line1[i]	== line2[i])
	i++;
      printf("Error in line %ld position %d\n", line_num, i + 1);
      printf("%s\n%s\n", line1, line2);
      if (report_all_errors)
	continue;
      break;
    }		   /* lines don't compare */
  }		   /* while */
  fclose(f1);
  fclose(f2);
}		   /* main */
