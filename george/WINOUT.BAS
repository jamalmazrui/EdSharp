' WinOut version 1.0
' Author: Anonymous Jhon
' Language: PowerBasic for Windows 7.0
' Begin Date: 09/27/2002

' Tom, I'm heavily commenting this code
' so you can see what is done and how.

' MetaStatements:
#COMPILE EXE ' compile to exe (default)
  ' other option is dll
#DIM ALL ' all variables, constants and arrays must be
  ' declared before use
' #includes come here such as win32api.inc, comndlg.inc and
  ' comnctrl.inc. none needed here
' constants here. % for numeric and $ for
  ' string e.g. %IdCancel = 100
$appTitle = "WinOut 1.0"

' main function it need not be declared
FUNCTION PBMAIN() AS LONG

  LOCAL parms AS STRING ' command line parameters
  LOCAL MsgString AS STRING ' message to be displayed
  LOCAL FileSpec AS STRING ' file to search
  LOCAL SrchString AS STRING ' string to search for
  LOCAL txt AS STRING ' file buffer
  LOCAL f AS LONG ' search result
  LOCAL fh AS LONG ' file handle
  ' local variables are not seen by any other sub or function and are
  ' deleted from the stack when sub or function exits.
  ' globals are available to all subs and functions
  ' statics can only be changed by calling sub or function and remain
  ' when sub or function exits.

  ' parse command line
  ' expected order: MsgString, FileSpec, SrchString
  parms = COMMAND$ ' get command line
  ' command$ is a system variable in basic containing command line
  ' parameters but not program path and name.
  ' how many parameters? find commas
  f = INSTR(parms, ",")
  IF f = 0 THEN f = LEN(parms) + 1 ' not found search not needed
  MsgString = LEFT$(parms, f - 1) ' all chars left of f
  INCR f ' f = f + 1
  ' if FileSpec is given, assume SrchString is too
  ' trim MsgString from parms
  IF f < LEN(parms) THEN Parms = MID$(Parms, f)
  ' hunt the comma
  f = INSTR(Parms, ",")
  ' if found get FileSpec and SrchString
  IF f > 0 THEN
    FileSpec = LEFT$(Parms, f - 1)
    SrchString = MID$(Parms, f + 1)
    ' read file into txt
    fh = FREEFILE ' next available file handle
    OPEN FileSpec FOR BINARY AS #fh
    GET$ #fh, LOF(fh), txt
    CLOSE #fh
    ' convert for non case sensitive search
    txt = LCASE$(txt)
    SrchString = LCASE$(SrchString)
    f = INSTR(SrchString, txt)
    ' if not found kill MsgString
    IF f = 0 THEN MsgString = ""
  END IF
  ' if msgString exists, display it.
  IF LEN(MsgString) THEN
    MSGBOX MsgString, , " "
  END IF
END FUNCTION
