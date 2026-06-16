
include "hjconst.jsh"

globals int initialised,
int ProcessMaths,
object latex_access,
object matrix,
int row,
int column


Void Function AutoStartEvent ()
if !initialised then
let latex_access=CreateObject ("latex_access")
let initialised=true
endif
EndFunction

Void Function SayLine ()
if  ProcessMaths then
var string input
let input = GetLine ()

if stringisblank(input) then
let input = "blank"
else
let input = latex_access.speech(input)
let input = StringReplaceSubstrings (input, "&", "&amp;")
let input = StringReplaceSubstrings (input, "<sub>", smmGetStartMarkupForAttributes (attrib_subscript|attrib_text))
let input = StringReplaceSubstrings (input, "</sub>", smmGetEndMarkupForAttributes (attrib_subscript|attrib_text))
let input = StringReplaceSubstrings (input, "<bold>", smmGetStartMarkupForAttributes (attrib_bold|attrib_text))
let input = StringReplaceSubstrings (input, "</bold>", smmGetEndMarkupForAttributes (attrib_bold|attrib_text))
endif
Say (input, ot_selected_item, true)
else
SayLine ()
endif
EndFunction





Script ToggleMaths ()
if ProcessMaths then
let ProcessMaths = false
SayMessage (ot_status, "maths to be read as plain latex", "Processing off")
else
let ProcessMaths = true
SayMessage(OT_STATUS,"Maths to be processed to a more verbal form","Processing on")
endif


EndScript

Script ToggleDollarsNemeth ()
var bool result 
let result=latex_access.toggle_dollars_nemeth()
if result then 
SayMessage (ot_status, "Dollars will now be ignored in nemeth", "nemeth dollars off")
else
SayMessage (ot_status, "Dollars will now be shown in nemeth", "nemeth dollars on")

endif
EndScript


Script ToggleDollarsSpeech ()
var bool result 
let result=latex_access.toggle_dollars_speech()
if result then 
SayMessage (ot_status, "Dollars will now be ignored in speech", "speech dollars off")
else
SayMessage (ot_status, "Dollars will now be shown in speech", "speech dollars on")
endif
EndScript





Int Function BrailleBuildLine ()
if  ProcessMaths then
var string input
let input = GetLine()
let input = StringReplaceSubstrings (input, "scroll down symbol", "")
let input=StringTrimTrailingBlanks (input)
let input = latex_access.nemeth(input)
; now sort out bad dots 456 
let input =StringReplaceSubstrings (input, "_", "\127") 
BrailleAddString (input, 0, 0, 0)

endif
return true
EndFunction




Script InputMatrix ()
let matrix=CreateObject ("latex_access_matrix")
let row=1
let column=1
matrix.tex_init(GetSelectedText ())
var string msg
let msg ="Initialised "
let msg=msg+inttostring(matrix.rows)
let msg=msg+" by "
let msg=msg+inttostring(matrix.columns)
let msg=msg +" matrix"
saystring(msg)
EndScript


Script MatrixRight ()
if column < matrix.columns then
let column = column+1
saystring(matrix.get_cell(row,column))
else
saystring("end row")
endif
EndScript


Script MatrixLeft ()
if column > 1 then 
let column = column - 1
saystring(matrix.get_cell(row,column))
else
saystring("start row")
endif
EndScript


Script MatrixDown ()
if row < matrix.rows then
let row = row+1
saystring(matrix.get_cell(row,column))
else
saystring("end column")
endif
EndScript


Script MatrixUp ()
if row > 1 then
let row = row - 1
saystring(matrix.get_cell(row,column))
else
saystring("start columnd")
endif
EndScript



Script SayRow (int i)
if i>0 && i <= matrix.rows then 
saystring(matrix.get_row(i," "))
else 
saystring("Invalid row")
endif
EndScript

Script SayColumn (int i)
if i>0 && i <=matrix.columns then
saystring(matrix.get_col(i," "))
else
saystring("Invalid column")
endif
EndScript



Script preprocessorAdd ()
var string input, int args, string strargs, string translation
if InputBox ("Enter the command you wish to re-define.", "Initial LaTeX", input) then 
if InputBox ("Enter the number of arguments of the command.", "Number of arguments", strargs) then 
if InputBox ("Enter the definition of the custom command, that is, the standard LaTeX to which it is equivalent.", "Translation", translation) then 
let args=StringToInt (strargs)
latex_access.preprocessor_add(input,args,translation)
endif
endif
endif
EndScript


Script PreprocessorFromString ()
latex_access.preprocessor_from_string(GetSelectedText ())


EndScript



Script PreprocessorWrite ()
var string filename
if InputBox ("enter full filename to save to", "Filename", filename) then 
if FileExists (filename) then 
var int result
let result=ExMessageBox ("The file you specified already exists.  Do you wish to replace it?", "File Exists", MB_YESNO)
if result==IDNO then 
return 
endif
endif
latex_access.preprocessor_write(filename)
endif
EndScript


Script PreprocessorRead ()
var string filename
if InputBox ("enter full filename to read from ", "Filename", filename) then 
if FileExists (filename) then 
latex_access.preprocessor_read(filename)
else
MessageBox ("File does not exist")
endif
endif
EndScript








