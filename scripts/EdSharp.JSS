Include "HJConst.jsh"
Use "Homer.jsb"
; Use "Editor.jsb"

Globals
Int iEdSharpInitialized, Int ProcessMaths, Int row, Int column, Int iIndentMode,
Handle hPrevious,
Object oEnvUser, Object latex_access, object matrix,
String sEdSharpIniFile, String sEdSharpIndentModeFile, String sEdSharpTempFile, String sEdSharpConfigFile

Void Function AutoStartEvent()
Var
int i,
Object oShell,
String s

;SwitchToConfiguration("default")
If !iEdSharpInitialized Then
Let s = GetActiveConfiguration()
Let s = GetJAWSSettingsDirectory()
; Let sEdSharpIniFile = GetJAWSSettingsDirectory() + "\\" + s + ".ini"
; Let sEdSharpIniFile = GetJAWSSettingsDirectory()
Let i = StringContains(s, "\\Freedom Scientific\\")
Let sEdSharpIniFile = StringLeft(s, i) + "EdSharp\\EdSharp.ini"
Let sEdSharpIndentModeFile = StringLeft(s, i) + "EdSharp\\IndentMode.tmp"
Let sEdSharpTempFile = StringLeft(s, i) + s + "\\" + s + ".tmp"
Let sEdSharpConfigFile = StringLeft(sEdSharpIniFile, i) + s + "\\" + s + ".ini"
; Let sEdSharpConfigFile = StringLeft(sEdSharpIniFile, i) + s + "\\" + s + ".cfg"
Let oShell = CreateObject("WScript.Shell")
; Let oEnvUser = oShell.Environment("USER")
; Let iIndentMode = StringToInt(IniReadString("Options", "IndentMode", "0", sEdSharpIniFile))
Let iIndentMode = IniReadInteger("Data", "IndentMode", 0, sEdSharpIniFile)
; SayStringIf(iIndentMode, "On", "Off")
Let iEdSharpInitialized = True
; MessageBox(sEdsharpConfigFile)
; MessageBox(sEdSharpTempFile)

Let s = IniReadString("Options", "Punctuation", "", sEdSharpIniFile)
If !StringIsBlank(s) Then
SetJCFOption(OPT_PUNCTUATION,  StringToInt(s))
EndIf

Let s = IniReadString("Options", "Indentation", "", sEdSharpIniFile)
If !StringIsBlank(s) Then
SetJCFOption(opt_indicate_Indentation,  StringToInt(s))
EndIf

If ProcessMaths && !latex_access Then
let latex_access=CreateObject ("latex_access")
EndIf
SayString(GetWindowName(GetTopLevelWindow(GetFocus())))
EndIf
EndFunction

void Function WindowActivatedEvent(handle hWnd)
If GetAppFileName() != "EdSharp.exe" Then
Return
EndIf

If GetWindowClass(hWnd) == "#32770" Then
Return
EndIf

; SayString(GetWindowName(GetTopLevelWindow(GetFocus())))
SayString(GetWindowName(hwnd))
EndFunction

Void Function WindowMinMaxEvent (handle hWnd, int nMinMaxRest, int nShow)
If hWnd == hPrevious Then
Return
EndIf

Let hPrevious = hWnd
If GetAppFileName() != "EdSharp.exe" Then
Return
EndIf

; saystring("nMinMaxRest")
; SayInteger(nMinMaxRest)
; SayString("nShow")
; SayInteger(nShow)
If nMinMaxRest != 1 Then
; If nMinMaxRest > 1 Then
Return
EndIf

; SayString(GetWindowName(GetTopLevelWindow(GetFocus())))
SayString(GetWindowName(hwnd))
EndFunction

Void Function AutoFinishEvent()
Var
Object null

Let latex_access = Null
Let matrix = Null
EndFunction

Void Function SayCharacter()
; CopyToClipboard(StringLower(IniReadString("Options", "HardPageAddress", "x", sEdSharpConfigFile)))
; CopyToClipboard(sEdSharpConfigFile)
; If GetCharacterValue(GetCharacter()) == 0 && StringLower(IniReadString("Options", "HardPageAddress", "n", sEdSharpConfigFile)) == "y" Then
If GetCharacterValue(GetCharacter()) == 0 && StringLower(IniReadStringEx("Options", "HardPageAddress", "n", 0, sEdSharpConfigFile)) == "y" Then
Return
EndIf
Return SayCharacter()
EndFunction

int Function HandleCustomWindows(handle h)
;copytoclipboard(getclipboardtext() + "\n" + GetWindowClass(h))
If MenusActive() && GetWindowClass (h) == "WindowsForms10.Window." Then
SayFocusedObject()
Return True
EndIf
Return HandleCustomWindows(h)
EndFunction

Void Function SaveVoiceSetting(String sSetting, Int iLevel)
Var
Int iLoop,
String sJcf, String sVoice, String sVoiceList

Let sJcf = GetActiveConfiguration() + ".jcf"
Let sVoiceList = "Global|Error|Keyboard|Screen|PCCursor|JAWSCursor|Message"
Let iLoop = 1
While iLoop
Let sVoice = StringSegment(sVoiceList, "|", iLoop)
If StringIsBlank(sVoice) Then
Let iLoop = 0
Else
Let sVoice = "eloq-" + sVoice + "Context"
IniWriteInteger(sVoice, sSetting, iLevel, sJcf)
Let iLoop = iLoop + 1
EndIf
EndWhile
EndFunction

Int Function new_SayFileByLine(String sFile)
Var
Object oFile, Object oNull, Object oSystem,
String sLine

Let oSystem =CreateObjectEx("Scripting.FileSystemObject", False)
Let oFile =oSystem.OpenTextFile(sFile, 1, 0)
While !oFile.AtEndOfStream && !IsKeyWaiting()
Let sLine =oFile.ReadLine()
If !StringIsBlank(sLine) Then
SayString(sLine)
EndIf
EndWhile
oFile.Close()
Let oFile =oNull
Let oSystem =oNull
EndFunction

String Function FileToString(String sFile)
;Get content of text (not binary) file

Var
Int iRead,
Object oSystem, Object oFile, Object oNull,
String sReturn

Let oSystem =CreateObject("Scripting.FilesystemObject")
Let oFile =oSystem.OpenTextFile(sFile)
Let sReturn =oFile.ReadAll()
oFile.close()

Let oFile = oNull
Let oSystem = oNull
Return sReturn
EndFunction

Int Function StringToFile(String sText, String sFile)
;Saves string to text file, replacing if it exists

Var
Int iReturn, Int iReplace,
Object oSystem, Object oFile, Object oNull

Let oSystem =CreateObject("Scripting.FilesystemObject")
Let iReplace = True
Let oFile =oSystem.CreateTextFile(sFile, iReplace)
oFile.write(sText)
oFile.close()
Let iReturn = FileExists(sFile)
Let oFile = oNull
Let oSystem = oNull

Return iReturn
EndFunction

Int Function SayTempFile()
If FileExists(sEdSharpTempFile) Then
Var
String sText

new_SayFileByLine(sEdSharpTempFile)
;SayString(FileToString(sEdSharpTempFile))
Return True
Else
Return False
EndIf
EndFunction

Void Function old_SayAllStoppedEvent ()
Var
Int iLength

SpeechOff()
SelectFromTop()
Let iLength = StringLength(GetSelectedText())
UserBufferDeactivate()
SpeechOn()
StringToFile(IntToString(iLength), sEdSharpTempFile)
Pause()
TypeKey("Shift+F9");
EndFunction

String Function LaTeXToSpeech(String input)
let input = latex_access.speech(input)
let input = StringReplaceSubstrings (input, "&", "&amp;")
let input = StringReplaceSubstrings (input, "<sub>", smmGetStartMarkupForAttributes (attrib_subscript|attrib_text))
let input = StringReplaceSubstrings (input, "</sub>", smmGetEndMarkupForAttributes (attrib_subscript|attrib_text))
let input = StringReplaceSubstrings (input, "<bold>", smmGetStartMarkupForAttributes (attrib_bold|attrib_text))
let input = StringReplaceSubstrings (input, "</bold>", smmGetEndMarkupForAttributes (attrib_bold|attrib_text))
Return input
EndFunction

String Function SayAllTempFile()
Var
Int i, Int iCount,
String s, String sText

If FileExists(sEdSharpTempFile) Then
UserBufferDeactivate()
UserBufferClear()
Let sText = StringLeft(FileToString(sEdSharpTempFile), 64000)
If ProcessMaths Then
/*
Let s = sText
Let sText = ""
Let i = 1
Let iCount = StringSegmentCount(s, "\n")
While i <= iCount
Let sText = sText + StringSegment(LaTeXToSpeech(s), "\n", i)
Let i = i + 1
EndWhile
*/
Let sText = LaTeXToSpeech(sText)
EndIf

UserBufferAddText(sText)
UserBufferActivate()
JAWSTopOfFile()
SayAll()
Return True
Else
Return False
EndIf
EndFunction

Script old_SayAll()
TypeKey("F9")
EndScript

Int Function UIIsEditorWindow()
Var
String sKey, String sClass
Let sClass = GetWindowClass(GetFocus())
Let sKey = GetCurrentScriptKeyName()

return !IsVirtualPCCursor() && IsPCCursor() && sClass == "WindowsForms10.RichEdit" || sClass == "WindowsForms10.MDICLIENT."
EndFunction

Int Function UIIsListWindow()
Var
String sKey, String sClass
Let sClass = StringLower(GetWindowClass(GetFocus()))
Let sKey = GetCurrentScriptKeyName()
return !IsVirtualPCCursor() && IsPCCursor() && StringContains(sClass, ".listbox.")
EndFunction

Script BottomOfFile()
if UIIsListWindow() Then
TypeCurrentScriptKey()
ElIf UIIsEditorWindow() Then
JAWSBottomOfFile()
SayLine()
Else
PerformScript BottomOfFile()
EndIf
EndScript

Script CloseDocumentWindow()
If UIIsEditorWindow() Then
TypeCurrentScriptKey()
Else
PerformScript CloseDocumentWindow()

EndIf
EndScript

Script CopySelectedTextToClipboard()
if UIIsEditorWindow() Then
TypeCurrentScriptKey()
Else
PerformScript CopySelectedTextToClipboard()
EndIf
EndScript

Script CutToClipboard()
if UIIsEditorWindow() Then
TypeCurrentScriptKey()
Else
PerformScript CutToClipboard()
EndIf
EndScript

Script DeleteWord()
;Delete word and say new one

If UIIsEditorWindow() Then
SpeechOff()
SelectNextWord()
TypeKey("Delete")
RefreshWindow(GetFocus())
Pause()
SpeechOn()
SayWord()
Else
PerformScript UISayAndTypeCurrentScriptKey()
EndIf
EndScript

Script DeleteWordBack()
;Delete word back and say new one
Var
String s

If UIIsEditorWindow() Then
SpeechOff()
PriorWord()
Let s = GetWord()
SelectNextWord()
TypeKey("Delete")
RefreshWindow(GetFocus())
Pause()
SpeechOn()
SayString(s)
Else
PerformScript ControlBackspace()
EndIf
EndScript

Script JAWSEnd()
if UIIsEditorWindow() Then
JAWSEnd()
SayCharacter()
Else
PerformScript JAWSEnd()
EndIf
EndScript

Script JAWSHome()
if UIIsEditorWindow() Then
JAWSHome()
SayCharacter()
Else
PerformScript JAWSHome()
EndIf
EndScript

Script JAWSPageDown()
if UIIsEditorWindow() Then
JAWSPageDown()
SayLine()
Else
PerformScript JAWSPageDown()
EndIf
EndScript

Script JAWSPageUp()
if UIIsEditorWindow() Then
JAWSPageUp()
SayLine()
Else
PerformScript JAWSPageUp()
EndIf
EndScript

Script MouseDown()
if UIIsListWindow() Then
TypeCurrentScriptKey()
Else
PerformScript MouseDown()
EndIf
EndScript

Script MouseUp()
if UIIsListWindow() Then
TypeCurrentScriptKey()
Else
PerformScript MouseUp()
EndIf
EndScript

Script PasteFromClipboard()
if UIIsEditorWindow() Then
TypeCurrentScriptKey()
Else
PerformScript PasteFromClipboard()
EndIf
EndScript

Script SayActiveCursor()
if UIIsEditorWindow() Then
TypeKey("Alt+A")
Else
PerformScript SayActiveCursor()
EndIf
EndScript

Script SayPriorLine()
Var
String sValue

; If IniReadInteger("Data", "IndentMode", 0, sEdSharpIniFile) && IsPCCursor() Then
If FileExists(sEdSharpIndentModeFile) && IsPCCursor() Then
PriorLine()
Delay(1)
; Let sValue = oEnvUser.Item("EdSharpIndent")
Let sValue = IniReadString("Data", "IndentChange", "", sEdSharpIndentModeFile)
If !StringIsBlank(sValue) Then SayString(sValue) EndIf
builtin::SayLine()
Else
PerformScript SayPriorLine()
EndIf
EndScript

Script SayNextLine()
Var
String sValue

; If IniReadInteger("Data", "IndentMode", 0, sEdSharpIniFile) && IsPCCursor() Then
If FileExists(sEdSharpIndentModeFile) && IsPCCursor() Then
; SayString("exists")
NextLine()
Delay(1)
; Let sValue = oEnvUser.Item("EdSharpIndent")
Let sValue = IniReadString("Data", "IndentChange", "", sEdSharpIndentModeFile)
If !StringIsBlank(sValue) Then SayString(sValue) EndIf
builtin::SayLine()
Else
PerformScript SayNextLine()
EndIf
EndScript

Script SayLine()
Var
Handle h

Let h = GetCurrentWindow()
If MenusActive() && GetWindowClass (h) == "WindowsForms10.Window." Then
SayFocusedObject()
Return
EndIf
PerformScript SayLine()
EndScript

Script SaySelectedText()
if UIIsEditorWindow() Then
TypeKey("Shift+Space")
Else
PerformScript SaySelectedText()
EndIf
EndScript

Script ScriptFileName ()
ScriptAndAppNames ("EdSharp")
EndScript

Script SpeedFaster ()
; Control+Accent = make voice faster
Var
Int iLevel, Int iMax, Int iMin,
String sSetting

SayString("Speed faster")
Let sSetting = "Speed"
Let iLevel =GetVoiceRate(VCTX_GLOBAL , True)
GetSynthRateRange(iMin, iMax)
If iLevel == iMax Then
SayString("Top")
Else
Let iLevel = iLevel +(5 *(iMax -iMin) /100)
Let iLevel =Min(iLevel, iMax)
SetVoiceRate(VCTX_GLOBAL , iLevel)
;GetVoiceRate(VCTX_GLOBAL, True)
SaveVoiceSetting(sSetting, iLevel)
SayString(IntToString(100 *(iLevel -iMin)/(iMax -iMin)) +" percent")
EndIf
EndScript

Script SpeedSlower ()
; Control+Shift+Accent = make voice slower
Var
Int iLevel, Int iMax, Int iMin,
String sSetting

SayString("Speed slower")
Let sSetting = "Speed"
Let iLevel =GetVoiceRate(VCTX_GLOBAL , True)
GetSynthRateRange(iMin, iMax)
If iLevel == iMin Then
SayString("Bottom")
Else
Let iLevel =iLevel -(5 * (iMax -iMin) /100)
Let iLevel =max(iLevel, iMin)
SetVoiceRate(VCTX_GLOBAL , iLevel)
;GetVoiceRate(VCTX_GLOBAL, True)
SaveVoiceSetting(sSetting, iLevel)
SayString(IntToString(100 *(iLevel -iMin)/(iMax -iMin)) +" percent")
EndIf
EndScript

Script ToggleIndentation()
; Alt+Control+Shift+I = toggle indent announcement
Var
Int iLevel,
String sSetting

Let sSetting = "Indentation"
Let iLevel = GetJCFOption(OPT_Indicate_Indentation)
If iLevel != 0 Then
SayString("No Announcement")
Let iLevel = 0
Else
SayString("Announce Indentation")
Let iLevel = 1
EndIf
SetJCFOption(opt_indicate_Indentation, iLevel)
IniWriteInteger("Options", "Indentation", iLevel, sEdSharpIniFile)
EndScript

Script TogglePunctuation()
; Insert+Accent = Toggle between all and no punctuation
Var
Int iLevel,
String sSetting

Let sSetting = "Punctuation"
Let iLevel = GetJCFOption(OPT_PUNCTUATION)
If iLevel != 0 Then
SayString("No Punctuation")
Let iLevel = 0
Else
SayString("All punctuation")
Let iLevel = 3
EndIf
SetJCFOption(OPT_PUNCTUATION, iLevel)
;SaveVoiceSetting(sSetting, iLevel)
IniWriteInteger("Options", "Punctuation", iLevel, sEdSharpIniFile)
EndScript

Script TopOfFile()
If UIIsListWindow() Then
TypeCurrentScriptKey()
ElIf UIIsEditorWindow() Then
JAWSTopOfFile()
SayLine()
Else
PerformScript TopOfFile()
EndIf
EndScript

Script TypeCurrentScriptKey()
Var
String sKey, String sClass
Let sClass = StringLower(GetWindowClass(GetFocus()))
Let sKey = GetCurrentScriptKeyName()
if UIIsEditorWindow() || StringContains(sClass, ".listbox.") Then
Else
SayCurrentScriptKeyLabel()
EndIf
TypeCurrentScriptKey()
EndScript

Script IndentMode()
if 0 then
SayString("test")
Let iIndentMode = !iIndentMode
SayStringIf(iIndentMode, "On", "Off")
IniWriteInteger("Options", "IndentMode", iIndentMode, sEdSharpIniFile)
endIf
TypeCurrentScriptKey()
EndScript

Script Undo()
Var
String sKey, String sClass
Let sClass = GetWindowClass(GetFocus())
Let sKey = GetCurrentScriptKeyName()

if UIIsEditorWindow() Then
TypeCurrentScriptKey()
Else
PerformScript Undo()
EndIf
EndScript

Script VolumeLouder()
; Alt+Accent = make voice louder
Var
Int iLevel, Int iMax, Int iMin,
String sSetting

SayString("Volume louder")
Let sSetting = "Volume"
Let iLevel =GetVoiceVolume(VCTX_GLOBAL , True)
GetVoiceVolumeRange(iMin, iMax)
Let iLevel =GetSystemVolume()
GetSystemVolumeRange(iMin, iMax)
If iLevel ==iMax Then
SayString("Top")
Else
Let iLevel =iLevel +(5 *(iMax -iMin) /100)
Let iLevel =Min(iLevel, iMax)
SetVoiceVolume(VCTX_GLOBAL , iLevel)
SetSystemVolume(iLevel)
SaveVoiceSetting(sSetting, iLevel)
SayString(IntToString(100 *(iLevel -iMin)/(iMax -iMin)) +" percent")
EndIf
EndScript

Script VolumeSofter()
; Alt+Shift+Accent = make voice softer
Var
Int iLevel, Int iMax, Int iMin,
String sSetting

SayString("Volume softer")
Let sSetting = "Volume"
Let iLevel =GetVoiceVolume(VCTX_GLOBAL , True)
GetVoiceVolumeRange(iMin, iMax)
Let iLevel =GetSystemVolume()
GetSystemVolumeRange(iMin, iMax)
If iLevel ==iMin Then
SayString("Bottom")
Else
Let iLevel =iLevel -(5 * (iMax -iMin) /100)
Let iLevel =max(iLevel, iMin)
SetVoiceVolume(VCTX_GLOBAL , iLevel)
SetSystemVolume(iLevel)
SaveVoiceSetting(sSetting, iLevel)
SayString(IntToString(100 *(iLevel -iMin)/(iMax -iMin)) +" percent")
EndIf
EndScript

Script InsertAllUsersPath()
Var
String sPath

If GetControlID(GetFocus()) == 1148 Then
SayString("Insert all users path")
Let sPath = GetJAWSSettingsDirectory()
;Let sPath = StringReplaceEx(sPath, GetJAWSUserName(), "All Users", False) + "\\"
Let sPath = StringReplaceSubstrings(StringLower(sPath), StringLower(GetJAWSUserName()), "All Users") + "\\"
Pause()
TypeString(sPath)
Else
PerformScript TypeCurrentScriptKey()
EndIf
EndScript

Script InsertScriptPath()
If GetControlID(GetFocus()) == 1148 Then
SayString("Insert script path")
Pause()
TypeString(GetJAWSSettingsDirectory() + "\\")
Else
PerformScript TypeCurrentScriptKey()
EndIf
EndScript

Script SelectNextLine()
If UIIsListWindow() Then
TypeCurrentScriptKey()
Else
PerformScript SelectNextLine()
EndIf
EndScript

Script SelectPriorLine()
If UIIsListWindow() Then
TypeCurrentScriptKey()
Else
PerformScript SelectPriorLine()
EndIf
EndScript




Script SelectToEndOfLine()
If UIIsListWindow() Then
TypeCurrentScriptKey()
Else
PerformScript SelectToEndOfLine()
EndIf
EndScript

Script SelectFromStartOfLine()
If UIIsListWindow() Then
TypeCurrentScriptKey()
Else
PerformScript SelectFromStartOfLine()
EndIf
EndScript




Script SelectToBottom()
If UIIsListWindow() Then
TypeCurrentScriptKey()
Else
PerformScript SelectToBottom()
EndIf
EndScript

Script SelectFromTop()
If UIIsListWindow() Then
TypeCurrentScriptKey()
Else
PerformScript SelectFromTop()
EndIf
EndScript




Script ControlDownArrow()
If UIIsEditorWindow() || UIIsListWindow() Then
TypeCurrentScriptKey()
Else
PerformScript ControlDownArrow()
EndIf
EndScript

Script ControlUpArrow()
If UIIsEditorWindow() || UIIsListWindow() Then
TypeCurrentScriptKey()
Else
PerformScript ControlUpArrow()
EndIf
EndScript


Script ProcessLaTeX()
Var
String sText

If latex_access && IsSameScript() Then
SayString("Examine")
Let sText = GetLine()
Let sText = LaTeXToSpeech(sText)
Say(sText, ot_selected_item, true)
UserBufferClear()
UserBufferAddText(sText)
UserBufferActivate()
JAWSTopOfFile()
Return
EndIf

SayString("Process LaTeX")
If !latex_access Then
let latex_access=CreateObject ("latex_access")
If !latex_access Then
SayString("Error creating latex_access object!")
Let ProcessMaths = False
SwitchToScheme("Default")
return
EndIf

EndIf

Let ProcessMaths = True
SwitchToScheme("latex")
EndScript

Script NoProcessing()
SayString("No processing")
Let ProcessMaths=False
SwitchToScheme("Default")
EndScript

String Function StringPlural(String sItem, Int iCount)
;Return singular or plural form of a string, depending on whether count equals one
Var
String sReturn

Let sReturn = IntToString(iCount) + " " + sItem
If iCount != 1 Then
Let sReturn = sReturn + "s"
EndIf
Return sReturn
EndFunction

Script InputMatrix ()
SayString("Input matrix")
If !matrix Then
let matrix=CreateObject ("latex_access_matrix")
EndIf

If !matrix Then
SayString("Error creating matrix object!")
Return
EndIf

let row=1
let column=1
matrix.tex_init(GetSelectedText ())
var string msg
SayString(StringPlural("row", matrix.rows))
SayString(StringPlural("column", matrix.columns))
EndScript


Void Function SayLine ()
var string input
let input = GetLine ()
if  ProcessMaths then
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
;SayMessage (ot_status, "maths to be read as plain latex", "Processing off")
else
let ProcessMaths = true
;SayMessage(OT_STATUS,"Maths to be processed to a more verbal form","Processing on")
endif


EndScript



Int Function BrailleBuildLine ()
if  ProcessMaths then
var string input
let input = GetLine()
let input = latex_access.nemeth(input)
; BrailleAddString (SubString (input, 1, 1), 0, 0, ATTRIB_HIGHLIGHT)
; BrailleAddString (SubString (input, 2, StringLength (input)), 0, 0, 0)
BrailleAddString (input, 0, 0, 0)

endif
return true
EndFunction

Script MatrixRight ()
if column < matrix.columns then
let column = column+1
SayString("c" + IntToString(column))
saystring(matrix.get_cell(row,column))
else
saystring("End!")
endif
EndScript


Script MatrixLeft ()
if column > 1 then 
let column = column - 1
SayString("c" + IntToString(column))
saystring(matrix.get_cell(row,column))
else
saystring("Start!")
endif
EndScript


Script MatrixDown ()
if row < matrix.rows then
let row = row+1
SayString("r" + IntToString(row))
saystring(matrix.get_cell(row,column))
else
saystring("Bottom!")
endif
EndScript


Script MatrixUp ()
if row > 1 then
let row = row - 1
SayString("r" + IntToString(row))
saystring(matrix.get_cell(row,column))
else
saystring("Top!")
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
var string input, string translation
InputBox ("Enter the custom LaTeX you wish to re-define.", "Initial LaTeX", input)
InputBox ("Enter the definition of the custom command, that is, the standard LaTeX to which it is equivalent.", "Translation", translation)
latex_access.preprocessor_add(input,translation)
EndScript

Script PreprocessorCsv ()
var string filename
InputBox ("Enter the name of the CSV file you wish to load", "Filename", filename)
if FileExists (filename) then 
latex_access.load_csv(filename)
else 
saystring("File not found")
endif
EndScript

Script MatrixTopLeft()
If ProcessMaths Then
Let row = 1
Let column = 1
SayString("r" + IntToString(row) + " c" + IntToString(column))
saystring(matrix.get_cell(row, column))

Else
PerformScript FirstCellInTable()
EndIf
EndScript

Script MatrixBottomRight()
If ProcessMaths Then
Let row = matrix.rows
Let column = matrix.columns
SayString("r" + IntToString(row) + " c" + IntToString(column))
saystring(matrix.get_cell(row, column))

Else
PerformScript LastCellInTable()
EndIf
EndScript

Script SayCurrentCell()
If ProcessMaths Then
If IsSameScript() Then
SayString("r" + IntToString(row) + " c" + IntToString(column))
spellString(matrix.get_cell(row, column))
Else
saystring(matrix.get_cell(row, column))
EndIf

Else
PerformScript SayCell()
EndIf
EndScript

Script SayCurrentColumn()
If ProcessMaths Then
If IsSameScript() Then
SayString("c" + IntToString(column))
spellstring(matrix.get_col(column, " "))
Else
saystring(matrix.get_col(column, " "))
EndIf

Else
PerformScript DecreaseVoiceRate()
EndIf
EndScript

Script SayCurrentRow()
If ProcessMaths Then
If IsSameScript() Then
SayString("r" + IntToString(row))
Spellstring(matrix.get_row(row, " "))
Else
saystring(matrix.get_row(row, " "))
EndIf

Else
PerformScript IncreaseVoiceRate()
EndIf
EndScript

Script SelectAll()
Var
String sKey, String sClass
Let sClass = GetWindowClass(GetFocus())
Let sKey = GetCurrentScriptKeyName()

if UIIsEditorWindow() Then
TypeCurrentScriptKey()
Else
PerformScript SelectAll()
EndIf
EndScript

Script SilentKey()
TypeCurrentScriptKey()
EndScript

String Function StringTrimWhiteSpace(String sText)
; Trim leading and trailing white space characters

Var
String sReturn

Let sReturn = StringTrim(sText)
Let sReturn = RegExpReplaceCase(sReturn, "^\\s+", "")
Let sReturn = RegExpReplaceCase(sReturn, "\\s+$", "")
Return sReturn
EndFunction

String Function PathGetBase(String sFile)
;Get base/root name

Var
Object oSystem, Object oNull,
String sReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let sReturn =oSystem.GetBaseName(sFile)

Let oSystem = oNull
Return sReturn
EndFunction

String Function DialogPickWithIndex(String sTitle, String sValues, Int bSort, Int iIndex)
;Get choice from a standard list box
Var
Handle h,
Int iChoice,
String sReturn

If StringIsBlank(sTitle) Then
Let sTitle = "Pick"
EndIf
If bSort Then
Let sValues = StringSegmentSort(sValues, "\7")
EndIf
Let iChoice = DlgSelectItemInList(sValues, sTitle, False, iIndex)
If iChoice Then
Let sReturn = StringSegment(sValues, "\7", iChoice)
EndIf
Return sReturn
EndFunction

String Function ConvertToUnixLineBreak(String sText)
;Convert to Unix line break, \n
Var
String sMatch, String sReplace

Let sMatch = "\r\n"
Let sReplace = "\n"
Let sText = RegExpReplaceCase(sText, sMatch, sReplace)
Let sMatch = "\r"
Let sText = RegExpReplaceCase(sText, sMatch, sReplace)
Return sText
EndFunction

Script UIWebClientUtilities()
Var
Int bSort,
Int i, Int iCount, Int iIndex,
String sWebClientFile, String sBody, String sCommand, String sExe, String sDir, String sFiles, String sFile, String sNames, String sName, String sValues, String sValue, String sBase, String sTitle, String sInputFile, String sOutputFile, String sCodeFile

Let sDir = GetJAWSSettingsDirectory()
Let sFiles = PathGetDir(sDir, "WebClient_*.py", "")
Let sFiles = ConvertToUnixLineBreak(sFiles)
Let i = 1
Let iCount = StringSegmentCount(sFiles, "\n")
While i <= iCount
Let sFile = StringSegment(sFiles, "\n", i)
Let sName = PathGetName(sFile)
Let sBase = PathGetBase(sName)
Let sBase = StringChopLeft(sBase, StringLength("WebClient_"))
Let sNames = sNames + sBase + "\7"
Let sValue = PathCombine(sDir, sName)
Let sValues = sValues + sValue + "\7"
Let i = i + 1
EndWhile
Let sNames = StringChopRight(sNames, 1)
Let sValues = StringChopRight(sValues, 1)

Let sBase = IniReadSetting("WebClientUtilities", "")
If StringLength(sBase) Then 
Let iIndex = StringSegmentIndex(sNames, "\7", sBase)
EndIf
If iIndex == 0 Then
Let iIndex = 1
EndIf

Let sTitle = "Web Client Utilities"
Let bSort = False
Let sName = DialogPickWithIndex(sTitle, sNames, bSort, iIndex)
If Not sName Then
Return
EndIf

IniWriteSetting("WebClientUtilities", sName)
Let iIndex = StringSegmentIndex(sNames, "\7", sName)
Let sFile = StringSegment(sValues, "\7", iIndex)

Let sExe = PathCombine(sDir, "InPy.exe")
Let sExe = PathGetShort(sExe)
Let sInputFile = PathCombine(sDir, GetActiveConfiguration() + ".ini")
Let sBase = PathGetBase(sFile)
Let sOutputFile = PathCombine(sDir, sBase + ".txt")

if 1 then
Let sCodeFile = sFile
Else
Let sWebClientFile = PathCombine(sDir, "WebClient.py")
Let sBody = StringTrimWhiteSpace(FileToString(sWebClientFile)) + "\r\n\r\n" + StringTrimWhiteSpace(FileToString(sFile)) + "\r\n"
Let sCodeFile = PathGetTempFile()
StringToFile(sBody, sCodeFile)
EndIf

Let sCommand = sExe + " " + StringQuote(sCodeFile) + " " + StringQuote(sInputFile) + " " + StringQuote(sOutputFile)
FileDelete(sOutputFile)
ShellRun(sCommand, 1, True)
; FileDelete(sCodeFile)
If FileExists(sOutputFile) Then
ShellRun(StringQuote(sOutputFile), 1, False)
EndIf
EndScript
