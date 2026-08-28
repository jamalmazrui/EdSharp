;Homer Script Library
;Copyright 2001 - 2008 by Jamal Mazrui
;Modified GPL License

Include "hjglobal.jsh"
Include "hjconst.jsh"
Include "msaa.jsh"
Include "Homer.jsh"

Int Function AppActivate(String sTitle)
;Activate a window by its title or a substring within it

Var
Int iReturn,
Object oShell, Object oNull

Let oShell =ObjectCreate("Wscript.Shell")
Let iReturn =oShell.AppActivate(sTitle)

Let oShell = oNull
Return iReturn
EndFunction

Int Function AppActivateHandle(Handle h)
; Activate a window by its handle using ForceWin.exe

Var
Int iReturn,
String sExe, String sCommand

Let sExe = PathGetHomer() + "ForceWin.exe"
Let sCommand = StringQuote(sExe) + " " + IntToString(h)
ShellRun(sCommand, 0, True)
Return iReturn
EndFunction

Int Function AppDestroy(Handle h)
;Try to force an application to close
Var
Int iMessage, Int iParam1, Int iParam2

If !h Then
Let h = GetAppMainWindow(GetFocus())
EndIf

Let iMessage = 0x2 ;WM_DESTROY
Return SendMessage(h, iMessage, iParam1, iParam2)
EndFunction

String Function AppGetTitle(Handle h)
;Replacement for GetAppTitle(), which truncates long titles that extend beyond the visible window

If !h Then
Let h = GetAppMainWindow(GetFocus())
EndIf

Return GetWindowName(h)
EndFunction

Int Function AppQuit(Handle h)
;Ask an application to exit
Var
Int iMessage, Int iParam1, Int iParam2

If !h Then
Let h = GetAppMainWindow(GetFocus())
EndIf

Let iMessage = 0x12 ;WM_QUIT
Return SendMessage(h, iMessage, iParam1, iParam2)
EndFunction

Int Function AppSetActive(Handle h, Int iState)
;Activate or inactivate an application window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetAppMainWindow(GetFocus())
EndIf

Let iMessage = 0x1C ;WM_ACTIVATEAPP
;Istate options
;WA_INACTIVE    = 0
;WA_ACTIVE      = 1
;WA_CLICKACTIVE = 2
Return SendMessage(h, iMessage, iState, 0)
EndFunction

String Function BooDialogMultiCheck(String sTitle, String sItems, String sSelected, Int iSort)
;Check items in a list using IronCOM.dll
Var
Handle h,
String sCode, String sBoo, String sReturn

Let h = GetFocus()
If stringIsBlank(sTitle) Then
Let sTitle = "Multi Check"
EndIf

Let sBoo = GetJAWSSettingsDirectory() + "\\Homer\\DialogMultiCheck.boo"
Let sCode = FileToString(sBoo)
;ScheduleFunction("BooDialogHelper", 0)
Let sReturn = BooEval(sCode, sTitle, sItems, sSelected, IntToString(iSort))

If h Then
SetFocus(h)
EndIf

return sReturn
EndFunction

String Function BooDialogChoose(String sTitle, String sButtons, Int iDefault)
;Choose from a set of buttons using IronCOM.dll
Var
Handle h,
String sCode, String sBoo, String sReturn

If !BooInit() Then
Return IniFormDialogChoose(sTitle, sButtons, iDefault)
EndIf

Let h = GetFocus()
If stringIsBlank(sTitle) Then
Let sTitle = "Choose"
EndIf

Let sBoo = GetJAWSSettingsDirectory() + "\\Homer\\DialogChoose.boo"
Let sCode = FileToString(sBoo)
;ScheduleFunction("BooDialogHelper", 0)
Let sReturn = BooEval(sCode, sTitle, sButtons, IntToString(iDefault), "")

If h Then
SetFocus(h)
EndIf

return sReturn
EndFunction

String Function BooDialogMultiInput(String sTitle, string sFields, String sValues)
;Get input from multiple edit boxes using IronCOM.dll
Var
Handle h,
Int iCount,
String sCode, String sBoo, String sReturn

If !BooInit() Then
Return IniFormDialogMultiInput(sTitle, sFields, sValues)
EndIf

Let h = GetFocus()
Let iCount = StringCountSegment(sFields, "\7")
If stringIsBlank(sTitle) Then
If iCount == 1 Then
Let sTitle = "Input"
Else
Let sTitle = "Fields"
EndIf
EndIf

If StringIsBlank(sFields) && iCount == 1 Then
Let sFields = "Text"
EndIf

Let sBoo = GetJAWSSettingsDirectory() + "\\Homer\\DialogMultiInput.boo"
Let sCode = FileToString(sBoo)
;ScheduleFunction("BooDialogHelper", 0)
Let sReturn = BooEval(sCode, sTitle, sFields, sValues, "")

If h Then
SetFocus(h)
EndIf

return sReturn
EndFunction

String Function BooEval(String sCode, String s1, String s2, String s3, String s4)
;Evaluate with BooEval object, passing four string parameters and returning a string result
Var
Int iCreate,
Object oNull,
String sReturn

If !BooInit() Then
Return
EndIf

Let sReturn = oHomerBoo.Eval(sCode, s1, s2, s3, s4)
Return sReturn
EndFunction

String Function BooEvalFile(String sFile, String s1, String s2, String s3, String s4)
;Evaluate file with BooEval object, passing four string parameters and returning a string result
Var
String sCode, String sReturn

If !StringContains(sFile, "\\") Then
Let sFile = GetJAWSSettingsDirectory() + "\\Homer\\" + sFile
EndIf

Let sCode = FileToString(sFile)
Let sReturn = BooEval(sCode, s1, s2, s3, s4)
Return sReturn
EndFunction

Int Function BooInit()
;Try to initiate global Boo object if it does not exist, and then test whether it exists

If iBooInitialized == -1 Then
Return False
ElIf oHomerBoo Then
Return True
EndIf

Let oHomerBoo = CreateObjectEx("Iron.COM", False)
If oHomerBoo Then
Let iBooInitialized = 1
Return True
Else
SayString("Error initializing Boo component!")
Let iBooInitialized = -1
Return False
EndIf
EndFunction

String Function BooPathGetShort(String sPath)
;get short path of a file or folder using IronCOM.dll
Var
String sCode, String sBoo, String sReturn

If !BooInit() Then
Return sPath
EndIf

Let sBoo = GetJAWSSettingsDirectory() + "\\Homer\\PathGetShort.boo"
Let sCode = FileToString(sBoo)
Let sReturn = BooEval(sCode, sPath, "", "", "")
return sReturn
EndFunction

String Function BooShellGetSoundDevices()
;Get list of sound devices using IronCOM.dll
Var
String sCode, String sBoo, String sReturn

If !BooInit() Then
Return ""
EndIf

Let sBoo = GetJAWSSettingsDirectory() + "\\Homer\\ShellGetSoundDevices.boo"
Let sCode = FileToString(sBoo)
Let sReturn = BooEval(sCode, "", "", "", "")
Let sReturn = StringChopRight(sReturn, 1)
return sReturn
EndFunction

Int Function BooShellUrlToFile(String sUrl, String sFile)
;Download file using IronCOM.dll
Var
Int iReturn,
String sCode

If !FileDelete(sFile) Then
Return
EndIf

Let sCode = "import Microsoft.VisualBasic.Devices from Microsoft.VisualBasic.dll\n"
Let sCode = sCode + "oNetwork = Network()\n"
Let sCode = sCode + "oNetwork.DownloadFile(s1, s2)\n"
BooEval(sCode, sUrl, sFile, "", "")

Let iReturn = FileExists(sFile)
Return iReturn
EndFunction

Int Function ButtonClick(Handle h)
;Activate a button

Var
Int BM_CLICK

If !h Then
Let h = GetFocus()
EndIf

Let BM_CLICK=245
Return SendMessage(h, BM_CLICK, 0, 0)
EndFunction

Int Function ButtonGet3State(Handle h)
;Test for 3 State check box style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x6 ;BS_AUTO3STATE
Return (iStyleBits & iStyle)
EndFunction

Int Function ButtonGetCheck(Handle h)
;Test for check box style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x3 ;BS_AUTOCHECKBOX
Return (iStyleBits & iStyle)
EndFunction

Int Function ButtonGetChecked(Handle h)
;Test whether a checkbox or radio button is checked
Var
Int BM_GETCHECK

If !h Then
Let h = GetFocus()
EndIf

Let BM_GETCHECK =240
Return SendMessage(h, BM_GETCHECK, 0, 0)
EndFunction

Int Function ButtonGetDefault(Handle h)
;Test for default button style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x1 ;BS_DEFPUSHBUTTON
Return (iStyleBits & iStyle)
EndFunction

Int Function ButtonGetGroup(Handle h)
;Test for group box style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x7 ;BS_GROUPBOX
Return (iStyleBits & iStyle)
EndFunction

Int Function ButtonGetPush(Handle h)
;Test for push button style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x0 ;BS_PUSHBUTTON
Return (iStyleBits & iStyle)
EndFunction

Int Function ButtonGetRadio(Handle h)
;Test for radio button style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x9 ;BS_AUTORADIOBUTTON
Return (iStyleBits & iStyle)
EndFunction

Int Function ButtonInvokeID(Handle hDlg, Int iID)
;Activate a button by its ID
If !hDlg Then
Let hDlg = GetRealWindow(GetFocus())
EndIf
Return ButtonClick(FindDescendantWindow(hDlg, iID))
EndFunction

Int Function ButtonSetChecked(Handle h, Int iState)
;Set checked state of a checkbox or radio button
Var
Int iMessage, Int iCurrent

If !h Then
Let h = GetFocus()
EndIf

Let iCurrent = ButtonIsChecked(h)
If (iState && !iCurrent) || (!iState && iCurrent) Then
;Unreliable
;Let iMessage = 241 ; BM_SETCHECK
;Return SendMessage(h, iMessage, iState, 0)
ButtonClick(h)
EndIf
EndFunction

Int Function ControlClear(Handle h)
;Tell a control to clear/delete
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x303 ;WM_CLEAR
Return SendMessage(h, iMessage, 0, 0)
EndFunction

Int Function ControlCopy(Handle h)
;Tell a control to copy to the clipboard
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x301 ;WM_COPY
Return SendMessage(h, iMessage, 0, 0)
EndFunction

Int Function ControlCut(Handle h)
;Tell a control to cut to the clipboard
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x300 ;WM_CUT
Return SendMessage(h, iMessage, 0, 0)
EndFunction

Int Function ControlGetTextLength(Handle h)
;Get the length of text in a control
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xE
Return SendMessage(h, iMessage, 0, 0)
EndFunction

Int Function ControlIsEdit(Handle h)
;Test whether a control is the standard Edit class
If !h Then
Let h = GetFocus()
EndIf

Return StringEqual(GetWindowClass(h), "Edit")
EndFunction

Int Function ControlIsIEDoc()
;Test whether control class is Internet Explorer document
If IEGetCurrentDocument() Then
Return True
Else
Return False
EndIf
EndFunction

Int Function ControlIsIEServer(Handle h)
;Test whether control class is Internet Explorer Server
Var
String sClass

Let h = GetFocus()
Let sClass = GetWindowClass(h)
Return StringEqual(sClass, "Internet Explorer_Server")
EndFunction

Int Function ControlIsMDIFrame(Handle h)
;Test whether control class is MDI frame (no child window open)
Var
String sClass

Let h = GetFocus()
Let sClass = GetWindowClass(h)
Return StringEqual(sClass, "MDIClient")
EndFunction

Int Function ControlIsRichEdit(Handle h)
;Test whether control class is RichEdit 
Var
String sClass

Let h = GetFocus()
Let sClass = StringLower(GetWindowClass(h))
Return StringContains(sClass, "richedit") || StringContains(sClass, "tclipdoc")
EndFunction

Int Function ControlIsRichEditDoc()
;Test whether control class is RichEdit document
If GetRichEditDocument() Then
Return True
Else
Return False
EndIf
EndFunction

Int Function ControlNext(Handle h)
;Go to the next control in the tab order
Var
Int iMessage, Int iParam1, Int iParam2

If !h Then
Let h = GetRealWindow(GetFocus())
EndIf

Let iMessage = 0x28 ;WM_NEXTDLGCTL
Let iParam1 = False
Let iParam2 = True
Return SendMessage(h, iMessage, iParam1, iParam2)
EndFunction

Int Function ControlPaste(Handle h)
;Tell a control to paste from the clipboard
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x302 ;WM_PASTE
Return SendMessage(h, iMessage, 0, 0)
EndFunction

Int Function ControlPrevious(Handle h)
;Go to the previous control in the tab order
Var
Int iMessage, Int iParam1, Int iParam2

If !h Then
Let h = GetRealWindow(GetFocus())
EndIf

Let iMessage = 0x28 ;WM_NEXTDLGCTL
Let iParam1 = False
Let iParam2 = False
Return SendMessage(h, iMessage, iParam1, iParam2)
EndFunction

Int Function ControlPrint(Handle h)
;Tell a control to print
Var
Int iMessage, Int iParam1, Int iParam2

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x317 ;WM_PRINT
Return SendMessage(h, iMessage, iParam1, iParam2)
EndFunction

Int Function ControlUndo(Handle h)
;Tell a control to undo the last editing operation
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x304 ;WM_UNDO
Return SendMessage(h, iMessage, 0, 0)
EndFunction

Int Function ConvertExcelToText(String sSource, String sTarget)
;Convert .xls to .txt file using Microsoft Excel
Var
Handle h, Handle hApp,
Int i, Int iAlerts, Int iGet, Int iSheet,
Int iSheetCount, Int iUpdating, Int iVisible,
Object oNull, Object oApp, Object oSheet,
Object oSheets, Object oXls, Object oXlss,
String s, String sBook, String sName

Let h =GetRealWindow(GetFocus())
Let oApp =GetObject("Excel.Application")
If oApp Then
Let iGet =True
;Let iVisible =oApp.visible
Let iAlerts =oApp.DisplayAlerts
Let iUpdating =oApp.ScreenUpdating
Else
Let oApp =ObjectCreate("Excel.Application")
EndIf
If oApp Then
;Let oApp.Visible =False
Let oApp.DisplayAlerts =False
Let oApp.ScreenUpdating =False
Let oXlss =oApp.Workbooks
Let oXls =oXlss.Open(sSource, 0, -1) ; parameters are UpdateLinks, ReadOnly
Let oSheets =oXls.Sheets
Let iSheetCount =oSheets.Count
Let iSheet =1
While iSheet <=iSheetCount
Let oSheet =oSheets.Item(iSheet)
FileDelete(sTarget)
oSheet.SaveAs(sTarget, 21) ; parameters are format
Let sName =oSheet.Name
Let oSheet = oNull
Let oSheets = oNull
oXls.Close(0) ; parameters are SaveChanges
Let oXls = oNull
Let s =sName +"\r\n\r\n" +FileToString(sTarget)
Let sBook =sBook +StringIf(iSheet >1, "\12", "") +s
Let oXls =oXlss.Open(sSource, 0, -1) ; parameters are UpdateLinks, ReadOnly
Let oSheets =oXls.Sheets
Let iSheet =iSheet +1
EndWhile
StringToFile(sBook, sTarget)
Let oSheet = oNull
Let oSheets = oNull
oXls.Close(0) ; parameters are SaveChanges
Let oXls = oNull
Let oXlss = oNull
If iGet Then
;Let oApp.Visible =iVisible
Let oApp.DisplayAlerts =iAlerts
Let oApp.ScreenUpdating =iUpdating
Else
oApp.Quit()
EndIf
Else
SSay("Error!")
EndIf

Let oApp = oNull
Return FileExists(sTarget)
EndFunction

Int Function ConvertFileToText(String sSource, String sTarget)
;Convert another format to text using a technique based on its extension
Var
Int iHtm, Int iReturn,
String s, String sExt

If FileDelete(sTarget) Then
Let sExt =PathGetExtension(sSource)
If sExt =="htm" Then
ConvertHTMLToText(sSource, sTarget)
ElIf sExt =="doc" || sExt =="rtf" Then
ConvertWordToText(sSource, sTarget)
ElIf sExt =="ppt" Then
ConvertPowerPointToText(sSource, sTarget)
ElIf sExt =="xls" Then
ConvertExcelToText(sSource, sTarget)
EndIf
Else
SayString("Cannot delete existing target file!")
EndIf
Return FileExists(sTarget)
EndFunction

String Function ConvertHTMLToString(String sFile)
;Get text of an HTML file
Var
Object oNull, Object oApp,
Object oBody, Object oDoc, Object oRange,
String sText

Let oApp =ObjectCreate("InternetExplorer.Application")
Let oApp.Visible =False
Let oApp.Silent =True
oApp.navigate(sFile)
While oApp.ReadyState!=4
Delay(1)
EndWhile
Let oDoc =oApp.Document
If oDoc Then
Let oBody =oDoc.Body
If oBody Then
Let oRange =oBody.CreateTextRange()
If oRange Then
Let sText =oRange.Text
Let oRange = oNull
EndIf
Let oBody = oNull
EndIf
oDoc.Close()
Let oDoc = oNull
oApp.Quit()
EndIf
Let oApp = oNull
Return sText
EndFunction

Int Function ConvertHTMLToText(String sSource, String sTarget)
;Convert .htm or .html to .txt file using Microsoft Internet Explorer
Var
String s

Let s =ConvertHTMLToString(sSource)
StringToFile(s, sTarget)
Return FileExists(sTarget)
EndFunction

String Function ConvertJssToJsd(String sJss)
;Refresh .jsd from .jss file
Var
Int i, Int iLeft, Int iRight, Int iLine, Int iLineCount, Int iParamCount, Int iParam,
String sLineBreak, String sTempJsd, String s, String sJsd, String sBody, String sLine, String sWord1, String sWord2, String sWord3, String sParamList, String sParam, String sOutput

Let sJsd = StringChopRight(sJss, 3) + "jsd"
If FileExists(sJsd) Then
FileCopy(sJsd, sJsd + ".bak")
Let sBody = FileToString(sJsd)
Let sBody = RegExpReplaceEquiv(sBody, "^\\s*:(Function|Script)\\s+(\\S+)", "[$1 $2]")
Let sBody = RegExpReplaceEquiv(sBody, "^\\s*:(Synopsis|Description)\\s+(.*?)\r\n", "$1=$2\r\n")
Let sBody = RegExpReplaceEquiv(sBody, "^\\s*:.*?\n", "")
Let sTempJsd = PathGetTempFolder() + "\\temp.ini"
;FileDelete(sTempJsd)
StringToFile(sBody, sTempJsd)
Pause()
EndIf
If !FileExists(sJss) Then
Return False
EndIf

Let sLineBreak = "\r\n"
Let sBody = FileToString(sJss)
Let sBody = RegExpReplaceCase(sBody, sLineBreak, "\n")
Let iLineCount = StringCountSegment(sBody, "\n")
Let iLine = 1
While (iLine <= iLineCount)
Let sLine = StringTrim(StringSegment(sBody, "\n", iLine))
If !StringLeadEquiv(sLine, ";") && StringContains(sLine, "(") && StringContains(sLine, ")") && (StringContainsEquiv(sLine, "Script") || StringContainsEquiv(sLine, "Function")) Then
Let sLine = StringReplaceCase(sLine, "(", " (")
Let sLine = StringReplaceAllCase(sLine, "  ", " ")
If StringLeadEquiv(sLine, "Function") Then
Let sLine = "Void " + sLine
EndIf
Let sWord1 = StringSegment(sLine, " ", 1)
Let sWord2 = StringSegment(sLine, " ", 2)
Let sWord3 = StringSegment(sLine, " ", 3)
If StringEquiv(sWord1, "Script") Then
Let sOutput = sOutput + ":Script " + sWord2 + sLineBreak
Let s = IniReadString("Script " + sWord2, "Synopsis", "", sTempJsd)
If StringIsBlank(s) Then
Let s = StringTrim(StringSegment(sBody, "\n", iLine + 1))
Let s = StringIf(StringLeadCase(s, ";"), StringTrim(StringChopLeft(s, 1)), "")
EndIf
If !StringIsBlank(s) Then
Let sOutput = sOutput + ":Synopsis " + s + sLineBreak
EndIf

Let s = IniReadString("Script " + sWord2, "Description", "", sTempJsd)
If StringIsBlank(s) Then
Let i = iLine + 2
While StringLeadCase(StringTrim(StringSegment(sBody, "\n", i)), ";")
Let s = s + StringTrim(StringChopLeft(StringTrim(StringSegment(sBody, "\n", i)), 1)) + "  "
Let i = i + 1
EndWhile
Let s = StringTrim(s)
EndIf
If !StringIsBlank(s) Then
Let sOutput = sOutput + ":Description " + s + sLineBreak
EndIf
Let sOutput = sOutput + sLineBreak
ElIf StringEquiv(sWord2, "Function") Then
Let sOutput = sOutput + ":Function " + sWord3 + sLineBreak
Let s = IniReadString("Function " + sWord3, "Synopsis", "", sTempJsd)
If StringIsBlank(s) Then
Let s = StringTrim(StringSegment(sBody, "\n", iLine + 1))
Let s = StringIf(StringLeadCase(s, ";"), StringTrim(StringChopLeft(s, 1)), "")
EndIf
If !StringIsBlank(s) Then
Let sOutput = sOutput + ":Synopsis " + s + sLineBreak
EndIf
Let s = IniReadString("Function " + sWord3, "Description", "", sTempJsd)
If StringIsBlank(s) Then
Let i = iLine + 2
While StringLeadCase(StringTrim(StringSegment(sBody, "\n", i)), ";")
Let s = s + StringTrim(StringChopLeft(StringTrim(StringSegment(sBody, "\n", i)), 1)) + "  "
Let i = i + 1
EndWhile
Let s = StringTrim(s)
EndIf

If !StringIsBlank(s) Then
Let sOutput = sOutput + ":Description " + s + sLineBreak
EndIf
If !StringEquiv(sWord1, "Void") Then
Let sOutput = sOutput + ":Returns " + sWord1 + sLineBreak
EndIf
Let iLeft = StringContains(sLine, "(")
Let iRight = StringContains(sLine, ")")
Let sParamList = Substring(sLine, iLeft + 1, iRight - iLeft - 1)
Let iParamCount = StringCountSegment(sParamList, ",")
Let iParam = 1
While iParam <= iParamCount
Let sParam = StringTrim(StringSegment(sParamList, ",", iParam))
If !StringIsBlank(sParam) Then
Let sParam = StringReplaceEx(sParam, " ", "/", True)
If iParamCount == 1 Then
Let sOutput = sOutput + ":Param " + sParam + sLineBreak
Else
Let sOutput = sOutput + ":Param" + IntToString(iParam) + " " + sParam + sLineBreak
EndIf
EndIf
Let iParam = iParam + 1
EndWhile
Let sOutput = sOutput + sLineBreak
EndIf
EndIf
Let iLine = iLine + 1
EndWhile

FileDelete(sTempJsd)
Return sOutput
EndFunction

Int Function ConvertPowerPointToText(String sSource, String sTarget)
;Convert .ppt to .txt file using PowerPoint
Var
Handle h, Handle hApp,
Int i, Int iAlerts, Int iGet, Int iNote,
Int iNoteLabel, Int iNoteCount, Int iOutlineLabel, Int iShape,
Int iShapeCount, Int iShip, Int iShipCount,
Int iSlide, Int iSlideCount, Int iType, Int iVisible,
Object oNull, Object oApp, Object oFrame,
Object oNote, Object oNotes, Object oPpt, Object oPpts,
Object oShape, Object oShapes, Object oShip, Object oShips,
Object oSlide, Object oSlides, Object oText, Object oTextEffect,
Object oTextFrame, Object oTextRange, Object oType,
String s, String sText

Let h =GetRealWindow(GetFocus())
Let oApp =GetObject("PowerPoint.Application")
If oApp Then
Let iGet =True
;Let iVisible =oApp.visible
Let iAlerts =oApp.DisplayAlerts
Else
Let oApp =ObjectCreate("PowerPoint.Application")
EndIf
If oApp Then
;Must be visible for COM automation to work
Let oApp.Visible =True
Let oApp.DisplayAlerts =False
Let oPpts =oApp.Presentations
;Let oPpt =oPpts.Open(sSource, -1, 0, 0) ; parameters are ReadOnly, Untitled, WithWindow
Let oPpt =oPpts.Open(sSource, True)

; get presentation name and slide count
Let s =oPpt.Name
Let sText =PathGetBase(s)

Let oSlides =oPpt.Slides
Let iSlideCount =oSlides.Count
Let sText =sText +"\r\n" +IntToString(iSlideCount) +" Slide" +StringIf(iSlideCount ==1, "", "s")
Let iSlide =1
While iSlide <=iSlideCount
Let oSlide =oSlides.Item(iSlide)
Let sText =sText +"\r\n\r\n\r\n" +"Slide " +IntToString(iSlide)

Let oNotes =oSlide.NotesPage
Let iNoteCount =oNotes.Count
Let iNoteLabel =True
Let iNote =1
While iNote <=iNoteCount
Let oNote =oNotes.Item(iNote)
Let oShips =oNote.Shapes
Let iShipCount =oShips.Count
Let iShip =1
While iShip <=iShipCount
Let oShip =oShips.Item(iShip)
If oShip.HasTextFrame Then
Let oFrame =oShip.TextFrame
Let oText =oFrame.TextRange
Let s =oText.Text
If s !="" Then
If iNoteLabel Then
Let sText =sText +"\r\n" +"Notes:" +"\r\n" +s
Let iNoteLabel =False
Else
Let sText =sText +"\r\n" +s
EndIf
EndIf
Let oText = oNull
Let oFrame = oNull
EndIf
Let oShip = oNull
Let iShip =iShip +1
EndWhile
Let oShips = oNull
Let oNote = oNull
Let iNote =iNote +1
EndWhile
Let oNotes = oNull

Let oShapes =oSlide.Shapes
Let iShapeCount =oShapes.Count
Let iOutlineLabel =True
Let iShape =1
While iShape <=iShapeCount
Let oShape =oShapes.Item(iShape)
If oShape.HasTextFrame Then
Let oTextFrame =oShape.TextFrame
Let oTextRange =oTextFrame.TextRange
Let s =oTextRange.Text
If s !="" && !StringEquiv(s, "outline") Then
If iOutlineLabel Then
Let sText =sText +"\r\n" +"Outline:" +"\r\n" +s
Let iOutlineLabel =False
Else
Let sText =sText +"\r\n" +s
EndIf
EndIf
Let oTextRange = oNull
Let oTextFrame = oNull
EndIf
If !oShape.HasTextFrame || (oShape.HasTextFrame && s =="") Then
Let s =oShape.AlternativeText
If s !="" Then
Let sText =sText +"\r\n" +s
EndIf
Let iType =oShape.Type
If iType ==15 Then ; texteffect
Let oTexteffect =oShape.TextEffect
Let s =oTexteffect.Text
If s !="" && (s !=oShape.alternativetext) Then
Let sText =sText +"\r\n" +"Text Effect: " +s
EndIf
Let oTextEffect = oNull
EndIf
Let oType = oNull
EndIf

Let oShape = oNull
Let iShape =iShape +1
EndWhile
Let oShapes = oNull
Let oSlide = oNull
Let iSlide =iSlide +1
EndWhile
Let oSlides = oNull
StringToFile(sText, sTarget)
Let oPpt.Saved =True
oPpt.close()
Let oPpt = oNull
Let oPpts = oNull
If iGet Then
Let oApp.Visible =iVisible
Let oApp.DisplayAlerts =iAlerts
Else
oApp.Quit()
EndIf
Else
SSay("Error!")
EndIf
Let oApp = oNull
return fileexists(sTarget)
EndFunction

Int Function ConvertWordToText(String sSource, String sTarget)
;Convert .doc or .rtf to .txt file using Microsoft Word
Var
Handle h, Handle hApp,
Int iLength, Int i, Int iAlerts, Int iGet,
Int iReturn, Int iVisible,
Object oNull, Object oApp, Object oDoc, Object oSelection,
Object oDocs, Object oRange, Object oTemplate,
String sText

Let iReturn =True
Let h =GetRealWindow(GetFocus())
Let oApp =GetObject("Word.Application")
If oApp Then
Let iGet =True
;Let iVisible =oApp.visible
Let iAlerts =oApp.DisplayAlerts
Else
Let oApp =ObjectCreate("Word.Application")
;Let oApp.Visible =True
EndIf
If oApp Then
Let oApp.DisplayAlerts =False
;Let oApp.Visible =False
Let oDocs =oApp.documents
Let oDoc =oDocs.Open(sSource, 0, -1, 0) ; parameters are confirm conversions, ReadOnly, AddToRecentFiles
Let oSelection = oApp.Selection
Let iLength = oSelection.StoryLength
oSelection.SetRange(0, iLength)
Let sText = oSelection.Text
Let sText = RegExpReplaceCase(sText, "\r\f", "\f\r")
Let sText = RegExpReplaceCase(sText, "\r", "\r\n")
StringToFile(sText, sTarget)
;Let i =oDoc.SaveAs(sTarget, 2, 0, "", 0) ; parameters are format, LockComments, password, AddToRecentFiles
oDoc.Close(0) ; parameters are SaveChanges
Let oDoc = oNull
Let oDocs = oNull
Let oTemplate =oApp.NormalTemplate
Let oTemplate.Saved =True
Let oTemplate = oNull
If iGet Then
;Let oApp.visible =iVisible
Let oApp.DisplayAlerts =iAlerts
Else
oApp.Quit(0) ; parameters are SaveChanges
EndIf
Else
SSay("Error!")
Let iReturn =False
EndIf
Let oApp = oNull
Return iReturn
EndFunction

String Function DialogBrowseForFolder(String sDir)
;Choose folder from a standard Windows dialog, or simple input box below Windows 2000
Var
Handle WINDOW_HANDLE,
Int NO_OPTIONS,
Object oShell, Object oFolder, Object oItem, Object oPath, Object oNull,
String sTitle, String sReturn

;If GetWindowsOS() == OS_WIN_NT Then
Let oShell = ObjectCreate("Shell.Application")
If oShell Then
;Let WINDOW_HANDLE = 0
;Let WINDOW_HANDLE = GetAppMainWindow(GetFocus())
;Let NO_OPTIONS = 0
ScheduleFunction("DialogBrowseForFolderHelper", 0)
Let oFolder = oShell.BrowseForFolder(WINDOW_HANDLE, sTitle, NO_OPTIONS, sDir)
If oFolder Then
Let oItem = oFolder.Self
Let sReturn = oItem.Path
EndIf
Else
If InputBox("Folder Name:", "Input", sDir) Then
Let sReturn = sDir
EndIf
EndIf

Let oItem = oNull
Let oFolder = oNull
Let oShell = oNull
Return sReturn
EndFunction

Void Function DialogBrowseForFolderHelper()
;AppActivateTitle("Browse for Folder")
Let sHomerTitle = "Browse for Folder"
UIWaitForTitleAndActivate(sHomerTitle, 10)
EndFunction

String Function DialogConfirm(String sTitle, String sMessage, String sDefault)
;Get choice from a standard Yes, No, or Cancel dialog
Var
Int iFlags, Int iChoice,
String sReturn

Let iFlags = MB_YESNOCANCEL | MB_ICONQuestion
Let sTitle = StringIf(StringIsBlank(sTitle), "Confirm", sTitle)
If StringEquiv(sDefault, "Y") Then
Let iFlags = iFlags | MB_DEFBUTTON1
Else
Let iFlags = iFlags | MB_DEFBUTTON2
EndIf
Let iChoice = ExMessageBox(sMessage, sTitle, iFlags)
If iChoice == IDYES Then
Let sReturn = "Y"
ElIf iChoice == IDNO Then
Let sReturn = "N"
EndIf
Return sReturn
EndFunction

Void Function DialogHelpMenu()
;Version of HelpMenuEx with common defaults
Return DialogHelpMenuEx("", "", True)
EndFunction

Void Function DialogHelpMenuEx(String sTitle, String sSection, Int iSort)
;Dynamically create list of scripts to invoke based on section of a .jkm file
Var
Int iCount, Int i, Int iLoop,
String sItem, String sItems, String sFile, String sKey, String sKeyList,
String sList, String sValue, String sValueList

Let sTitle = StringIf(StringIsBlank(STitle), "Help Menu", sTitle)
Let sSection = StringIf(StringIsBlank(SSection), "Common Keys", sSection)
Let sFile =GetJAWSSettingsDirectory() +"\\" +GetActiveConfiguration() +".jkm"
Let sKeyList =iniReadSectionKeys(sSection, sFile)
Let iCount = StringCountSegment(sKeyList, "|")
Let iLoop =1
While iLoop <= iCount
Let sKey =StringSegment(sKeyList, "|", iLoop)
If !StringIsBlank(sKey) && !StringLeadCase(StringTrim(sKey), ";") Then
Let sValue = StringTrim(iniReadString(sSection, sKey, "", sFile))
If StringLeadEquiv(sValue, "UI") Then
Let sValue = StringPadRight(StringChopLeft(sValue, 2), " ", 40) + StringPadRight(sValue, " ", 40)
Else
Let sValue = StringPadRight(sValue, " ", 40) + StringPadRight(sValue, " ", 40)
EndIf
Let sValueList =sValueList + sValue + "\7"
Let sList =sList + StringPadRight(StringTrim(StringLeft(sValue, 40)) + "\t" + sKey, " ", 80) + "\7"
EndIf
Let iLoop =iLoop +1
EndWhile
Let sValueList =StringChopRight(sValueList, 1)
Let sList = StringChopRight(sList, 1)
If iSort Then
Let sValueList = StringSegmentSort(sValueList, "\7")
Let sList = StringSegmentSort(sList, "\7")
EndIf
If True Then
;If False Then
Let sValue = DialogPick(sTitle, sList, False)
If StringIsBlank(sValue) Then
Return
EndIf
Let i = StringIndexSegmentCase(sList, "\7", sValue)
Let sValue = StringSegment(sValueList, "\7", i)
Let sValue = StringTrim(StringRight(sValue, 40))
Else
Let i =dlgSelectItemInList(sList, sTitle, False)
Let sValue =StringSegment(sValueList, "\7", i)
If StringIsBlank(sValue) Then
Return
EndIf
EndIf

ScheduleFunction("$" +sValue, 0)
EndFunction

String Function DialogInput(String sTitle, String sField, String sValue)
;Get input from a single edit box
Var
Handle h,
Int iChoice,
String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
If StringIsBlank(sTitle) Then
Let sTitle = "Input"
EndIf
If !StringIsBlank(sField) && StringChopRight(sField, 1) != ":" Then
Let sField = sField + ":"
EndIf

If InputBox(sField, sTitle, sValue) Then
Let sReturn = sValue
EndIf
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
Return sReturn
EndFunction

String Function DialogPick(String sTitle, String sValues, Int iSort)
;Get choice from a standard list box
Var
Handle h,
Int iChoice,
String sReturn

If StringIsBlank(sTitle) Then
Let sTitle = "Pick"
EndIf
If iSort Then
Let sValues = StringSegmentSort(sValues, "\7")
EndIf
Let iChoice = DlgSelectItemInList(sValues, sTitle, False)
Pause()
If iChoice Then
Let sReturn = StringSegment(sValues, "\7", iChoice)
EndIf
Return sReturn
EndFunction

String Function DialogOpenFile(String sDir)
;Choose file from a standard Windows dialog, or simple input box below Windows XP

;String Function DialogOpenFile(String sDir, String sFilter, Int iIndex)
Var
Int iIndex,
Object oDlg, Object oNull,
String sFilter, String sReturn

Let oDlg = ObjectCreate("UserAccounts.CommonDialog")
If oDlg Then
If StringIsBlank(sFilter) Then
;Let oDlg.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
Let oDlg.Filter = "All Files (*.*)|*.*"
Else
Let oDlg.Filter = sFilter
EndIf
If iIndex == 0 Then
Let oDlg.FilterIndex = 1
Else
Let oDlg.FilterIndex = iIndex
EndIf
Let oDlg.InitialDir = sDir
ScheduleFunction("DialogOpenFileHelper", 0)
If oDlg.ShowOpen() Then
Let sReturn = oDlg.FileName
EndIf
Else
Let sDir = StringIf(FolderExists(sDir), sDir + "\\", "")
If InputBox("File Name:", "Open", sDir) Then
Let sReturn = sDir
EndIf
EndIf

Let oDlg = oNull
Return sReturn
EndFunction

Void Function DialogOpenFileHelper()
;AppActivateTitle("Open")
Let sHomerTitle = "Open"
UIWaitForTitleAndActivate(sHomerTitle, 10)
EndFunction

String Function DialogSaveFile(String sFile)
;Choose file from a standard Windows dialog, or simple input box below Windows XP

;Int Function DialogSaveFile(String sFile, String sType)
Var
Object oDlg, Object oNull,
String sType, String sReturn

Let oDlg = ObjectCreate("SAFRCFileDlg.FileSave")
If oDlg Then
Let oDlg.FileName = sFile
;Let oDlg.Title = "Save"
If StringIsBlank(sType) Then
Let oDlg.FileType = "All Files (*.*)"
Else
Let oDlg.FileType = sType
EndIf
ScheduleFunction("DialogSaveFileHelper", 0)
If oDlg.OpenFileSaveDlg() Then
Let sReturn = oDlg.FileName
EndIf
Else
If InputBox("File Name:", "Save", sFile) Then
Let sReturn = sFile
EndIf
EndIf

Let oDlg = oNull
Return sReturn
EndFunction

Void Function DialogSaveFileHelper()
;AppActivateTitle("Save")
Let sHomerTitle = "Save As"
UIWaitForTitleAndActivate(sHomerTitle, 10)
EndFunction

Int Function EditGetIndex(Handle h)
;Get cursor position as index from beginning of text
Return EditGetSelectionEnd(h)
EndFunction

Int Function EditGetIndexColumn(Handle h, Int iIndex)
;Get column of cursor position
Var
Int iLine, Int iReturn, Int iStart

If !h Then
Let h = GetFocus()
EndIf

If iIndex == - 1 Then
Let iIndex = EditGetSelectionEnd(h)
EndIf

Let iLine = EditGetIndexLine(h, iIndex)
;Pause()
Let iStart = EditGetLineStart(h, iLine)
Let iReturn = iIndex - iStart
Return iReturn
EndFunction

Int Function EditGetIndexLine(Handle h, Int iIndex)
;Get line of cursor position
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xC9 ;EM_LINEFROMCHAR
Return SendMessage(h, iMessage, iIndex, 0)
EndFunction

Int Function EditGetLineCount(Handle h)
;Get number of lines of text
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xBA ;EM_GETLINECOUNT
Return SendMessage(h, iMessage, 0, 0)
EndFunction

Int Function EditGetLineEnd(Handle h, Int iLine)
;Get index of position following a line
Var
Int iReturn, Int iStart

If !h Then
Let h = GetFocus()
EndIf

Let iStart = EditGetLineStart(h, iLine)
If iStart < 0 Then
Return -1
EndIf
Pause()
Let iReturn = EditGetLineStart(h, iLine + 1)
Pause()
If iReturn < 0 Then
Let iReturn = EditGetSize(h)
EndIf
Return iReturn
EndFunction

Int Function EditGetLineLength(Handle h, Int iLine)
;Get length of a line
Var
Int iReturn, Int iIndex

If !h Then
Let h = GetFocus()
EndIf

Let iIndex = EditGetLineStart(h, iLine)
If iIndex < 0 Then
Return -1
EndIf

Pause()
Let iReturn = EditGetLineLengthFromIndex(h, iIndex)
Return iReturn
EndFunction

Int Function EditGetLineLengthFromIndex(Handle h, Int iIndex)
;Get length of line containing index
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xC1 ;EM_LINELENGTH
Return SendMessage(h, iMessage, iIndex, 0)
EndFunction

Int Function EditGetLineStart(Handle h, Int iLine)
;Get index at start of line
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xBB ;EM_LINEINDEX
Return SendMessage(h, iMessage, iLine, 0)
EndFunction

String Function EditGetLineText(Handle h, Int iLine)
;Get text of a line
Var
Int iIndex, Int iStart, Int iEnd, Int iLength,
String sReturn

If !h Then
Let h = GetFocus()
EndIf

If iLine == -1 Then
Let iIndex = EditGetIndex(h)
Let iLine = EditGetIndexLine(h, iIndex)
EndIf

Let iStart = EditGetLineStart(h, iLine)
If iStart < 0 Then
Return ""
EndIf
Pause()
Let iEnd = EditGetLineStart(h, iLine + 1)
Pause()
If iEnd < 0 Then
Let iEnd = EditGetSize(h)
EndIf
Let sReturn = EditGetRange(h, iStart, iEnd)
Return sReturn
EndFunction

Int Function EditGetModified(Handle h)
;Test whether text in control has been saved to disk
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xB8 ;EM_GETMODIFY
Return SendMessage(h, iMessage, 0, 0)
EndFunction

Int Function EditGetMultiline(Handle h)
;Test for Multiline style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x4 ;ES_MULTILINE
Return (iStyleBits & iStyle)
EndFunction

String Function EditGetRange(Handle h, Int iStart, Int iEnd)
;Get text between lower and upper positions, excluding end point
Var
String sText, String sReturn

If !h Then
Let h = GetFocus()
EndIf

Let sText = EditGetText(h)
Let iStart = iStart + 1
Let iEnd = iEnd + 1
Let sReturn = StringGetRange(sText, iStart, iEnd)
Return sReturn
EndFunction

Int Function EditGetReadOnly(Handle h)
;Test for ReadOnly style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x800 ;%ES_READONLY
Return (iStyleBits & iStyle)
EndFunction

String Function EditGetSelectedText(Handle h)
;Get selected text
Var
Int iStart, Int iEnd,
String sText, String sReturn

If !h Then
Let h = GetFocus()
EndIf

Let sText = EditGetText(h)
Let iStart = EditGetSelectionStart(h) + 1
Let iEnd = EditGetSelectionEnd(h) + 1
Let sReturn = StringGetRange(sText, iStart, iEnd)
Return sReturn
EndFunction

Int Function EditGetSelectionEnd(Handle h)
;Get index at end of selection
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xB0 ;EM_GETSEL
Return MathHighWord(SendMessage(h, iMessage, 0, 0))
EndFunction

Int Function EditGetSelectionStart(Handle h)
;Get index at start of selection
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xB0 ;EM_GETSEL
Return MathLowWord(SendMessage(h, iMessage, 0, 0))
EndFunction

Int Function EditGetSize(Handle h)
;Get size of text
Var
Int iReturn, Int iLine, Int iStart, Int iLength

If !h Then
Let h = GetFocus()
EndIf

Let iLine = EditGetLineCount(h) - 1
Let iStart = EditGetLineStart(h, iLine)
Let iLength = EditGetLineLengthFromIndex(h, iStart)
Let iReturn = iStart + iLength
Return iReturn
EndFunction

String Function EditGetText(Handle h)
;Get all text
Var
Int i, Int iCol, Int iRow,
Object oNull, Object o,
String sReturn

If !h Then
Let h = GetFocus()
EndIf

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let sReturn = o.AccValue

Let o= oNull
Return sReturn
EndFunction

Int Function EditGoBottom(Handle h)
;Go to bottom position
Var
Int iIndex

If !h Then
Let h = Getfocus()
EndIf
;Return EditSetSelection(h, 65000, 65000)
Let iIndex = EditGetSize(h) - 1
EditSetIndex(h, iIndex)
EndFunction

Int Function EditGoTop(Handle h)
;Go to top position
If !h Then
Let h = Getfocus()
EndIf
;Return EditSetSelection(h, 0, 0)
EditSetIndex(h, 0)
EndFunction

Int Function EditReplaceRange(Handle h, Int iStart, Int iEnd, String sReplace)
;Replace text between lower and upper positions, excluding end point
Var
String sText

If !h Then
Let h = GetFocus()
EndIf

Let sText = EditGetText(h)
Let iStart = iStart + 1
Let iEnd = iEnd + 1
Let sText = StringReplaceRange(sText, iStart, iEnd, sReplace)
EditSetText(h, sText)
EndFunction

Int Function EditScrollCaret(Handle h)
;Adjust view of text to ensure that caret is visible
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xB7 ;EM_SCROLLCARET
Return SendMessage(h, iMessage, 0, 0)
EndFunction

Int Function EditSelectAll(Handle h)
;Select all text
If !h Then
Let h = Getfocus()
EndIf
EditSetSelection(h, 0, -1)
EditScrollCaret(h)
EndFunction

Int Function EditSelectChunk(Handle h)
;Select chunk of text at cursor, or extend selection to next chunk
Var
Int iIndex, Int iStart, Int iEnd, Int iForward, Int iBackward,
String sValue, String sText, String sMatch, String sReturn, String sResult

If !h Then
Let h = GetFocus()
EndIf

Let iIndex = EditGetIndex(h)
Let iStart = iIndex
Let iEnd = EditGetSize(h)
Let sText = EditGetRange(h, iStart, iEnd)
Let sMatch = "\\s+"
Let sResult = RegExpContainsCase(sText, sMatch)
Let iForward = StringToInt(StringSegment(sResult, "\7", 1))
Let iForward = iIndex + iForward - 1
Let iStart = 0
Let iEnd = iIndex
Let sText = EditGetRange(h, iStart, iEnd)
Let sResult = RegExpContainsLastEquiv(sText, sMatch)
Let iBackward = StringToInt(StringSegment(sResult, "\7", 1))
Let sValue = StringSegment(sResult, "\7", 2)
Let iBackward = iBackward + StringLength(sValue) - 1
EditSetSelection(h, iBackward, iForward)
EndFunction

Int Function EditSetIndex(Handle h, Int iIndex)
;Set cursor position
If !h Then
Let h = GetFocus()
EndIf

EditSetSelection(h, iIndex, iIndex)
Pause()
EditScrollCaret(h)
EndFunction

Int Function EditSetModified(Handle h, Int iState)
;Set modified status of text
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xB9 ;EM_SETMODIFY
Return SendMessage(h, iMessage, iState, 0)
EndFunction

Int Function EditSetSelection(Handle h, Int iStart, Int iEnd)
;Select text within range
Var
Int iMessage

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0xB1 ;EM_SETSEL
Return SendMessage(h, iMessage, iStart, iEnd)
EndFunction

Int Function EditSetText(Handle h, String sText)
;Set all text
Var
Int i, Int iCol, Int iRow, Int iReturn,
Object oNull, Object o,
String sReturn

If !h Then
Let h = GetFocus()
EndIf

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Pause()
Let o.AccValue = sText
;Let iReturn = StringEqual(o.AccValue, sText)

Let o= oNull
Return iReturn
EndFunction

Int Function EditUnselectAll(Handle h)
;Unselect all text
Var
Int iIndex

If !h Then
Let h = GetFocus()
EndIf

Let iIndex = EditGetIndex(h)
EditSetSelection(h, -1, 0)
EditSetIndex(h, iIndex)
EditScrollCaret(h)
EndFunction

Int Function FileCopy(String sSource, String sTarget)
;Copy source to destination file, replacing if it exists

Var
Object oSystem, Object oNull,
Int iReturn

If FileDelete(sTarget) Then
Let oSystem =ObjectCreate("Scripting.FileSystemObject")
oSystem.CopyFile(sSource, sTarget)
Let iReturn = FileExists(sTarget)
Let oSystem = oNull
EndIf

Return iReturn
EndFunction

Int Function FileDelete(String sFile)
;Delete a file if it exists, and test whether it is subsequently absent
;either because it was successfully deleted or because it was not present in the first place
Var
Object oSystem, Object oNull

If FileExists(sFile) Then
Let oSystem =ObjectCreate("Scripting.FileSystemObject")
oSystem.DeleteFile(sFile, True)
Let oSystem = oNull
EndIf

Return !FileExists(sFile)
EndFunction

Int Function FileGetSize(String sFile)
;Get size of a file
Var
Object oSystem, Object oFile, Object oNull,
Int iReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let oFile =oSystem.GetFile(sFile)
Let iReturn =oFile.size

Let oFile = oNull
Let oSystem = oNull
Return iReturn
EndFunction

Int Function FileMove(String sSource, String sTarget)
;Move source to destination file, replacing if it exists

Var
Object oSystem, Object oNull,
Int iReturn

If FileDelete(sTarget) Then
Let oSystem =ObjectCreate("Scripting.FileSystemObject")
oSystem.MoveFile(sSource, sTarget)
Let iReturn = FileExists(sTarget)
Let oSystem = oNull
EndIf

Return iReturn
EndFunction

String Function FileToString(String sFile)
;Get content of text (not binary) file

Var
Int iRead,
Object oSystem, Object oFile, Object oNull,
String sReturn

Let oSystem =ObjectCreate("Scripting.FilesystemObject")
Let oFile =oSystem.OpenTextFile(sFile)
Let sReturn =oFile.ReadAll()
oFile.close()

Let oFile = oNull
Let oSystem = oNull
Return sReturn
EndFunction

Int Function FolderCopy(String sSource, String sTarget)
;Copy source to destination Folder, replacing if it exists

Var
Object oSystem, Object oNull,
Int iReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
oSystem.CopyFolder(sSource, sTarget)
Let iReturn = FolderExists(sTarget)
Let oSystem = oNull

Return iReturn
EndFunction

Int Function FolderCreate(String sFolder)
;Create folder
Var
Object oSystem, Object oNull

If FolderExists(sFolder) Then
Return True
EndIf
Let oSystem =ObjectCreate("Scripting.FileSystemObject")
oSystem.Createfolder(sFolder)
Let oSystem = oNull

Return FolderExists(sFolder)
EndFunction

Int Function FolderDelete(String sFolder)
;Delete a Folder if it exists, and test whether it is subsequently absent
;either because it was successfully deleted or because it was not present in the first place
Var
Object oSystem, Object oNull

If FolderExists(sFolder) Then
Let oSystem =ObjectCreate("Scripting.FileSystemObject")
oSystem.DeleteFolder(sFolder, True)
Let oSystem = oNull
EndIf

Return !FolderExists(sFolder)
EndFunction

Int Function FolderExists(String sPath)
;Test whether folder exists

Var
Int iReturn,
Object oSystem, Object oNull

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let iReturn =oSystem.FolderExists(sPath)

Let oSystem =oNull
Return iReturn
EndFunction

Int Function FolderGetSize(String sFolder)
;Get size of folder, summing the sizes of files and subfolders it contains

Var
Object oSystem,
Object oFolder, Object oNull,
Int iReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let oFolder =oSystem.GetFolder(sFolder)
Let iReturn =oFolder.size

Let oFolder = oNull
Let oSystem = oNull
Return iReturn
EndFunction

Int Function FolderIsRoot(String sFolder)
;Test whether folder is root of a drive
Return StringEqual(sFolder, "\\") || (IntEqual(StringLength(sFolder), 3) && (StringContainsEquiv("abcdefghijklmnopqrstuvwxyz", StringLeft(sFolder, 1)) && StringEqual(StringRight(sFolder, 2), ":\\")))
EndFunction

Int Function FolderMove(String sSource, String sTarget)
;Move source to destination Folder, replacing if it exists

Var
Object oSystem, Object oNull,
Int iReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
oSystem.MoveFolder(sSource, sTarget)
Let iReturn = FolderExists(sTarget)
Let oSystem = oNull

Return iReturn
EndFunction

Void Function IBox(Int i)
Var
Int iSpeech

Let iSpeech = VoiceSetSpeech(True)
MessageBox(IntToString(i))
VoiceSetSpeech(iSpeech)
EndFunction

Int Function IniDeleteFile(String sIni)
;Delete a .ini file, including cached version in memory
Var
Int iLoop, Int iCount,
String sSection, String sSections

If !FileExists(sIni) Then
Return True
EndIf

Let sSections = IniReadSectionNames(sIni)
Let iCount = StringCountSegment(sSections, "|")
Let iLoop = 1
While iLoop <= iCount
Let sSection = StringSegment(sSections, "|", iLoop)
IniRemoveSection(sSection, sIni)
Let iLoop = iLoop + 1
EndWhile
IniFlushFile(sIni)
Pause()
FileDelete(sIni)
Pause()
EndFunction

Int Function IniFlushFile(String sIni)
;Flush cached version in memory to disk
Return IniWriteString("", "", "", sIni)
;IniFlush(sIni)
;Pause()
EndFunction

String Function IEEval(String sText)
;Get evaluation of a JavaScript expression using IE
Var
Object oApp, Object oNull,
String sClipboard, String sUrl, String sReturn

Let sClipboard = GetClipboardText()
Pause()
Let sUrl = "javascript:window.clipboardData.setData('Text', eval(\34" + sText + "\34).toString());"
Let oApp = ObjectCreate("InternetExplorer.Application")
Let oApp.Visible = False
oApp.Navigate(sUrl)Let oApp.Silent = True

Pause()
oApp.Quit()
Let oApp = oNull

Let sReturn = GetClipboardText()
CopyToClipboard(sClipboard)
Return sReturn
EndFunction

String Function IniFormDialogMultiInput(String sTitle, string sFields, String sValues)
;Get input from multiple edit boxes using IniForm.exe
Var
Handle h,
Int iLoop, Int iCount,
String sTempDir, String sBody, String sField, String sValue, String sCommand, String sExe, String sIni, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
FileCopy(GetJAWSSettingsDirectory() + "\\" + GetActiveConfiguration() + ".jcf", GetJAWSSettingsDirectory() + "\\IniForm.jcf")
Let sExe = GetJAWSSettingsDirectory() + "\\Homer\\IniForm.exe"
Let sExe = PathGetShort(sExe)
Let sTempDir = PathCreateTempFolder()
Let sIni = sTempDir + "\\Input.ini"
Let sCommand = sExe + " " + sTempDir + "\\"
Let iCount = StringCountSegment(sFields, "\7")
If stringIsBlank(sTitle) Then
If iCount == 1 Then
Let sTitle = "Input"
Else
Let sTitle = "Fields"
EndIf
EndIf
SayString(sTitle)

If StringIsBlank(sFields) && iCount == 1 Then
Let sFields = "Text"
EndIf

Let sBody = "[" + sTitle + "]\r\nMisc=NoStatus\r\n"

Let iLoop = 1
While iLoop <= iCount
Let sField = StringSegment(sFields, "\7", iLoop)
Let sValue = StringSegment(sValues, "\7", iLoop)
Let sBody = sBody + "[" + sField + "]\r\nValue=" + sValue + "\r\n"
Let iLoop = iLoop + 1
EndWhile
Let sBody = sBody + "[OK]\r\nControl=Button\r\n"
Let sBody = sBody + "[Cancel]\r\nControl=Button\r\n"
IniDeleteFile(sIni)
StringToFile(sBody, sIni)
Let h = GetFocus()
ShellRun(sCommand, 1, True)
Pause()
If h Then
SetFocus(h)
Pause()
EndIf

FileDelete(sIni)
Let sIni = sTempDir + "\\Output.ini"
If FileExists(sIni) Then
Let iLoop = 1
While iLoop <= iCount
Let sField = StringSegment(sFields, "\7", iLoop)
If iLoop == iCount Then
Let sReturn = sReturn + IniReadString("Results", sField, "", sIni)
Else
Let sReturn = sReturn + IniReadString("Results", sField, "", sIni) + "\7"
EndIf
Let iLoop = iLoop + 1
EndWhile
FileDelete(sIni)
EndIf
FolderDelete(sTempDir)
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
return sReturn
EndFunction

Void Function IniFormDialogHelper()
;TypeKey("Home")
UIWaitForTitleAndActivate(sHomerTitle, 10)
EndFunction

Void Function IniFormDialogInfo(String sTitle, String sText)
;Display information in a multiline edit box using IniForm.exe
Var
Handle h,
Int iLoop, Int iCount,
String sTempDir, String sBody, String sTxt, String sCommand, String sExe, String sIni, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
FileCopy(GetJAWSSettingsDirectory() + "\\" + GetActiveConfiguration() + ".jcf", GetJAWSSettingsDirectory() + "\\IniForm.jcf")
Let sExe = GetJAWSSettingsDirectory() + "\\Homer\\IniForm.exe"
Let sExe = PathGetShort(sExe)
Let sTempDir = PathCreateTempFolder()
Let sIni = sTempDir + "\\Input.ini"
Let sCommand = sExe + " " + sTempDir + "\\"
Let stxt = sTempDir + "\\Input.txt"
If StringIsBlank(sTitle) Then
Let sTitle = "Info"
EndIf

Let sBody = "[" + sTitle + "]\r\nMemoWidth=300\r\nMemoHeight=300\r\nMisc=NoStatus\r\n"
Let sBody = sBody + "[InfoView]\r\nControl=Memo\r\n"
Let sBody = sBody + "Misc=NoLabel|ReadOnly\r\n"
Let sBody = sBody + "[Close]\r\nControl=Button\r\nID=2\r\n"
StringToFile(sBody, sIni)

Let sText = "[[InfoView]]\r\n" + sText
StringToFile(sText, sTxt)

Let h = GetFocus()
ShellRun(sCommand, 1, True)
Pause()
If h Then
SetFocus(h)
Pause()
EndIf
FolderDelete(sTempDir)
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
EndFunction

String Function IniFormDialogInput(String sTitle, string sField, String sValue)
Return IniFormDialogMultiInput(sTitle, sField, sValue)
EndFunction

String Function IniFormDialogPick(String sTitle, String sValues, Int iSort)
;Get choice from a single selection list box using IniForm.exe
Var
Handle h,
Int iLoop, Int iCount,
String sTempDir, String sBody, String sTxt, String sCommand, String sExe, String sIni, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
FileCopy(GetJAWSSettingsDirectory() + "\\" + GetActiveConfiguration() + ".jcf", GetJAWSSettingsDirectory() + "\\IniForm.jcf")
Let sExe = GetJAWSSettingsDirectory() + "\\Homer\\IniForm.exe"
Let sExe = PathGetShort(sExe)
Let sTempDir = PathCreateTempFolder()
Let sIni = sTempDir + "\\Input.ini"
Let sCommand = sExe + " " + sTempDir + "\\"
Let stxt = sTempDir + "\\Input.txt"
If StringIsBlank(sTitle) Then
Let sTitle = "Pick"
EndIf
SayString(sTitle)

Let sBody = "[" + sTitle + "]\r\nListWidth=300\r\nListHeight=300\r\nMisc=NoStatus\r\n"
Let sBody = sBody + "[List]\r\nControl=List\r\n"
;Let sValues = StringReplaceEx(sValues, "\7", "|", True)
;Let sBody = sBody + "Range=" + sValues + "\r\n"
Let sValues = RegExpReplaceCase(sValues, "\7", "\r\n")
Let sValues = "[[List]]\r\n" + sValues
StringToFile(sValues, sTxt)
Let sBody = sBody + "Selection=1" + "\r\n"
If iSort Then
Let sBody = sBody + "Misc=NoLabel|Sort\r\n"
Else
Let sBody = sBody + "Misc=NoLabel\r\n"
EndIf
Let sBody = sBody + "[OK]\r\nControl=Button\r\n"
Let sBody = sBody + "[Cancel]\r\nControl=Button\r\n"
StringToFile(sBody, sIni)
Let h = GetFocus()
ScheduleFunction("IniFormDialogHelper", 0)
ShellRun(sCommand, 1, True)
Pause()
If h Then
SetFocus(h)
Pause()
EndIf

FileDelete(sIni)
Let sIni = sTempDir + "\\Output.ini"
If FileExists(sIni) Then
;CopyToClipboard(FileToString(sini))
Let sReturn = IniReadString("Results", "List", "", sIni)
Let sReturn = StringReplaceEx(sReturn, "|", "\7", True)
FileDelete(sIni)
EndIf
FolderDelete(sTempDir)
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
return sReturn
EndFunction

String Function IniFormDialogMemo(String sTitle, String sText)
;Get input from a multiline edit box using IniForm.exe
Var
Handle h,
Int iLoop, Int iCount,
String sTempDir, String sBody, String sTxt, String sCommand, String sExe, String sIni, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
FileCopy(GetJAWSSettingsDirectory() + "\\" + GetActiveConfiguration() + ".jcf", GetJAWSSettingsDirectory() + "\\IniForm.jcf")
Let sExe = GetJAWSSettingsDirectory() + "\\Homer\\IniForm.exe"
Let sExe = PathGetShort(sExe)
Let sTempDir = PathCreateTempFolder()
Let sIni = sTempDir + "\\Input.ini"
Let sCommand = sExe + " " + sTempDir + "\\"
Let stxt = sTempDir + "\\Input.txt"
If StringIsBlank(sTitle) Then
Let sTitle = "Memo"
EndIf

Let sBody = "[" + sTitle + "]\r\nMemoWidth=300\r\nMemoHeight=300\r\nMisc=NoStatus\r\n"
Let sBody = sBody + "[MemoEdit]\r\nControl=Memo\r\n"
Let sBody = sBody + "Misc=NoLabel\r\n"
Let sBody = sBody + "[OK]\r\nControl=Button\r\n"
Let sBody = sBody + "[Cancel]\r\nControl=Button\r\n"
StringToFile(sBody, sIni)

Let sText = "[[MemoEdit]]\r\n" + sText
StringToFile(sText, sTxt)

Let h = GetFocus()
ShellRun(sCommand, 1, True)
Pause()
If h Then
SetFocus(h)
Pause()
EndIf

FileDelete(sIni)
FileDelete(sTxt)
Let sIni = sTempDir + "\\Output.ini"
If FileExists(sIni) Then
Let stxt = sTempDir + "\\Output.txt"
Let sReturn = FileToString(sTxt)
Let sReturn = Substring(sReturn, StringLength("[[MemoEdit]]\r\n"), StringLength(sReturn))
FileDelete(sIni)
FileDelete(sTxt)
EndIf
FolderDelete(sTempDir)
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
Return sReturn
EndFunction

String Function IniFormDialogMultiPick(String sTitle, String sValues, Int iSort)
;Get choices from a multiple selection list box using IniForm.exe
Var
Handle h,
Int iLoop, Int iCount,
String sTempDir, String sBody, String sTxt, String sCommand, String sExe, String sIni, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
FileCopy(GetJAWSSettingsDirectory() + "\\" + GetActiveConfiguration() + ".jcf", GetJAWSSettingsDirectory() + "\\IniForm.jcf")
Let sExe = GetJAWSSettingsDirectory() + "\\Homer\\IniForm.exe"
Let sExe = PathGetShort(sExe)
Let sTempDir = PathCreateTempFolder()
Let sIni = sTempDir + "\\Input.ini"
Let sCommand = sExe + " " + sTempDir + "\\"
Let stxt = sTempDir + "\\Input.txt"
If StringIsBlank(sTitle) Then
Let sTitle = "Multi Pick"
EndIf
;SayString(sTitle)

;Let sBody = "[" + sTitle + "]\r\nMultiWidth=300\r\nMultiHeight=300\r\nMisc=NoStatus\r\n"
Let sBody = "[" + sTitle + "]\r\nMultiWidth=600\r\nMultiHeight=600\r\nMisc=NoStatus\r\n"
Let sBody = sBody + "[Multi]\r\nControl=Multi\r\n"
;Let sValues = StringReplaceEx(sValues, "\7", "|", True)
;Let sBody = sBody + "Range=" + sValues + "\r\n"
Let sValues = RegExpReplaceCase(sValues, "\7", "\r\n")
Let sValues = "[[Multi]]\r\n" + sValues
StringToFile(sValues, sTxt)
If iSort Then
Let sBody = sBody + "Misc=NoLabel|Sort\r\n"
Else
Let sBody = sBody + "Misc=NoLabel\r\n"
EndIf
Let sBody = sBody + "[OK]\r\nControl=Button\r\n"
Let sBody = sBody + "[Cancel]\r\nControl=Button\r\n"
StringToFile(sBody, sIni)
Let h = GetFocus()
Let sHomerTitle = sTitle
ScheduleFunction("IniFormDialogHelper", 0)
ShellRun(sCommand, 1, True)
Pause()
If h Then
SetFocus(h)
Pause()
EndIf

FileDelete(sIni)
Let sIni = sTempDir + "\\Output.ini"
If FileExists(sIni) Then
Let sReturn = IniReadString("Results", "Multi", "", sIni)
Let sReturn = StringReplaceEx(sReturn, "|", "\7", True)
FileDelete(sIni)
EndIf
FolderDelete(sTempDir)
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
return sReturn
EndFunction

Int Function IniFormIniEditSection(String sSection, String sIni)
;Edit values in a section of a .ini file using IniForm.exe
Var
Handle h,
Int iReturn, Int iLoop, Int iCount,
String sFields, String sValues, String sField, String sValue

Let sFields = IniReadSectionKeys(sSection, sIni)
Let sFields = StringReplaceEx(sFields, "|", "\7", True)
Let iCount = StringCountSegment(sFields, "\7")
Let iLoop = 1
While iLoop <= iCount
Let sField = StringSegment(sFields, "\7", iLoop)
Let sValue = IniReadString(sSection, sField, "", sIni)
If iLoop == iCount Then
Let sValues = sValues + sValue
Else
Let sValues = sValues + sValue + "\7"
EndIf
Let iLoop = iLoop + 1
EndWhile
Let sValues = IniFormDialogMultiInput("Fields", sFields, sValues)
Let iReturn = !StringIsBlank(sValues)
If iReturn Then
Let iLoop = 1
While iLoop <= iCount
Let sField = StringSegment(sFields, "\7", iLoop)
Let sValue = StringSegment(sValues, "\7", iLoop)
IniWriteQuote(sSection, sField, sValue, sIni)
Let iLoop = iLoop + 1
EndWhile
EndIf
Return iReturn
EndFunction

String Function IniReadSetting(String sKey, String sDefaultValue)
;Get an application setting
Return IniReadString("Internal", sKey, sDefaultValue, GetActiveConfiguration() + ".ini")
EndFunction

Int Function IniWriteQuote(String sSection, String sKey, String sValue, String sIni)
;Write a .ini value enclosed in quotes
Let sValue = "\34" + sValue + "\34"
;Write a .ini value enclosed in quotes
IniWriteString(sSection, sKey, sValue, sIni, True)
;Pause()
IniFlush(sIni)
EndFunction

Int Function IniWriteSetting(String sKey, String sValue)
;Write an application setting
Return IniWriteQuote("Internal", sKey, sValue, GetActiveConfiguration() + ".ini")
EndFunction

Int Function IntEqual(Int i1, Int i2)
;Test whether two integers are equal
Return (i1 == i2)
EndFunction

Int Function IntIf(Int iCondition, Int iTrue, Int iFalse)
;Return one of two integers, depending on condition

Var
Int iReturn

If iCondition Then
Let iReturn = iTrue
Else
Let iReturn = iFalse
EndIf

Return iReturn
EndFunction

Void Function iSay(Int i)
;Say an integer regardless of speech on/off status (used in debugging)

SSay(IntToString(i))
EndFunction

String Function JSDialogMultiInput(String sTitle, string sFields, String sValues)
;Get input from multiple edit boxes using IronCOM.dll
Var
Handle h,
Int iCount,
String sCode, String sJS, String sReturn

If !JSInit() Then
Return IniFormDialogMultiInput(sTitle, sFields, sValues)
EndIf

Let h = GetFocus()
Let iCount = StringCountSegment(sFields, "\7")
If stringIsBlank(sTitle) Then
If iCount == 1 Then
Let sTitle = "Input"
Else
Let sTitle = "Fields"
EndIf
EndIf

If StringIsBlank(sFields) && iCount == 1 Then
Let sFields = "Text"
EndIf

Let sJS = GetJAWSSettingsDirectory() + "\\Homer\\MultiInput.JS"
Let sCode = FileToString(sJS)
;ScheduleFunction("JSDialogHelper", 0)
Let sReturn = JSEval(sCode, sTitle, sFields, sValues, "\7")

If h Then
SetFocus(h)
EndIf

return sReturn
EndFunction

String Function JSEval(String sCode, String s1, String s2, String s3, String s4)
;Evaluate with JSEval object, passing four string parameters and returning a string result
Var
Int iCreate,
Object oNull,
String sReturn

If !JSInit() Then
Return
EndIf

Let sReturn = oHomerJS.Eval(sCode, s1, s2, s3, s4)
Return sReturn
EndFunction

String Function JSEvalFile(String sFile, String s1, String s2, String s3, String s4)
;Evaluate file with JSEval object, passing four string parameters and returning a string result
Var
String sCode, String sReturn

If !StringContains(sFile, "\\") Then
Let sFile = GetJAWSSettingsDirectory() + "\\Homer\\" + sFile
EndIf

Let sCode = FileToString(sFile)
Let sReturn = JSEval(sCode, s1, s2, s3, s4)
Return sReturn
EndFunction

Int Function JSInit()
;Try to initiate global JScript object if it does not exist, and then test whether it exists

If iJSInitialized == -1 Then
Return False
ElIf oHomerJS Then
Return True
EndIf

Let oHomerJS = CreateObjectEx("Iron.JS", False)
If oHomerJS Then
Let iJSInitialized = 1
Return True
Else
SayString("Error initializing JScript component!")
Let iJSInitialized = -1
Return False
EndIf
EndFunction

Int Function JSShellUrlToFile(String sUrl, String sFile)
;Download file using IronCOM.dll
Var
Int iReturn,
String sCode

If !FileDelete(sFile) Then
Return
EndIf

Let sCode = "var o = new Microsoft.VisualBasic.Devices.Network();\n"
Let sCode = sCode + "o.DownloadFile(s1, s2);\n"
JSEval(sCode, sUrl, sFile, "", "")
Let iReturn = FileExists(sFile)
;isay(ireturn)
Return iReturn
EndFunction

Int Function KeyboardSend(String sKeys)
;Send keystrokes
Var
Int iReturn,
Object oNull, Object oShell

Let oShell =ObjectCreate("Wscript.Shell")
If oShell Then
Pause()
Delay(1)
Let iReturn =oShell.KeyboardSendKeys(sKeys)
Else
Let iReturn =-2
EndIf

Let oShell = oNull
Return iReturn
EndFunction

Void Function KeyboardType(String sText)
;Type a string of characters
Var
Int i, Int iLength,
String s

if 0 then
Return TypeString(sText)
endif

Let iLength =StringLength(sText)
Let i =1
While i <=iLength
Let s =Substring(sText, i, 10)
TypeString(s)
Pause()
Let i = i +10
EndWhile
EndFunction

Int Function KeyboardTypeAndWait(String sKeyList, String sControl, Int iMax)
;Type a list of keys and wait for a control class to receive focus
Var
Int i, Int iReturn,
String sClass, String sKey

SpeechOff()
Let i = 1
While i
Let sKey = StringSegment(sKeyList, "|", i)
If StringLength(sKey) == 0 Then
Let i = 0
Else
TypeKey(sKey)
Delay(1, False)
Let i = i + 1
EndIf
EndWhile
Let sClass = GetWindowClass(GetFocus())
Let i = 0
While StringLength(sClass) != StringLength(sControl) && sClass != sControl && i <= iMax
Delay(1, False)
Let i = i + 1
Let sClass = GetWindowClass(GetFocus())
EndWhile
If sClass == sControl Then
Let iReturn = True
EndIf
SpeechOn()
Return iReturn
EndFunction

Void Function KeyboardTypeList(String sKeyList, String sDelimiter)
;Type keys from a delimited list

Var
Int i,
String sKey

SpeechOff()
Pause()
Let i =1
While i >0
Let sKey =StringSegment(sKeyList, sDelimiter, i)
If sKey =="" Then
Let i =0
Else
TypeKey(sKey)
Pause()
Let i =i +1
EndIf
EndWhile
SpeechOn()
EndFunction

String Function LbcDialogMultiInput(String sTitle, string sFields, String sValues)
;Get input from multiple edit boxes using LbC.dll
Var
Handle h,
Int iLoop, Int iCount,
Object oLbc, Object oNull,
String sTempDir, String sBody, String sField, String sValue, String sCommand, String sExe, String sIni, String sReturn

If !LbcInit() Then
Return IniFormDialogMultiInput(sTitle, sFields, sValues)
EndIf

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
Let iCount = StringCountSegment(sFields, "\7")
If stringIsBlank(sTitle) Then
If iCount == 1 Then
Let sTitle = "Input"
Else
Let sTitle = "Fields"
EndIf
EndIf
;SayString(sTitle)

If StringIsBlank(sFields) && iCount == 1 Then
Let sFields = "Text"
EndIf

Let sReturn = oHomerLbc.FieldDialog(sTitle, sFields, sValues)
If h Then
SetFocus(h)
EndIf

If StringIsBlank(sReturn) Then
Return
EndIf

Let sReturn = RegExpReplaceCase(sReturn, "\t", "\7")
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
Let oLbc = oNull
return sReturn
EndFunction

Void Function LbcDialogHelper()
if 0 then
;SpeechOff()
AppActivateTitle("JAWS")
Pause()
;SpeechOn()
endif
UIWaitForTitleAndActivate(sHomerTitle, 10)
EndFunction

String Function LbcDialogInput(String sTitle, string sField, String sValue)
Return LbcDialogMultiInput(sTitle, sField, sValue)
EndFunction

String Function LbcDialogPick(String sTitle, String sValues, Int iSort)
;Get choices from a single selection list box using lbc.dll
Var
Handle h,
Int iLoop, Int iCount,
String sTempDir, String sBody, String sField, String sValue, String sCommand, String sExe, String sIni, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
If !LbcInit() Then
Return IniFormDialogPick(sTitle, sValues, iSort)
EndIf

Let h = GetFocus()
If stringIsBlank(sTitle) Then
Let sTitle = "Pick"
EndIf
Let sHomerTitle = sTitle
Let sValues = RegExpReplaceCase(sValues, "\7", "\t")
ScheduleFunction("LbcDialogHelper", 0)
Let sReturn = oHomerLbc.ListDialog(sTitle, sValues, iSort)
If h Then
SetFocus(h)
EndIf

If StringIsBlank(sReturn) Then
Return
EndIf

Let sReturn = RegExpReplaceCase(sReturn, "\t", "\7")
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
return sReturn
EndFunction

String Function LbcDialogMultiPick(String sTitle, String sValues, Int iSort)
;Get choices from a multiple selection list box using lbc.dll
Var
Handle h,
Int iLoop, Int iCount,
Object oLbc, Object oNull,
String sTempDir, String sBody, String sField, String sValue, String sCommand, String sExe, String sIni, String sReturn

If !LbcInit() Then
Return IniFormDialogMultiPick(sTitle, sValues, iSort)
EndIf

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
If stringIsBlank(sTitle) Then
Let sTitle = "Multi Pick"
EndIf
;SayString(sTitle)

Let sValues = RegExpReplaceCase(sValues, "\7", "\t")
ScheduleFunction("LbcDialogHelper", 0)
Let sReturn = oHomerLbc.MultiListDialog(sTitle, sValues, iSort)
If h Then
SetFocus(h)
EndIf

If StringIsBlank(sReturn) Then
Return
EndIf

Let sReturn = RegExpReplaceCase(sReturn, "\t", "\7")
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
Let oLbc = oNull
return sReturn
EndFunction

Int Function LbcInit()
Var
Int iReturn

If !oHomerLbc Then
SayString("Initializing .NET component")
SayString("Please wait")
Let oHomerLbc = CreateObjectEx("LayoutByCode.Lbc", False)
SayString("Done!")
EndIf

If oHomerLbc Then
Let iReturn = True
Else
Let iReturn = False
EndIf
Return iReturn
EndFunction

Int Function ListGetExtended(Handle h)
;Test for extended selection style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x800 ;LBS_EXTENDEDSEL
Return (iStyleBits & iStyle)
EndFunction

Int Function ListGetMulti(Handle h)
;Test for multiple selection style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x8 ;LBS_MULTIPLESEL
Return (iStyleBits & iStyle)
EndFunction

Int Function ListGetSort(Handle h)
;Test for Sort style
Var
Int iStyleBits, Int iStyle

If !h Then
Let h = GetFocus()
EndIf

Let iStyleBits = GetWindowStyleBits(h)
Let iStyle = 0x2 ;LBS_SORT
Return (iStyleBits & iStyle)
EndFunction

String Function MathDecIntToHexString(Int iDec)
;Convert decimal integer to hex string
Var
Int iLeadingZero, Int i, Int iPower, Int iDiv,
String s, String sLookup, String sReturn

Let iLeadingZero = True
Let sReturn = "0x"
Let sLookup = "123456789abcdef"
Let iPower = 7
While iPower >= 0
Let iDiv = MathIntToPower(16, iPower)
Let i = iDec / iDiv
If i Then
Let s = Substring(sLookup, i, 1)
Let iLeadingZero = False
Else
Let s = "0"
EndIf
If !iLeadingZero Then
Let sReturn = sReturn + s
EndIf
Let iDec = iDec % iDiv
Let iPower = iPower - 1
EndWhile
Return sReturn
EndFunction

Int Function MathHexStringToDecInt(String sHex)
;Convert hex string to decimal integer
Var
Int iReturn, Int i, Int iPower,
String s, String sLookup

Let sLookup = "123456789abcdef"
Let sHex = StringLower(sHex)
If StringLeft(sHex, 2) == "0x" Then
Let sHex = StringChopLeft(sHex, 2)
EndIf
Let sHex = StringRight("0x0000000" + sHex, 8)

Let iPower = 7
While iPower >= 0
Let s = SubString(sHex, 8 - iPower, 1)
Let i = StringContains(sLookup, s)
Let iReturn = iReturn + i * MathIntToPower(16, iPower)
Let iPower = iPower - 1
EndWhile
Return iReturn
EndFunction

Int Function MathHighWord(Int iDec)
;Get high word of an integer
Var
Int i, Int iPower, Int iDiv, Int iReturn

Let iPower = 7
While iPower >= 4
Let iDiv = MathIntToPower(16, iPower)
Let i = iDec / iDiv
Let iReturn = iReturn + i * MathIntToPower(16, iPower - 4)
Let iDec = iDec % iDiv
Let iPower = iPower - 1
EndWhile
Return iReturn
EndFunction

Int Function MathIntToPower(Int i, Int iPower)
;Get result of a number raised to a power
Var
Int iReturn

Let iReturn = 1
While iPower > 0
Let iReturn = iReturn * i
Let iPower = iPower - 1
EndWhile
Return iReturn
EndFunction

Int Function MathLowWord(Int iDec)
;Get low word of an integer
Var
Int iPower, Int iDiv, Int iReturn

Let iPower = 4
Let iDiv = MathIntToPower(16, iPower)
Let iReturn = iDec % iDiv
Return iReturn
EndFunction

Int Function MenuInvokeID(Handle hAppWindow, Int iID)
;Invoke a standard menu item by its ID
Var
Int iMessage

If !hAppWindow Then
Let hAppWindow = GetAppMainWindow(GetFocus())
EndIf
Let iMessage = 273 ;WM_SYSCOMMAND
Return SendMessage(hAppWindow, iMessage, iID, 0)
EndFunction

Int Function MSAAGetChildCount(Handle h)
;Get the number of MSAA children in the client area of a window
Var
Int i,
Object o, Object oNull,
String s

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let i = o.AccChildCount
Let o = oNull
Return i
EndFunction

String Function MSAAGetDefaultAction(Handle h)
;Get the MSAA default action
Var
Int i,
Object o, Object oNull,
String s

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let s = o.AccDefaultAction(0)
Let o = oNull
Return s
EndFunction

String Function MSAAGetDescription(Handle h)
; Get the MSAA description
Var
Int i,
Object o, Object oNull,
String s

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let s = o.AccDescription(0)
Let o = oNull
Return s
EndFunction

Int Function MSAAGetFocus(Handle h)
; Get the MSAA Focus of a subitem
Var
Int i, Int iReturn,
Object o, Object oNull

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let iReturn = o.AccFocus(0)

Let o = oNull
Return iReturn
EndFunction

String Function MSAAGetHelp(Handle h)
;Get MSAA help text
Var
Int i,
Object o, Object oNull,
String s

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let s = o.AccHelp(0)
Let o = oNull
Return s
EndFunction

String Function MSAAGetInfo(Handle h)
;Get string of various MSAA information
Var
Int i, Int iCount, Int iFocus, Int iChildren,
Object o, Object oNull,
String sReturn, String sRole, String sName, String sValue, String sState, String sHelp, String sDescription, String sKeyboardShortcut, String sDefaultAction

Let h = GetFocus()
Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let iCount = o.AccChildCount(0)
Let i = 0
While i <= iCount
Let sRole = GetRoleText(o.AccRole(i))
If !StringIsBlank(sRole) Then Let sReturn = sReturn + "Role=" + sRole + "\n" EndIf
Let sName = o.AccName(i)
If !StringIsBlank(sName) Then Let sReturn = sReturn + "Name=" + sName + "\n" EndIf
Let sValue = o.AccValue(i)
If !StringIsBlank(sValue) Then Let sReturn = sReturn + "Value=" + sValue + "\n" EndIf
Let iFocus = o.AccFocus(i)
If iFocus Then Let sReturn = sReturn + "Focus=" + IntToString(iFocus) + "\n" EndIf
Let sState = MSAAGetStateText(o.AccState(i))
If !StringIsBlank(sState) Then Let sReturn = sReturn + "State=" + sState + "\n" EndIf
Let sHelp = o.AccHelp(i)
If !StringIsBlank(sHelp) Then Let sReturn = sReturn + "Help=" + sHelp + "\n" EndIf
Let sDescription = o.AccDescription(i)
If !StringIsBlank(sDescription) Then Let sReturn = sReturn + "Description=" + sDescription + "\n" EndIf
Let sKeyboardShortcut = o.AccKeyboardShortcut(i)
If !StringIsBlank(sKeyboardShortcut) Then Let sReturn = sReturn + "KeyboardShortcut=" + sKeyboardShortcut + "\n" EndIf
Let sDefaultAction = o.AccDefaultAction(i)
If !StringIsBlank(sDefaultAction) Then Let sReturn = sReturn + "DefaultAction=" + sDefaultAction + "\n" EndIf
Let iChildren = o.AccChildCount(i)
If iChildren Then Let sReturn = sReturn + "Children=" + IntToString(iChildren) + "\n" EndIf
Let i = i + 1
EndWhile
Let o = oNull
Return sReturn
EndFunction

String Function MSAAGetKeyboardShortcut(Handle h)
;Get MSAA keyboard shortcut
Var
Int i,
Object o, Object oNull,
String s

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let s = o.AccKeyboardShortcut(0)
Let o = oNull
Return s
EndFunction

String Function MSAAGetName(Handle h)
;Get MSAA name
Var
Int i,
Object o, Object oNull,
String s

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let s = o.AccName(0)
Let o = oNull
Return s
EndFunction

Int Function MSAAGetRole(Handle h)
; Get MSAA numeric role
Var
Int i, Int iReturn,
Object o, Object oNull

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let iReturn = o.AccRole(0)

Let o = oNull
Return iReturn
EndFunction

String Function MSAAGetRoleString(Handle h)
;Get text of MSAA numeric role
Return GetRoleText(MSAAGetRole(h))
EndFunction

Int Function MSAAGetState(Handle h)
;Get MSAA numeric state
Var
Int i,
Object o, Object oNull,
String s

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let i = o.AccState(0)
Let o = oNull
Return i
EndFunction

String Function MSAAGetStateString(Handle h)
;Get MSAA state text
Return GetStateText(MSAAGetState(h))
EndFunction

String Function MSAAGetStateText(Int iBits)
;Get composite text of MSAA states

Var
string sReturn

If iBits & msaa_state_unavailable then
let sReturn = sReturn + "unavailable "
endIf
If iBits & msaa_state_selected then
let sReturn = sReturn + "selected "
endIf
If iBits & msaa_state_focused then
let sReturn = sReturn + "focused "
endIf
If iBits & msaa_state_pressed then
let sReturn = sReturn + "pressed "
endIf
If iBits & msaa_state_checked then
let sReturn = sReturn + "checked "
endIf
If iBits & msaa_state_mixed then
let sReturn = sReturn + "mixed "
endIf
If iBits & msaa_state_readonly then
let sReturn = sReturn + "readonly "
endIf
If iBits & msaa_state_hottracked then
let sReturn = sReturn + "hottracked "
endIf
If iBits & msaa_state_default then
let sReturn = sReturn + "default "
endIf
If iBits & msaa_state_expanded then
let sReturn = sReturn + "expanded "
endIf
If iBits & msaa_state_collapsed then
let sReturn = sReturn + "collapsed "
endIf
If iBits & msaa_state_busy then
let sReturn = sReturn + "busy "
endIf
If iBits & msaa_state_floating then
let sReturn = sReturn + "floating "
endIf
If iBits & msaa_state_marqueed then
let sReturn = sReturn + "marqueed "
endIf
If iBits & msaa_state_animated then
let sReturn = sReturn + "animated "
endIf
If iBits & msaa_state_invisible then
let sReturn = sReturn + "invisible "
endIf
If iBits & msaa_state_offscreen then
let sReturn = sReturn + "offscreen "
endIf
If iBits & msaa_state_sizeable then
let sReturn = sReturn + "sizeable "
endIf
If iBits & msaa_state_moveable then
let sReturn = sReturn + "moveable "
endIf
If iBits & msaa_state_selfvoicing then
let sReturn = sReturn + "selfvoicing "
endIf
If iBits & msaa_state_focusable then
let sReturn = sReturn + "focusable "
endIf
If iBits & msaa_state_selectable then
let sReturn = sReturn + "selectable "
endIf
If iBits & msaa_state_linked then
let sReturn = sReturn + "linked "
endIf
If iBits & msaa_state_traversed then
let sReturn = sReturn + "traversed "
endIf
If iBits & msaa_state_multiselectable then
let sReturn = sReturn + "multiselectable "
endIf
If iBits & msaa_state_extselectable then
let sReturn = sReturn + "extselectable "
endIf
If iBits & msaa_state_alert_low then
let sReturn = sReturn + "alert_low "
endIf
If iBits & msaa_state_alert_medium then
let sReturn = sReturn + "alert_medium "
endIf
If iBits & msaa_state_alert_high then
let sReturn = sReturn + "alert_high "
endIf
If iBits & msaa_state_protected then
let sReturn = sReturn + "protected "
endIf
If iBits & msaa_state_valid then
let sReturn = sReturn + "valid "
endIf
Let sReturn = StringTrimTrailingBlanks(sReturn)
Return sReturn
EndFunction

String Function MSAAGetValue(Handle h)
;Get MSAA value
Var
Int i,
Object o, Object oNull,
String s

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, 1, i)
Let s = o.AccValue(0)
Let o = oNull
Return s
EndFunction

Int Function MSAASetSelection(Handle h, Int iState)
;Select subitems using MSAA
Var
Int i, Int iCol, Int iRow, Int iReturn,
Object oNull, Object o,
String sReturn

If !h Then
Let h = GetFocus()
EndIf

Let o = GetObjectFromEvent(h, msaa_OBJID_CLIENT, iState, i)
o.AccSelect(iState, i)

Let o= oNull
Return iReturn
EndFunction

Object Function ObjectCreate(String sObject)
;Create a COM object from a ProgID string, trying an alternate method if the default one fails

Var
Object oReturn

Let oReturn =CreateObjectEx(sObject, True)
If !oReturn Then
Let oReturn =CreateObjectEx(sObject, False)
EndIf
Return oReturn
EndFunction

String Function PathCombine(String sFolder, String sPath)
;Combine a folder and path, ensuring a single backslash seperater
Var
String sReturn

Let sReturn = sFolder + "\\" + sPath
Let sReturn = StringReplaceAllCase(sReturn, "\\\\", "\\")
Return sReturn
EndFunction

String Function PathCreateTempFolder()
;Create temporary folder
Var
String sReturn

Let sReturn = PathGetTempFolder() + "\\" + PathGetTempName()
FolderCreate(sReturn)
If FolderExists(sReturn) Then
Return sReturn
EndIf
EndFunction

String Function PathGetBase(String sFile)
;Get base/root name

Var
Object oSystem, Object oNull,
String sReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let sReturn =oSystem.PathGetBase(sFile)

Let oSystem = oNull
Return sReturn
EndFunction

String Function PathGetCurrentDirectory()
;Get current directory of JAWS application

Var
Object oShell, Object oNull,
String sReturn

Let oShell =ObjectCreate("Wscript.Shell")
Let sReturn =oShell.CurrentDirectory

Let oShell = oNull
Return sReturn
EndFunction

String Function PathGetDir(String sDir, String sWildcards, String sFlags)
;Get a list of paths, specifying folder, wild card pattern, and sort order
Var
Int iWait, Int iWindowStyle,
String sReturn, String sCommand, String sTempFile

Let sCommand = "%COMSPEC% /c dir /b " + sFlags + " \34" + sDir + "\\" + sWildcards + "\34"
Let sTempFile = PathGetTempFolder() + "\\" + PathGetTempname()
Let sCommand = sCommand + " >" + sTempFile
Let iWindowStyle = 0 ;hidden
Let iWait = True
ShellRun(sCommand, iWindowStyle, iWait)
pause()
Let sReturn = StringTrim(FileToString(sTempFile))
FileDelete(sTempFile)
Return sReturn
EndFunction

String Function PathGetExtension(String sFile)
;Get extention

Var
Object oSystem, Object oNull,
String sReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let sReturn =oSystem.GetExtensionName(sFile)

Let oSystem = oNull
Return sReturn
EndFunction

String Function PathGetFolder(String sPath)
;Get the folder containing a path

Var
Object oSystem, Object oNull,
String sReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let sReturn =oSystem.GetParentFolderName(sPath)
If !StringIsBlank(sReturn) && !StringContains(sReturn, "\\") Then
Let sReturn =sReturn +"\\"
EndIf

Let oSystem = oNull
Return sReturn
EndFunction

String Function PathGetHomer()
; Get path of Homer subfolder

Return GetJAWSSettingsDirectory() + "\\Homer\\"
EndFunction

String Function PathGetInternetCacheFolder()
;Get Windows folder for temporary Internet files
Var
Int TEMPORARY_INTERNET_FILES,
Object oShell, Object oFolder, Object oItem, Object oNull,
String sReturn

Let oShell = ObjectCreate("Shell.Application")
Let TEMPORARY_INTERNET_FILES = 32
Let oFolder = oShell.Namespace(TEMPORARY_INTERNET_FILES)
Let oItem = oFolder.Self
Let sReturn = oItem.Path

Let oItem = oNull
Let oFolder = oNull
Let oShell = oNull
Return sReturn
EndFunction

String Function PathGetLong(String sPath)
;Get long file name of file or folder
Var
Object oShell, Object oShortcut, Object oNull,
String sReturn

Let oShell = ObjectCreate("WScript.Shell")
Let oShortcut = oShell.CreateShortcut("temp.lnk")
Let oShortcut.TargetPath = sPath
Let sReturn = oShortcut.TargetPath

Let oShortcut = oNull
Let oShell = oNull
Return sReturn
EndFunction

String Function PathGetName(String sPath)
;Get the file or folder name at the end of a path

Var
Object oSystem, Object oNull,
String sReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let sReturn =oSystem.GetFileName(sPath)

Let oSystem = oNull
Return sReturn
EndFunction

String Function PathGetShort(String sPath)
;Get short path (8.3 style) of a file or folder

Var
Object oSystem, Object oFile, Object oFolder, Object oNull,
String sReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
If FolderExists(sPath) Then
Let oFolder =oSystem.GetFolder(sPath)
Let sReturn =oFolder.ShortPath
Else
Let oFile =oSystem.GetFile(sPath)
Let sReturn =oFile.ShortPath
EndIf

Let oFile = oNull
Let oFolder = oNull
Let oSystem = oNull
Return sReturn
EndFunction

String Function PathGetSpecialFolder(String s)
;Get a special folder of Windows
Var
Int i, Int iCount,
Object oNull, Object o, Object oShell,
String sReturn

Let oShell =ObjectCreate("WScript.Shell")
Let o =oShell.SpecialFolders
Let iCount =o.Count
Let i =1
While i <=iCount
Let sReturn =o.Item(i)
If StringEquiv(("\\" +s), Substring(sReturn, (StringLength(sReturn) -StringLength(s)), StringLength(sReturn))) Then
Let i = iCount +1
Else
Let i =i +1
EndIf
EndWhile

Let o = oNull
Let oShell = oNull
Return sReturn
EndFunction

String Function PathGetTempFile()
;Get a temporary file name
Return PathGetTempFolder() + "\\" + PathGetTempName()
EndFunction

String Function PathGetTempFolder()
;Get Windows folder for temporary files

Var
Int iFolder,
Object oSystem, Object oNull,
String sReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let iFolder = 2
Let sReturn =oSystem.GetSpecialFolder(iFolder).path

Let oSystem = oNull
Return sReturn
EndFunction

String Function PathGetTempName()
;Get Name for temporary file or folder
Var
Object oSystem, Object oNull,
String sReturn

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let sReturn = oSystem.GetTempName()

Let oSystem = oNull
Return sReturn
EndFunction

String Function PathGetType(String sPath)
;Get type as a folder or as a file type based on extension
Var
Object oSystem, Object oFile, Object oNull,
String sReturn

If FolderExists(sPath) Then
Let sReturn = "Folder"
Else
Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let oFile =oSystem.GetFile(sPath)
Let sReturn =oFile.Type
EndIf

Let oFile = oNull
Let oSystem = oNull
Return sReturn
EndFunction

String Function PathGetValidName(String sDir, String sBase, String sExt, Int iUnique)
;Get a Valid file name for a proposed name and parent folder, optionally unique with a suffix like _02
Var
Int i, Int iCount, 
String s, String sIllegal, String sLine, String sPrintable, String sSourceDir, String sSourceExt, String sSourceBase, String sTargetDir, String sTargetExt, String sTargetFile, String sTargetBase, String sViewable

Let sIllegal ="@%*+\\|;:'\34<>/?"
Let sViewable ="!\34#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
Let sPrintable =" " +sViewable

Let sSourceDir = sDir
Let sSourceBase = sBase
Let sSourceExt = sExt

Let sLine =sSourceBase
Let iCount = StringLength(sIllegal)
While i <= iCount
Let s = Substring(sIllegal, i, 1)
If StringContains(sLine, s) Then
Let sLine = StringReplaceAllCase(sLine, s, "_")
EndIf
Let i = i + 1
EndWhile

Let sLine = StringReplaceAllCase(sLine, "  ", " ")
Let sLine = StringReplaceAllCase(sLine, " _", "_")
Let sLine = StringReplaceAllCase(sLine, "_ ", "_")
Let sLine = StringReplaceAllCase(sLine, "__", "_")
Let sLine =StringTrim(sLine)

While StringLeft(sLine, 1) == "_"
Let sLine = StringChopLeft(sLine, 1)
EndWhile

While StringRight(sLine, 1) == "_"
Let sLine = StringChopRight(sLine, 1)
EndWhile

Let sLine =StringTrim(sLine)
Let sTargetBase =sLine
;Let sTargetFile =sSourceDir +"\\" +sTargetBase +sSourceExt
Let sTargetFile = PathCombine(sSourceDir, sTargetBase) + sSourceExt
If iUnique && FileExists(sTargetFile) Then
Let s ="_01"
;Let sTargetFile =sSourceDir +"\\" +sTargetBase +s +sSourceExt
Let sTargetFile = PathCombine(sSourceDir, sTargetBase) + s + sSourceExt
Let i =1
While FileExists(sTargetFile) && i <=99
Let i =i +1
Let s =IntToString(i)
Let s ="_" + StringPadLeft(s, "0", 2)
;Let sTargetFile =sSourceDir +"\\" +sTargetBase +s +sSourceExt
Let sTargetFile = PathCombine(sSourceDir, sTargetBase) + s + sSourceExt
EndWhile
EndIf
Return sTargetFile
EndFunction

Void Function PathSetCurrentDirectory(String sDir)
;Set current directory of JAWS application

Var
Object oShell, Object oNull,
String sReturn

Let oShell =ObjectCreate("Wscript.Shell")
Let oShell.CurrentDirectory = sDir

Let oShell = oNull
EndFunction

String Function PerlComboGetContents(Handle h)
;Get contents of item with focus in a combo box using PerlEval.dll
Var
String sCode, String sReturn

If !PerlInit() Then
Return
EndIf

If !h Then
Let h = GetFocus()
EndIf

Let sCode = "join '\7', Win32::GuiTest::GetComboContents($p1)"
Let sReturn = PerlEval(sCode, IntToString(h), "", "", "")
Return sReturn
EndFunction

String Function PerlComboGetText(Handle h)
;Get text of item with focus in a combo box using PerlEval.dll
Var
String sCode, String sReturn

If !PerlInit() Then
Return
EndIf

If !h Then
Let h = GetFocus()
EndIf

Let sCode = "Win32::GuiTest::GetComboText($p1)"
Let sReturn = PerlEval(sCode, IntToString(h), "", "", "")
Return sReturn
EndFunction

String Function PerlDateCalc(String sYear, String sMonth, String sWeek, String sDay)
;Calculate English date phrase
Var
String sReturn, String sCode

Let sYear = StringTrim(sYear)
Let sMonth = StringTrim(sMonth)
Let sWeek = StringTrim(sWeek)
Let sDay = StringTrim(sDay)

Let sCode = "package Date::Calc;\n"
If !StringToInt(sMonth) Then
Let sCode = sCode + "$p2 = Decode_Month($p2);\n"
EndIf
If !StringToInt(sDay) Then
Let sCode = sCode + "$p4 = Decode_Day_of_Week($p4);\n"
EndIf
If StringToInt(sWeek) Then
Let sCode = sCode + "($y, $m, $d) = Nth_Weekday_of_Month_Year($p1, $p2, $p4, $p3);\n"
Else
Let sCode = sCode + "($y, $m, $d) = ($p1, $p2, $p4);\n"
EndIf
Let sCode = sCode + "Date_to_Text_Long($y, $m, $d);\n"
Let sReturn = PerlEval(sCode, sYear, sMonth, sWeek, sDay)
Let sReturn = StringReplaceCase(sReturn, "st", ",")
Let sReturn = StringReplaceCase(sReturn, "nd", ",")
Let sReturn = StringReplaceCase(sReturn, "rd", ",")
Let sReturn = StringReplaceCase(sReturn, "th", ",")
Return sReturn
EndFunction

String Function PerlDialogAppendFile(String sFile)
Var
Handle h,
String sReturn, String sCode

If !PerlInit() Then
Return DialogSaveFile(sFile)
EndIf

Let h = GetFocus()
Let sCode = "Win32::GUI::GetSaveFileName(-file => $p1, -overwriteprompt => 0);"
ScheduleFunction("DialogSaveFileHelper", 0)
Let sReturn = oHomerPerl.Eval(sCode, sFile, "", "", "")
Pause()
SetFocus(h)
Return sReturn
EndFunction

String Function PerlDialogBrowseForFolder(String sDir)
Var
Handle h,
String sReturn, String sCode

If !PerlInit() Then
Return DialogBrowseForFolder(sDir)
EndIf

Let h = GetFocus()
;Let sCode = "Win32::GUI::BrowseForFolder(-root => $p1, -editbox => 1);"
PathSetCurrentDirectory(sDir)
Let sCode = "Win32::GUI::BrowseForFolder(-folderonly => 1, -editbox => 1);"
ScheduleFunction("DialogBrowseForFolderHelper", 0)
Let sReturn = oHomerPerl.Eval(sCode, sDir, "", "", "")
Pause()
SetFocus(h)
Return sReturn
EndFunction

String Function PerlDialogMultiInput(String sTitle, string sFields, String sValues)
;Get input from multiple edit boxes using PerlEval.dll
Var
Handle h,
Int iCount,
String sCode, String sPl, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
If !PerlInit() Then
Return IniFormDialogMultiInput(sTitle, sFields, sValues)
EndIf

Let h = GetFocus()
Let iCount = StringCountSegment(sFields, "\7")
If stringIsBlank(sTitle) Then
If iCount == 1 Then
Let sTitle = "Input"
Else
Let sTitle = "Fields"
EndIf
EndIf
;SayString(sTitle)

If StringIsBlank(sFields) && iCount == 1 Then
Let sFields = "Text"
EndIf

Let sPl = GetJAWSSettingsDirectory() + "\\Homer\\DialogMultiInput.pl"
Let sCode = FileToString(sPl)
;Let sValues = RegExpReplaceCase(sValues, "\7", "\t")
;Let sValues = StringReplaceCase(sValues, "\7", "\t")
;ScheduleFunction("PerlDialogHelper", 0)
Let sReturn = PerlEval(sCode, sTitle, sFields, sValues, "")

If h Then
SetFocus(h)
EndIf

If StringIsBlank(sReturn) Then
Return
EndIf

;Let sReturn = RegExpReplaceCase(sReturn, "\t", "\7")
;Let sReturn = StringReplaceCase(sReturn, "\t", "\7")
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
return sReturn
EndFunction

Void Function PerlDialogInfo(String sTitle, string sText)
;Display information in a multiline edit box using PerlEval.dll
Var
Handle h,
Int iDefault, 
String sCode, String sPl, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
If !PerlInit() Then
Return IniFormDialogInfo(sTitle, sText)
EndIf

Let h = GetFocus()
If stringIsBlank(sTitle) Then
Let sTitle = "Info"
EndIf
SayString(sTitle)

Let sPl = GetJAWSSettingsDirectory() + "\\Homer\\DialogInfo.pl"
Let sCode = FileToString(sPl)
;ScheduleFunction("PerlDialogHelper", 0)
Let sReturn = PerlEval(sCode, sTitle, sText, "", "")

If h Then
SetFocus(h)
EndIf

If StringIsBlank(sReturn) Then
Return
EndIf

IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
return sReturn
EndFunction

String Function PerlDialogInput(String sTitle, string sField, String sValue)
Return PerlDialogMultiInput(sTitle, sField, sValue)
EndFunction

String Function PerlDialogPick(String sTitle, String sValues, Int iSort)
;Get choices from a single selection list box using PerlEval.dll
Var
Handle h,
Int iDefault, 
String sCode, String sPl, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
If !PerlInit() Then
Return IniFormDialogPick(sTitle, sValues, iSort)
EndIf

Let h = GetFocus()
If stringIsBlank(sTitle) Then
Let sTitle = "Pick"
EndIf
Let sHomerTitle = sTitle

Let sPl = GetJAWSSettingsDirectory() + "\\Homer\\DialogPick.pl"
Let sCode = FileToString(sPl)
;ScheduleFunction("PerlDialogHelper", 0)
Let sValues = RegExpReplaceCase(sValues, "\7", "|")
Let sReturn = PerlEval(sCode, sTitle, sValues, IntToString(iSort), IntToString(iDefault))
If h Then
SetFocus(h)
EndIf

If StringIsBlank(sReturn) Then
Return
EndIf

IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
return sReturn
EndFunction

String Function PerlDialogMemo(String sTitle, string sText)
;Get input from a multiline edit box using PerlEval.dll
Var
Handle h,
Int iDefault, 
String sCode, String sPl, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
If !PerlInit() Then
Return IniFormDialogMemo(sTitle, sText)
EndIf

Let h = GetFocus()
If stringIsBlank(sTitle) Then
Let sTitle = "Memo"
EndIf
SayString(sTitle)

Let sPl = GetJAWSSettingsDirectory() + "\\Homer\\DialogMemo.pl"
Let sCode = FileToString(sPl)
;ScheduleFunction("PerlDialogHelper", 0)
Let sReturn = PerlEval(sCode, sTitle, sText, "", "")

If h Then
SetFocus(h)
EndIf

If StringIsBlank(sReturn) Then
Return
EndIf

IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
return sReturn
EndFunction

String Function PerlDialogMultiPick(String sTitle, String sValues, Int iSort)
;Get choices from a multiple selection Multi box using PerlEval.dll
Var
Handle h,
Int iDefault, 
String sCode, String sPl, String sReturn

IniWriteInteger("Internal", "SuppressTitle", True, "Homer.ini")
If !PerlInit() Then
Return IniFormDialogMultiPick(sTitle, sValues, iSort)
EndIf

Let h = GetFocus()
If stringIsBlank(sTitle) Then
Let sTitle = "Multi Pick"
EndIf
Let sHomerTitle = sTitle

Let sPl = GetJAWSSettingsDirectory() + "\\Homer\\DialogMultiPick.pl"
Let sCode = FileToString(sPl)
Let sValues = RegExpReplaceCase(sValues, "\7", "|")
;ScheduleFunction("PerlDialogHelper", 0)
Let sReturn = PerlEval(sCode, sTitle, sValues, IntToString(iSort), IntToString(iDefault))
If h Then
SetFocus(h)
EndIf

If StringIsBlank(sReturn) Then
Return
EndIf

Let sReturn = RegExpReplaceCase(sReturn, "\\|", "\7")
IniWriteInteger("Internal", "SuppressTitle", False, "Homer.ini")
return sReturn
EndFunction

String Function PerlDialogOpenFile(String sDir)
Var
Handle h,
String sReturn, String sCode

If !PerlInit() Then
Return DialogOpenFile(sDir)
EndIf

Let h = GetFocus()
Let sCode = "Win32::GUI::GetOpenFileName(-directory => $p1, -filemustexist => 1, -pathmustexist => 1);"
ScheduleFunction("DialogOpenFileHelper", 0)
Let sReturn = oHomerPerl.Eval(sCode, sDir, "", "", "")
Pause()
SetFocus(h)
Return sReturn
EndFunction

String Function PerlDialogSaveFile(String sFile)
Var
Handle h,
String sReturn, String sCode

If !PerlInit() Then
Return DialogSaveFile(sFile)
EndIf

Let h = GetFocus()
Let sCode = "Win32::GUI::GetSaveFileName(-file => $p1);"
ScheduleFunction("DialogSaveFileHelper", 0)
Let sReturn = oHomerPerl.Eval(sCode, sFile, "", "", "")
Pause()
SetFocus(h)
Return sReturn
EndFunction

String Function PerlEditGetText(Handle h)
;Get all text using PerlEval.dll
Var
String sCode, String sReturn

If !PerlInit() Then
Return EditGetText(h)
EndIf

If !h Then
Let h = GetFocus()
EndIf

Let sCode = "Win32::GuiTest::WMGetText($p1)"
Let sReturn = PerlEval(sCode, IntToString(h), "", "", "")
Return sReturn
EndFunction

Int Function PerlEditSetText(Handle h, String sText)
;Set all text using PerlEval.dll
Var
Int iReturn,
String sCode

If !PerlInit() Then
Return EditSetText(h, sText)
EndIf

If !h Then
Let h = GetFocus()
EndIf

Let sCode = "Win32::GuiTest::WMSetText($p1, $p2)"
Let iReturn = StringToInt(PerlEval(sCode, IntToString(h), sText, "", ""))
Return iReturn
EndFunction

String Function PerlEval(String sCode, String s1, String s2, String s3, String s4)
;Evaluate with PerlEval object, passing four string parameters and returning a string result
Var
Int iCreate,
Object oNull,
String sReturn

If !PerlInit() Then
Return
EndIf

;Let sCode = StringTrim(sCode)
if 0 then
;If StringTrailCase(sCode, ";") Then
Let sCode = StringChopRight(sCode, 1)
EndIf
Let sReturn = oHomerPerl.Eval(sCode, s1, s2, s3, s4)
Return sReturn
EndFunction

String Function PerlGetWMIProperty(String sObject, String sProperty)
;Get property of the first instance of WMI object
Var
String sCode, String sPl, String sReturn

Let sPl = GetJAWSSettingsDirectory() + "\\Homer\\GetWMIProperty.pl"
Let sCode = FileToString(sPl)
Let sReturn = PerlEval(sCode, sObject, sProperty, "", "")
return sReturn
EndFunction

Int Function PerlInit()
;Try to initiate global Perl object if it does not exist, and then test whether it exists
Var
Handle h

If iPerlInitialized == -1 Then
Return False
ElIf oHomerPerl Then
Return True
EndIf

ReloadAllConfigs()
If !FileExists(GetJAWSSettingsDirectory() + "\\Homer\\PerlEval.dll") Then
Let iPerlInitialized = -1
Return False
EndIf

SayString("Initializing component")
Let h = GetFocus()
Let oHomerPerl = CreateObjectEx("Perl.Eval", False)
If h Then
SetFocus(h)
EndIf

If oHomerPerl Then
SayString("Done!")
Let iPerlInitialized = 1
Return True
Else
SayString("Error!")
Let iPerlInitialized = -1
Return False
EndIf
EndFunction

String Function PerlListGetContents(Handle h)
;Get contents of item with focus in a list box using PerlEval.dll
Var
String sCode, String sReturn

If !PerlInit() Then
Return
EndIf

If !h Then
Let h = GetFocus()
EndIf

Let sCode = "join '\7', Win32::GuiTest::GetListContents($p1)"
Let sReturn = PerlEval(sCode, IntToString(h), "", "", "")
Return sReturn
EndFunction

String Function PerlListGetText(Handle h)
;Get text of item with focus in a list box using PerlEval.dll
Var
String sCode, String sReturn

If !PerlInit() Then
Return
EndIf

If !h Then
Let h = GetFocus()
EndIf

Let sCode = "Win32::GuiTest::GetListText($p1)"
Let sReturn = PerlEval(sCode, IntToString(h), "", "", "")
Return sReturn
EndFunction

Void Function PerlRunExampleCode(String sTask, String sCode, String s1, String s2, String s3, String s4)
;Evaluate a string and show the code using PerlEval object
Var
String sParams, String sResult

Let sResult = PerlEval(sCode, s1, s2, s3, s4)
If s4 Then Let sParams = "p4=" + s4 +"\r\n" + sParams EndIf
If s3 Then Let sParams = "p3=" + s3 +"\r\n" + sParams EndIf
If s2 Then Let sParams = "p2=" + s2 +"\r\n" + sParams EndIf
If s1 Then Let sParams = "p1=" + s1 +"\r\n" + sParams EndIf

If sParams then Let sParams = StringChopRight(sParams, 2) EndIf
If StringContains(sParams, "\r\n") Then Let sParams = "\r\n" + sParams EndIf
If sParams Then Let sParams = "Parameters: " + sParams + "\r\n\r\n" EndIf

Let sCode = StringReplaceAllCase(sCode, ";", ";\r\n")
If sCode then Let sCode = StringChopRight(sCode, 2) EndIf
If StringContains(sCode, "\r\n") Then Let sCode = "\r\n" + sCode EndIf
If sCode Then Let sCode = "Code: " + sCode + "\r\n\r\n" EndIf

Let sResult = StringReplaceAllCase(sResult, "\r\n", "\n")
Let sResult = StringReplaceAllCase(sResult, "\n", "\r\n")
Let sResult = StringReplaceAllCase(sResult, "\r\n\r\n", "\r\n")
If StringContains(sResult, "\r\n") Then Let sResult = "\r\n" + sResult EndIf
Let sResult = "Result: " + sResult
Let sResult = sParams + sCode + sResult

Let sTask = "Task: " + sTask
PerlDialogInfo(sTask, sResult)
EndFunction

Void Function PerlRunExampleFile(String sTask, String sFile, String s1, String s2, String s3, String s4)
;Evaluate a file and show the code using PerlEval object
Var
String sCode

Let sCode = FileToString(sFile)
PerlRunExampleCode(sTask, sCode, s1, s2, s3, s4)
EndFunction

Int Function PerlShellUrlToFile(String sUrl, String sFile)
;Download file using PerlEval.dll
Var
Int iReturn,
String sCode

If !FileDelete(sFile) Then
Return
EndIf

Let sCode = "$UrlToFile = new Win32::API('urlmon.dll', 'URLDownloadToFileA', 'NPPNN', 'N');"
;$Url = " " x 260;
;$File = " " x 260;
Let sCode = sCode + "$UrlToFile->Call(0, $p1, $p2, 0, 0);"
Let iReturn = StringToInt(PerlEval(sCode, sUrl, sFile, "", ""))
Let iReturn = FileExists(sFile)
Return iReturn
EndFunction

Handle Function PerlWindowFind(Handle h, String sTitle, String sClass, Int iID)
;Find first window starting from h and matching a RegExp Title, RegExp Class, and/or ID
Var
Handle hReturn,
String sCode

If !PerlInit() Then
Return
EndIf

Let sCode = "(Win32::GuiTest::FindWindowLike($p1, $p2, $p3, $p4))[0]"
Let hReturn = StringToHandle(PerlEval(sCode, IntToString(h), sTitle, sClass, IntToString(iID)))
Return hReturn
EndFunction

Handle Function PerlWindowGetActive()
;Get active window using PerlEval.dll
Var
Handle hReturn,
String sCode

If !PerlInit() Then
Return
EndIf

Let sCode = "Win32::GuiTest::GetActiveWindow()"
Let hReturn = StringToHandle(PerlEval(sCode, "", "", "", ""))
Return hReturn
EndFunction

Handle Function PerlWindowGetDeskTop()
;Get desktop window using PerlEval.dll
Var
Handle hReturn,
String sCode

If !PerlInit() Then
Return
EndIf

Let sCode = "Win32::GuiTest::GetDesktopWindow()"
Let hReturn = StringToHandle(PerlEval(sCode, "", "", "", ""))
Return hReturn
EndFunction

Handle Function PerlWindowGetForeground()
;Get foreground window using PerlEval.dll
Var
Handle hReturn,
String sCode

If !PerlInit() Then
Return
EndIf

Let sCode = "Win32::GuiTest::GetForegroundWindow()"
Let hReturn = StringToHandle(PerlEval(sCode, "", "", "", ""))
Return hReturn
EndFunction

Int Function PerlWindowSetActive(Handle h)
;Set active window using PerlEval.dll
Var
Int iReturn,
String sCode

If !PerlInit() Then
Return
EndIf

Let sCode = "Win32::GuiTest::SetActiveWindow($p1)"
Let iReturn = StringToInt(PerlEval(sCode, IntToString(h), "", "", ""))
Return iReturn
EndFunction

Int Function PerlWindowSetForeground(Handle h)
;Set foreground window using PerlEval.dll
Var
Int iReturn,
String sCode

If !PerlInit() Then
Return
EndIf

Let sCode = "Win32::GuiTest::SetForegroundWindow($p1)"
Let iReturn = StringToInt(PerlEval(sCode, IntToString(h), "", "", ""))
Return iReturn
EndFunction

Int Function PerlWindowSetState(Handle h, Int iState)
;Set a window state using PerlEval.dll
Var
Int iReturn,
String sCode

If !PerlInit() Then
Return WindowSetState(h, iState)
EndIf

Let sCode = "Win32::GuiTest::ShowWindow($p1, $p2)"
Let iReturn = StringToInt(PerlEval(sCode, IntToString(h), IntToString(iState), "", ""))
Return iReturn
EndFunction

Variant Function PyEval(String sCode, Variant v1, Variant v2, Variant v3, Variant v4)
Var
Variant vReturn

If Not SayToolsInit() Then
Return
EndIf

Let vReturn = oHomerST.Eval(sCode, v1, v2, v3, v4)
Return vReturn
EndFunction

Variant Function PyEvalFile(String sFile, Variant v1, Variant v2, Variant v3, Variant v4)
Var
String sCode,
Variant vReturn

If !StringContains(sFile, "\\") Then
Let sFile = GetJAWSSettingsDirectory() + "\\Homer\\" + sFile
EndIf

Let sCode = FileToString(sFile)
Let vReturn = PyEval(sCode, v1, v2, v3, v4)
Return vReturn
EndFunction

String Function RegExpContainsCase(String sText, String sMatch)
;Case sensitive Version of RegExpContainsEx with common defaults
Return RegExpContainsEx(sText, sMatch, True, "\7")
EndFunction

String Function RegExpContainsEquiv(String sText, String sMatch)
;Case equivalent Version of RegExpContainsEx with common defaults
Return RegExpContainsEx(sText, sMatch, False, "\7")
EndFunction

String Function RegExpContainsEx(String sText, String sMatch, Int iCaseSensitive, String sDelimiter)
;Get starting index and text of the first match of a regular expression
;where sText is the string to search,
;sMatch is the regular expression to match,
;iCaseSensitive indicates whether capitalization matters,
Var
Int iIndex, Int i, Int iCount,
Object oExp, Object oMatches, Object oMatch, Object oNull,
String sValue, String sReturn

Let oExp =ObjectCreate("VBScript.RegExp")
Let oExp.Pattern =sMatch
Let oExp.Ignorecase = !iCaseSensitive
Let oExp.Multiline = False
Let oExp.Global = False
Let oMatches = oExp.Execute(sText)
Let iCount = oMatches.Count
If iCount Then
Let oMatch = oMatches.Item(0)
Let iIndex = oMatch.FirstIndex + 1
Let sValue = oMatch.Value
Let sReturn = IntToString(iIndex) + sDelimiter + sValue
EndIf

Let oMatch = oNull
Let oMatches = oNull
Let oExp = oNull
Return sReturn
EndFunction

String Function RegExpContainsLastCase(String sText, String sMatch)
;Case sensitive version of RegExpContainsLastEx with common defaults
Return RegExpContainsLastEx(sText, sMatch, True, "\7")
EndFunction

String Function RegExpContainsLastEquiv(String sText, String sMatch)
;Case equivalent version of RegExpContainsLastEx with common defaults
Return RegExpContainsLastEx(sText, sMatch, False, "\7")
EndFunction

String Function RegExpContainsLastEx(String sText, String sMatch, Int iCaseSensitive, String sDelimiter)
;Get starting index and text of the last match of a regular expression
;where sText is the string to search,
;sMatch is the regular expression to match,
;iCaseSensitive indicates whether capitalization matters,
Var
Int iIndex, Int i, Int iCount,
Object oExp, Object oMatches, Object oMatch, Object oNull,
String sValue, String sReturn

Let oExp =ObjectCreate("VBScript.RegExp")
Let oExp.Pattern =sMatch
Let oExp.Ignorecase = !iCaseSensitive
Let oExp.Multiline = False
Let oExp.Global = True
Let oMatches = oExp.Execute(sText)
Let iCount = oMatches.Count
If iCount Then
Let oMatch = oMatches.Item(iCount - 1)
Let iIndex = oMatch.FirstIndex + 1
Let sValue = oMatch.Value
Let sReturn = IntToString(iIndex) + sDelimiter + sValue
EndIf

Let oMatch = oNull
Let oMatches = oNull
Let oExp = oNull
Return sReturn
EndFunction

Int Function RegExpCountCase(String sText, String sMatch)
;Case sensitive version of RegExpCountEx
Return RegExpCountEx(sText, sMatch, True)
EndFunction

Int Function RegExpCountEquiv(String sText, String sMatch)
;Case equivalent version of RegExpCountEx
Return RegExpCountEx(sText, sMatch, False)
EndFunction

Int Function RegExpCountEx(String sText, String sMatch, Int iCaseSensitive)
;Count matches of a regular expression
;where sText is the string to search,
;sMatch is the regular expression to match,
;iCaseSensitive indicates whether capitalization matters,
Var
Int iReturn, Int i, Int iCount,
Object oExp, Object oMatches, Object oMatch, Object oNull,
String sValue, String sReturn

Let oExp =ObjectCreate("VBScript.RegExp")
Let oExp.Pattern =sMatch
Let oExp.Ignorecase = !iCaseSensitive
Let oExp.Multiline = False
Let oExp.Global = True
Let oMatches = oExp.Execute(sText)
Let iReturn = oMatches.Count

Let oMatches = oNull
Let oExp = oNull
Return iReturn
EndFunction

String Function RegExpExtractCase(String sText, String sMatch)
;Case sensitive version of RegExpExtractEx with common defaults
Return RegExpExtractEx(sText, sMatch, True, "\7")
EndFunction

String Function RegExpExtractEquiv(String sText, String sMatch)
;Case equivalent version of RegExpExtractEx with common defaults
Return RegExpExtractEx(sText, sMatch, False, "\7")
EndFunction

String Function RegExpExtractEx(String sText, String sMatch, Int iCaseSensitive, String sDelimiter)
;Get matches of a regular expression
;where sText is the string to search,
;sMatch is the regular expression to match,
;iCaseSensitive indicates whether capitalization matters,
;sDelimiter seperates matches
Var
Int i, Int iCount,
Object oExp, Object oMatches, Object oMatch, Object oNull,
String sValue, String sReturn

Let oExp =ObjectCreate("VBScript.RegExp")
Let oExp.Pattern =sMatch
Let oExp.Ignorecase = !iCaseSensitive
Let oExp.Multiline = False
Let oExp.Global = True
Let oMatches = oExp.Execute(sText)
Let iCount = oMatches.Count
Let i = 0
While i < iCount
Let oMatch = oMatches.Item(i)
Let sValue = oMatch.Value
Let sReturn = sReturn + sValue + sDelimiter
Let i = i + 1
EndWhile
Let sReturn = StringChopRight(sReturn, 1)

Let oMatch = oNull
Let oMatches = oNull
Let oExp = oNull
Return sReturn
EndFunction

String Function RegExpReplaceCase(String sText, String sMatch, String sReplace)
;Case sensitive version of RegExpReplaceEx
Return RegExpReplaceEx(sText, sMatch, sReplace, True, True)
EndFunction

String Function RegExpReplaceEquiv(String sText, String sMatch, String sReplace)
;Case equivalent version of RegExpReplaceEx
Return RegExpReplaceEx(sText, sMatch, sReplace, False, True)
EndFunction

String Function RegExpReplaceEx(String sText, String sMatch, String sReplace, Int iCaseSensitive, Int iGlobal)
;Replace text matching a regular expression
;where sText is the string to search,
;sMatch is the regular expression to match,
;sReplace is the replacement text,
;iCaseSensitive indicates whether capitalization matters,
;iGlobal indicates whether to replace the first or all matches

Var
Object oExp, Object oNull,
String sReturn

Let oExp =ObjectCreate("VBScript.RegExp")
Let oExp.Pattern =sMatch
Let oExp.Ignorecase = !iCaseSensitive
Let oExp.Multiline = False
Let oExp.Global =iGlobal
Let sReturn =oExp.Replace(sText, sReplace)

Let oExp = oNull
Return sReturn
EndFunction

String Function RegistryRead(String sKey)
;Get a string from the registry

Var
Object oShell, Object oNull,
String sReturn

Let oShell =ObjectCreate("Wscript.Shell")
Let sReturn =oShell.RegRead(sKey)
Let sReturn =oShell.ExpandEnvironmentStrings(sReturn)

Let oShell = oNull
Return sReturn
EndFunction

Int Function RegistryWrite(String sKey, String sValue)
;Write a string to the registry

Var
Int iReturn,
Object oShell, Object oNull,
String sReturn

Let oShell =ObjectCreate("Wscript.Shell")
Let iReturn =oShell.RegWrite(sKey, sValue, "REG_SZ")

Let oShell = oNull
Return iReturn
EndFunction

Void Function SayFileByLine(String sFile, Int iRepeat)
;Say a text file a line at a time, with repeat spelling support, silenced by any key

Var
Int iBlank,
Object oSystem, Object oFile, Object oNull,
String sLine

Let oSystem =ObjectCreate("Scripting.FileSystemObject")
Let oFile =oSystem.OpenTextFile(sFile)
While !oFile.AtEndOfStream && !IsKeyWaiting()
Let sLine =oFile.ReadLine()
If !StringIsBlank(sLine) Then
SayRepeat(sLine, iRepeat)
EndIf
EndWhile
oFile.Close()

Let oFile = oNull
Let oSystem = oNull
EndFunction

Void Function SayIfVerbose(String s)
;Say a string unless verbosity is advanced
If GetVerbosity() < 2 Then SayString(s) EndIf
EndFunction

Void Function SayRepeat(String sText, Int iRepeat)
;Say a string, or spell if repeated, or phonetic if triple key presses

If iRepeat >=2 Then
SpellPhonetic(sText)
ElIf iRepeat Then
SpellString(sText)
Else
SayString(sText)
EndIf
EndFunction

Void Function SayStringByLine(String sText, Int iRepeat)
Var
Int iCount, Int i,
String sLine
Let iCount = StringCountSegment(sText, "\n")
Let i =1
While (i <= iCount) && !IsKeyWaiting()
Let sLine =StringSegment(sText, "\n", i)
If !StringIsBlank(sLine) Then
SayRepeat(sLine, iRepeat)
EndIf
Let i =i +1
EndWhile
EndFunction

Void Function SayStringIf(Int iCondition, String sTrue, String sFalse)
;Say one of two strings, depending on condition

If iCondition Then
SayString(sTrue)
Else
SayString(sFalse)
EndIf
EndFunction

Void Function SayStringIfObject(Object o, String s1, String s2)
;Say one of two strings, depending on whether an object exists
If o Then
SayString(s1)
Else
SayString(s2)
Endif
EndFunction

Void Function SayTempFile()
Var
String sFile, String sText

Let sFile = GetJAWSSettingsDirectory() + "\\" + GetActiveConfiguration() + ".tmp"
Let sText = FileToString(sFile)
SayString(sText)
EndFunction

Int Function SayToolsInit()
;Try to initiate global SayTools object if it does not exist, and then test whether it exists

If iSayToolsInitialized == -1 Then
Return False
ElIf oHomerST Then
Return True
EndIf

;Let oHomerST = CreateObjectEx("Say.Tools", False)
Let oHomerST = CreateObjectEx("Say.Tools", True)
If oHomerST Then
Let iSayToolsInitialized = 1
Return True
Else
SayString("Error initializing SayTools component!")
Let iSayToolsInitialized = -1
Return False
EndIf
EndFunction

Int Function SayVirtual(String sText)
;Display text in the virtual viewer
UserBufferDeactivate()
UserBufferClear()
UserBufferAddText(sText)
UserBufferActivate()
JAWSTopOfFile()
SayAll()
EndFunction

Void Function SBox(String s)
Var
Int iSpeech

Let iSpeech = VoiceSetSpeech(True)
MessageBox(s)
VoiceSetSpeech(iSpeech)
EndFunction

Int Function ShellCreateShortcut(String sFile, String sTargetPath, String sWorkingDirectory, Int iWindowStyle, String sHotkey)
;Create a .lnk or .url file
Var
Object oShell, Object oShortcut, Object oNull

Let oShell = ObjectCreate("WScript.Shell")
Let oShortcut = oShell.CreateShortcut(sFile)
Let oShortcut.TargetPath = sTargetPath
Let oShortcut.WorkingDirectory = sWorkingDirectory
Let oShortcut.WindowStyle = iWindowStyle
Let oShortcut.Hotkey = sHotkey
oShortcut.Save()

Let oShortcut = oNull
Let oShell = oNull
Return FileExists(sFile)
EndFunction

String Function ShellExec(String sCommand)
;Run a console mode command and return its standard output

Var
Object oShell, Object oExec, Object oOutput, Object oNull,
String sReturn

Let oShell =CreateObjectEx("Wscript.Shell", False)
SpeechOff()
Let oExec =oShell.Exec(sCommand)
While oExec.Status ==0
Delay(1)
EndWhile

Let oOutput =oExec.StdOut
Let sReturn =oOutput.ReadAll()
oExec.Terminate()
SpeechOn()

Let oOutput = oNull
Let oExec = oNull
Let oShell = oNull
Return sReturn
EndFunction

String Function ShellGetDrives()
Var
Int i,
String sReturn, String sDrive, String sDrives,
Object oSystem, Object oDrive, Object oNull

Let oSystem = ObjectCreate("Scripting.FileSystemObject")
Let sDrives = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Let i = 1
While i <= 26
Let sDrive = Substring(sDrives, i, 1)
If oSystem.DriveExists(sDrive) Then
Let oDrive = oSystem.GetDrive(sDrive)
If oDrive.IsReady Then
Let sReturn = sReturn + sDrive
EndIf
Let oDrive = oNull
EndIf
Let i = i + 1
EndWhile
Let oSystem = oNull
Return sReturn
EndFunction

String Function ShellGetEnvironmentVariable(String sVariable)
;Get the value of an environment variable

Var
Object oShell, Object oEnv, Object oNull,
String sReturn

Let oShell =ObjectCreate("Wscript.Shell")
Let oEnv =oShell.Environment
Let sReturn =oEnv.Item(sVariable)

Let oEnv = oNull
Let oShell = oNull
Return sReturn
EndFunction

String Function ShellGetShortcutTargetPath(String sFile)
;Get the target path of a shortcut file
Var
Object oShell, Object oShortcut, Object oNull,
String sReturn

Let oShell = ObjectCreate("WScript.Shell")
Let oShortcut = oShell.CreateShortcut(sFile)
Let sReturn = oShortcut.TargetPath
Return sReturn
EndFunction

String Function ShellGetWindowsName()
;Get name of Windows installed
Return GetVersionInfoString(GetWindowsSystemDirectory() + "\\win.com", "ProductName")
EndFunction

String Function ShellGetWindowsNTName()
;Get name of the Windows NT version installed
Var
Int iKey,
String sSubkey, String sValueName

Let iKey = 2 ; HKEY_LOCAL_MACHINE
Let sSubkey = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion"
Let sValueName = "ProductName"

If StringLeft(IntToString(GetJFWVersion()), 1) >= "6" Then
Return GetRegistryEntryString(iKey, sSubkey, sValueName)
Else
Return RegistryRead("HKEY_LOCAL_MACHINE\\" + sSubkey + "\\" + sValueName)
EndIf
EndFunction

Int Function ShellRun(String sFile, Int iStyle, Int iWait)
;Launch a program or file, indicating its window style and whether to wait before returning
;window styles:
;0 Hides the window and activates another window.
;1 Activates and displays a window. If the window is minimized or maximized, the
;system restores it to its original size and position. This flag should be used
;when specifying an application for the first time.
;2 Activates the window and displays it minimized.
;3 Activates the window and displays it maximized.
;4 Displays a window in its most recent size and position. The active window
;remains active.
;5 Activates the window and displays it in its current size and position.
;6 Minimizes the specified window and activates the next top-level window in the Z
;order.
;7 Displays the window as a minimized window. The active window remains active.
;8 Displays the window in its current state. The active window remains active.
;9 Activates and displays the window. If it is minimized or maximized, the system
;restores it to its original size and position. An application should specify
;this flag when restoring a minimized window.
;10 Sets the show state based on the state of the program that started the
;application.

Var
Int iReturn,
Object oShell, Object oNull

Let oShell =CreateObjectEx("Wscript.Shell", False)
Let iReturn =oShell.Run(sFile, iStyle, iWait)

Let oShell = oNull
Return iReturn
EndFunction

Int Function ShellRunCommandLine(String sDir)
;Open a command prompt in the directory specified
ShellRun("%COMSPEC% /k cd \34" + sDir + "\34", 1, False)
EndFunction

Int Function ShellRunExplorer(String sDir)
;Open Windows Explorer in the directory specified
Run("\34" + sDir + "\34")
EndFunction

Int Function ShellSetHomerDir()
Var
Int iKey,
String sSubkey, String sValueName, String sValueData

Let iKey = 2 ; HKEY_LOCAL_MACHINE
Let sSubkey = "SOFTWARE\\EmpowermentZone\\Homer\\CurrentVersion"
Let sValueName = "UtilDir"
Let sValueData = GetJAWSSettingsDirectory() + "\\Homer"

Return RegistryWrite("HKEY_LOCAL_MACHINE\\" + sSubkey + "\\" + sValueName, sValueData)
EndFunction

Int Function ShellUrlToFile(String sURL, String sFile)
;Download file from Internet to disk
Var
Int iStyle, Int iWait,
Int adTypeBinary, Int adSaveCreateOverWrite, Int adSaveCreateNotExist,
String sText, String sVBS, String sCommand

If !FileDelete(sFile) Then
Return False
EndIf

Let iStyle = 0
Let iWait = True
Let sText = sText + "sURL = \34" + sURL + "\34\n"
Let sText = sText + "sFile = \34" + sFile + "\34\n"
Let sText = sText + "adTypeBinary = 1\n"
Let sText = sText + "Const adSaveCreateOverWrite = 2\n"
Let sText = sText + "Const adSaveCreateNotExist = 1\n"
;Let sText = sText + "oXML = CreateObject(\34Microsoft.XMLHTTP\34)\n"
Let sText = sText + "Set oXML = CreateObject(\34MSXML2.XMLHTTP\34)\n"
Let sText = sText + "Call oXML.Open(\34GET\34, sURL, False)\n"
Let sText = sText + "oXML.Send()\n"
Let sText = sText + "Set oStream = CreateObject(\34Adodb.Stream\34)\n"
Let sText = sText + "oStream.Type = adTypeBinary\n"
Let sText = sText + "oStream.open()\n"
Let sText = sText + "oStream.Write(oXML.ResponseBody)\n"
;Let sText = sText + " Do not overwrite an existing file\n"
;Let sText = sText + "Call oStream.SaveToFile(sFile, adSaveCreateNotExist)\n"
;Let sText = sText + " Use this form to overwrite a file if it already exists\n"
Let sText = sText + "Call oStream.SaveToFile(sFile, adSaveCreateOverWrite)\n"
Let sText = sText + "oStream.Close()\n"
Let sText = sText + "Set oStream = Nothing\n"
Let sText = sText + "Set oXML = Nothing\n"
;Let sVBS = PathGetTempFolder() + "\\temp.vbs"
Let sVBS = PathGetTempFile()
;Let sVBS = GetJAWSSettingsDirectory() + "\\kitsetup.tmp"
StringToFile(sText, sVBS)
;Let sCommand = "cscript.exe " + sVBS
Let sCommand = "cscript.exe /e:vbscript " + StringQuote(sVBS)
;CopyToClipboard(sText + "\n"+ sCommand)
ShellRun(sCommand, iStyle, iWait)
FileDelete(sVBS)
Return FileExists(sFile)
EndFunction

Void Function SpellANSI(String sText)
;Spell string using ANSI codes

Var
String s, String sAlphaList, String sPhoneticList,
Int i, Int iIndex, Int iLength

Let iLength =StringLength(sText)
Pause()
Let i =1
While i <=iLength && !IsKeyWaiting()
Let s = Substring(sText, i, 1)
Let iIndex = GetCharacterValue(s)
SayInteger(iIndex)
delay(2)
Let i =i +1
EndWhile
EndFunction

Void Function SpellPhonetic(String sText)
;Spell string using phonetic alphabet

Var
String s, String sAlphaList, String sPhoneticList,
Int i, Int iIndex, Int iLength

Let sPhoneticList ="space|alpha|bravo|charlie|delta|echo|foxtrot|golf|hotel|india|juliette|kilo|lema"
Let sPhoneticList =sPhoneticList +"|mike|november|oscar|poppa|Quebec|romeo|sierra|tango|uniform|victor|Whiskey|Xray|yankee|zulu"
Let sAlphaList =" abcdefghijklmnopqrstuvwxyz"

Let iLength =StringLength(sText)
Let i =1
While i <=iLength && !IsKeyWaiting()
Let s =Substring(sText, i, 1)
Let iIndex =StringContains(sAlphaList, StringLower(s))
If iIndex Then
Let s =StringSegment(sPhoneticList, "|", iIndex)
EndIf
SayString(s)
delay(2)
Let i =i +1
EndWhile
EndFunction

Void Function SSay(String s)
;Say a string regardless of speech on/off status (used in debugging)

If IsSpeechOff() Then
SpeechOn()
SayString(s)
SpeechOff()
Else
SayString(s)
EndIf
EndFunction

String Function StatusBarGetText()
; Get status line text of main application window

Var
Handle h,
Int iRestriction,
String sReturn

Let h = FindWindow(GetAppMainWindow(GetFocus()), "msctls_statusbar32")
If h Then
Let sReturn = GetWindowText(h, READ_EVERYTHING)
else
SaveCursor()
InvisibleCursor()
RouteInvisibleToPC()
Let iRestriction = GetRestriction()
SetRestriction(RESTRICTAPPWINDOW)
JawsPageDown()
Let sReturn =GetLine()
SetRestriction(iRestriction)
RestoreCursor()
EndIf

Return sReturn
EndFunction

String Function StringAddSegment(String sText, String sDelimiter, String sSegment)
;Add segment to a string, omitting delimiter if the first one
If StringIsBlank(sText) Then
Let sText = sSegment
Else
Let sText = sText + sDelimiter + sSegment
EndIf
Return sText
EndFunction

Int Function StringAppendToClipboard(String sText, String sDivider)
;Append string to clipboard, omitting delimiter if the first one
Var
String sClipboard

Let sClipboard = GetClipboardText()
Pause()
If !StringIsBlank(sClipboard) Then
Let sText = sClipboard + sDivider + sText
EndIf
CopyToClipboard(sText)
EndFunction

Int Function StringAppendToFile(String sText, String sFile, String sDivider)
;Append string to File, omitting divider if the first one

If FileExists(sFile) Then
Let sText = FileToString(sFile) + sDivider + sText
EndIf
StringToFile(sText, sFile)
Return FileExists(sFile)
EndFunction

String Function StringConcat(String s1, String s2)
;Concatonate two strings, working around JAWS bug with strings that look like numbers
Var
String sReturn

;Return StringChopLeft("|" + s1 + s2, 1)
Let sReturn = "|" + s1
Let sReturn = sReturn + s2
Let sReturn = StringChopLeft(sReturn, 1)
Return sReturn
EndFunction

Int Function StringContainsEquiv(String sText, String sMatch)
;Case-insensitive version of StringContains
Return StringContains(StringLower(sText), StringLower(sMatch))
EndFunction

Int Function StringContainsLast(String sText, String sMatch)
;Get starting index of last match of substring

Var
Int iReturn

Let iReturn =StringContains(StringReverse(sText), StringReverse(sMatch))
If iReturn Then
Let iReturn =StringLength(sText) - iReturn - StringLength(sMatch) + 2
EndIf

Return iReturn
EndFunction

Int Function StringContainsLastEquiv(String sText, String sMatch)
;Case equivalent version of StringContains Last

Var
Int iReturn

Let iReturn =StringContainsEquiv(StringReverse(sText), StringReverse(sMatch))
If iReturn Then
Let iReturn =StringLength(sText) - iReturn - StringLength(sMatch) + 2
EndIf

Return iReturn
EndFunction

Int Function StringCountSegment(String sText, String sDelimiter)
;Version of StringSegmentCount that works on JAWS 5.1 and above
;Return StringLength(sText) - StringLength(StringReplaceEx(sText, sDelimiter, "", True)) + 1
Return StringSegmentCount(sText, sDelimiter)
EndFunction

String Function StringDeleteSegment(String sText, String sDelimiter, Int iIndex)
;Delete segment from a string
Var
Int iCount, Int iStart, Int iEnd, Int iLength,
String sSegment, String sReturn

Let sSegment = StringSegment(sText, sDelimiter, iIndex)
Let iStart = StringSegmentStartCase(sText, sDelimiter, sSegment)
Let iLength = StringLength(sSegment)
Let iEnd = iStart + iLength
Let iCount = StringCountSegment(sText, sDelimiter)
If iIndex < iCount Then
Let iEnd = iEnd + 1
EndIf
Let sReturn = StringReplaceRange(sText, iStart, iEnd, "")
Return sReturn
EndFunction

Int Function StringEqual(String s1, String s2)
;Test whether first and second strings are the same, case sensitive

Return StringContains(s1, s2) && StringLength(s1) ==StringLength(s2)
EndFunction

Int Function StringEquiv(String s1, String s2)
;Case equivalent version of StringEqual

Return s1 ==s2 && StringLength(s1) ==StringLength(s2)
EndFunction

String Function StringGetBounds(String sText, Int iStart, Int iEnd)
;Get text between lower and upper positions, inclusive
Var
Int iLength

Let iLength = StringLength(sText)
If iStart < 0 Then
Let iStart = iLength + iStart + 1
EndIf
If iEnd < 0 Then
Let iEnd = iLength + iEnd + 1
EndIf
Return Substring(sText, iStart, iEnd - iStart + 1)
EndFunction

String Function StringGetRange(String sText, Int iStart, Int iEnd)
;Get text between lower and upper positions, excluding end point
Return StringGetBounds(sText, iStart, iEnd - 1)
EndFunction

String Function StringGetValueCharacter(Int i)
;Get a character from its ASCII value
Return Substring(sHomerASCII, i, 1)
EndFunction

String Function StringIf(Int iCondition, String sTrue, String sFalse)
;Return one of two strings, depending on condition

Var
String sReturn

If iCondition Then
Let sReturn =sTrue
Else
Let sReturn =sFalse
EndIf

Return sReturn
EndFunction

Int Function StringIndexSegmentCase(String sText, String sDelimiter, String sSegment)
;Case sensitive version of StringIndexSegmentEx
;Return StringIndexSegmentEx(sText, sDelimiter, sSegment, True)
Return StringSegmentIndex(sText, sDelimiter, sSegment, True)
EndFunction

Int Function StringIndexSegmentEquiv(String sText, String sDelimiter, String sSegment)
;Case equivalent version of StringIndexSegmentEx
Return StringIndexSegmentEx(sText, sDelimiter, sSegment, False)
EndFunction

Int Function StringIndexSegmentEx(String sText, String sDelimiter, String sSegment, Int iCaseSensitive)
;Get the index of a string segment
Var
Int i, Int iCount, Int iReturn,
String s

Let iCount = StringCountSegment(sText, sDelimiter)
Let i = 1
While i <= iCount
Let s = StringSegment(sText, sDelimiter, i)
If (iCaseSensitive && StringEqual(s, sSegment)) || (!iCaseSensitive && StringEquiv(s, sSegment)) Then
Let iReturn = i
Let i = iCount + 2
Else
Let i = i + 1
EndIf
EndWhile
Return iReturn
EndFunction

Int Function StringLeadCase(String s1, String s2)
;Case sensitive version of StringLeadEx
Return StringLeadEx(s1, s2, True)
EndFunction

Int Function StringLeadEquiv(String s1, String s2)
;Case equivalent version of StringLeadEx
;Return s1 ==s2
Return StringLeadEx(s1, s2, False)
EndFunction

Int Function StringLeadEx(String s1, String s2, Int iCaseSensitive)
;Test whether second string matches leading part of first string

If iCaseSensitive Then
Return StringEqual(StringLeft(s1, StringLength(s2)), s2)
Else
Return StringEquiv(StringLeft(s1, StringLength(s2)), s2)
Endif
EndFunction

String Function StringPadLeft(String sText, String sPadCharacter, Int iLength)
;Pad a string with characters to the left
Var
Int i,
String sPadding, String sReturn

If StringEqual(sPadCharacter, " ") Then
Let sPadding = "|" + sHomerSpace
ElIf StringEqual(sPadCharacter, "0") Then
Let sPadding = "|" + sHomerZero
Else
Let sPadding = "|"
Let i = 1
While i <= iLength
Let sPadding = sPadding + sPadCharacter
Let i = i + 1
EndWhile
EndIf

Let sReturn = StringRight(sPadding + sText, iLength)
Return sReturn
EndFunction

String Function StringPadRight(String sText, String sPadCharacter, Int iLength)
;Pad a string with characters to the right
Var
Int i,
String sPadding, String sReturn

If StringEqual(sPadCharacter, " ") Then
Let sPadding = sHomerSpace + "|"
ElIf StringEqual(sPadCharacter, "0") Then
Let sPadding = sHomerZero + "|"
Else
Let sPadding = "|"
Let i = 1
While i <= iLength
Let sPadding = sPadCharacter + sPadding
Let i = i + 1
EndWhile
EndIf

Let sReturn = StringLeft(sText + sPadding, iLength)
Return sReturn
EndFunction

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

String Function StringProper(String sText)
;Capitalize the first letter and lower case the rest
Return StringUpper(StringLeft(sText, 1)) + StringLower(Substring(sText, 2, StringLength(sText)))
EndFunction

String Function StringProperWords(String sText)
;Capitalize the first letter of each word and lower case the rest
Var
Int iWord, Int iWordCount, Int iLength,
String sWord, String sReturn

Let iWordCount = StringCountSegment(sText, " ")
Let iWord = 1
While iWord <= iWordCount
Let sWord = StringSegment(sText, " ", iWord)
Let iLength = StringLength(sWord)
Let sWord = StringUpper(StringLeft(sWord, 1)) + StringLower(Substring(sWord, 2, iLength - 1))
If iWord == iWordCount Then
Let sReturn = sReturn + sWord
Else
Let sReturn = sReturn + sWord + " "
EndIf
Let iWord = iWord + 1
EndWhile
Return sReturn
EndFunction

String Function StringQuote(String sText)
;Ensure that a string is enclosed by a single set of quotes
Return "\34" + StringUnquote(sText) + "\34"
EndFunction

String Function StringReplaceAllCase(String sText, String sMatch, String sReplace)
;Case sensitive version of StringReplaceAllEx
Return StringReplaceAllEx(sText, sMatch, sReplace, True)
EndFunction

String Function StringReplaceAllEquiv(String sText, String sMatch, String sReplace)
;Case equivalent version of StringReplaceAllEx
Return StringReplaceAllEx(sText, sMatch, sReplace, False)
EndFunction

String Function StringReplaceAllEx(String sText, String sMatch, String sReplace, Int iCaseSensitive)
;Repeat a search and replace until no matches remain
If (iCaseSensitive && StringContains(sReplace, sMatch)) || (!iCaseSensitive && StringContainsEquiv(sReplace, sMatch)) Then
Return StringReplaceEx(sText, sMatch, sReplace, iCaseSensitive)
EndIf

While (iCaseSensitive &&  StringContains(sText, sMatch)) || (!iCaseSensitive && StringContainsEquiv(sText, sMatch))
Let sText = StringReplaceEx(sText, sMatch, sReplace, iCaseSensitive)
EndWhile
Return sText
EndFunction

String Function StringReplaceBounds(String sText, Int iStart, Int iEnd, String sReplace)
;Replace text between lower and upper positions, inclusive
Var
Int iLength

Let iLength = StringLength(sText)
If iStart < 0 Then
Let iStart = iLength + iStart + 1
EndIf
If iEnd < 0 Then
Let iEnd = iLength + iEnd + 1
EndIf
Return StringLeft(sText, iStart - 1) + sReplace + StringRight(sText, iLength - iEnd)
EndFunction

String Function StringReplaceCase(String sText, String sMatch, String sReplace)
;Case sensitive version of StringReplaceEx
;Return StringReplaceEx(sText, sMatch, sReplace, True)
Return StringReplaceSubstrings(sText, sMatch, sReplace)
EndFunction

String Function StringReplaceEquiv(String sText, String sMatch, String sReplace)
;Case equivalent version of StringReplaceEx
Return StringReplaceEx(sText, sMatch, sReplace, False)
EndFunction

String Function StringReplaceEx(String sText, String sMatch, String sReplace, Int iCaseSensitive)
;Replace matching text

Var
Int iFound, Int iLoop, Int iMatch, Int iPre,
Int iReplace, Int iSource, Int iRest, Int iReturn,
String sPre, String sRest, String sReturn

Let iMatch =StringLength(sMatch)
Let sReturn =sText
Let sRest =sReturn
Let iRest =StringLength(sRest)

Let iLoop =True
While iLoop
Let iReturn =StringLength(sReturn)
Let iRest =StringLength(sRest)
Let iPre =iReturn -iRest
If iCaseSensitive Then
Let iFound =StringContains(sRest, sMatch)
Else
Let iFound =StringContains(StringLower(sRest), StringLower(sMatch))
EndIf
If iFound Then
Let sRest =Substring(sRest, (iFound +iMatch), iRest)
Let sReturn =Substring(sReturn, 1, (iPre +iFound -1)) +sReplace +sRest
Else
Let iLoop =False
EndIf
EndWhile

Return sReturn
EndFunction

String Function StringReplaceRange(String sText, Int iStart, Int iEnd, String sReplace)
;Replace text between lower and upper positions, excluding end point
Return StringReplaceBounds(sText, iStart, iEnd - 1, sReplace)
EndFunction

String Function StringReplaceSegment(String sText, String sDelimiter, Int iIndex, String sReplace)
;Replace a string segment
Var
Int iStart, Int iEnd, Int iLength,
String sSegment, String sReturn

Let sSegment = StringSegment(sText, sDelimiter, iIndex)
Let iStart = StringSegmentStartCase(sText, sDelimiter, sSegment)
Let iLength = StringLength(sSegment)
Let iEnd = iStart + iLength
Let sReturn = StringReplaceRange(sText, iStart, iEnd, sReplace)
Return sReturn
EndFunction

String Function StringReplaceTokens(String sText)
;Replace white space tokens with the actual characters they represent
Var
Int i, Int iCount,
String sReturn, String sToken, String sChar, String sTokens, String sChars

Let sReturn = sText
Let sTokens = "\\r|\\n|\\t|\\f"
Let sChars = "\r\n\t\f"
Let iCount = StringCountSegment(sTokens, "|")
Let i = 1
While i <= icount
Let sToken = StringSegment(sTokens, "|", i)
Let sChar = Substring(sChars, i, 1)
Let sReturn = StringReplaceCase(sReturn, sToken, sChar)
Let i = i + 1
EndWhile
Return sReturn
EndFunction

String Function StringSegmentFilter(String sItems, String sDelimiter, String sMatch)
;Get matches of an SQL filter expression
Var
Int iItem, Int iCount, Int iDataType, Int iMaxLength,
Object oRS, Object oFields, Object oItem, Object oNull,
String sItem, String sReturn

Let oRS = ObjectCreate("AdoDb.RecordSet")
Let iDataType = 202 ; adVarWChar
Let iMaxLength = 260 ; maximum length of a file path
Let oFields = oRS.Fields
oFields.Append("Item", iDataType, iMaxLength)
oRS.Open()

Let iCount = StringCountSegment(sItems, sDelimiter)
Let iItem = 1
While iItem <= iCount
Let sItem = StringSegment(sItems, sDelimiter, iItem)
oRS.AddNew("Item", sItem)
Let iItem = iItem + 1
EndWhile
oRS.Update()

Let oRS.Filter = "Item LIKE '" + sMatch + "'"
Let oItem = oFields("Item")
oRS.MoveFirst()
While !ORS.EOF
Let sItem = oItem.Value
Let sReturn = sReturn + sItem + sDelimiter
oRS.MoveNext()
EndWhile
Let sReturn = StringChopRight(sReturn, 1)
oRS.Close()

Let oItem = oNull
Let oFields = oNull
Let oRS = oNull
Return sReturn
EndFunction

String Function StringSegmentSort(String sItems, String sDelimiter)
;Sort string segments alphabetically
Return StringSegmentSortEx(sItems, sDelimiter, False)
EndFunction

String Function StringSegmentSortEx(String sItems, String sDelimiter, Int iReverse)
;Sort string segments alphabetically or the reverse
Var
Int iItem, Int iCount, Int iDataType, Int iMaxLength,
Object oRS, Object oFields, Object oItem, Object oNull,
String sItem, String sReturn

Let oRS = ObjectCreate("AdoDb.RecordSet")
Let iDataType = 202 ; adVarWChar
Let iMaxLength = 260 ; maximum length of a file path
Let oFields = oRS.Fields
oFields.Append("Item", iDataType, iMaxLength)
oRS.Open()

Let iCount = StringCountSegment(sItems, sDelimiter)
Let iItem = 1
While iItem <= iCount
Let sItem = StringSegment(sItems, sDelimiter, iItem)
oRS.AddNew("Item", sItem)
Let iItem = iItem + 1
EndWhile
oRS.Update()

If iReverse Then
Let oRS.Sort = "Item DESC"
else
Let oRS.Sort = "Item"
EndIf
Let oItem = oFields("Item")
oRS.MoveFirst()
While !ORS.EOF
Let sItem = oItem.Value
Let sReturn = sReturn + sItem + sDelimiter
oRS.MoveNext()
EndWhile
Let sReturn = StringChopRight(sReturn, 1)
oRS.Close()

Let oItem = oNull
Let oFields = oNull
Let oRS = oNull
Return sReturn
EndFunction

Int Function StringSegmentStartCase(String sText, String sDelimiter, String sSegment)
;Case sensitive version of StringSegmentStartEx
Return StringSegmentStartEx(sText, sDelimiter, sSegment, True)
EndFunction

Int Function StringSegmentStartEquiv(String sText, String sDelimiter, String sSegment)
;Case equivalent version of StringSegmentStartEx
Return StringSegmentStartEx(sText, sDelimiter, sSegment, False)
EndFunction

Int Function StringSegmentStartEx(String sText, String sDelimiter, String sSegment, Int iCaseSensitive)
;Get starting index of a string segment
Var
Int iReturn

If iCaseSensitive Then
Let iReturn = StringContains(sDelimiter + sText + sDelimiter, sDelimiter + sSegment + sDelimiter)
Else
Let iReturn = StringContainsEquiv(sDelimiter + sText + sDelimiter, sDelimiter + sSegment + sDelimiter)
EndIf
Return iReturn
EndFunction

String Function StringSetTokens(String sText)
;Replace white space characters with symbolic tokens
Var
Int i, Int iCount,
String sReturn, String sToken, String sChar, String sTokens, String sChars

Let sReturn = sText
Let sTokens = "\\r|\\n|\\t|\\f"
Let sChars = "\r\n\t\f"
Let iCount = StringCountSegment(sTokens, "|")
Let i = 1
While i <= icount
Let sToken = StringSegment(sTokens, "|", i)
Let sChar = Substring(sChars, i, 1)
Let sReturn = StringReplaceCase(sReturn, sChar, sToken)
Let i = i + 1
EndWhile
Return sReturn
EndFunction

String Function StringSortJkm(String sJkm)
;Sort key assignments alphabetically by script names
Var
Int iKey, Int iKeyCount, Int iSection, Int iSectionCount, Int iItem, Int iItemCount,
String sLineBreak, String sValue, String sSection, String sSections, String sKey, String sKeys, String sItem, String sItems, String sTempIni,
Int iStart, Int i, Int iCount, Int iLeft, Int iRight, Int iLine, Int iLineCount, Int iParamCount, Int iParam,
String sStart, String sScripts, String sScript, String s, String s1, String s2, String s3,
String sMatch, String sReplace, String sExp, String sTempJsd, String sJsd, String sText, String sLine, String sWord1, String sWord2, String sWord3, String sParamList, String sParam, String sOutput

If !FileExists(sJkm) Then
Return False
EndIf

FileCopy(sJkm, sJkm + ".bak")
Let sTempIni = PathGetTempFolder() + "\\temp.ini"
IniFlush(sTempIni)
Pause()
IniDeleteFile(sTempIni)
Pause()
Let sSections = IniReadSectionNames(sJkm)
Let sSections = StringSegmentSort(sSections, "|")
Let iSectionCount = StringCountSegment(sSections, "|")
Let iSection = 1
While iSection <= iSectionCount
Let sSection = StringSegment(sSections, "|", iSection)
Let sKeys = IniReadSectionKeys(sSection, sJkm)
Let iKeyCount = StringCountSegment(sKeys, "|")
Let iKey = 1
While iKey <= iKeyCount
Let sKey = StringSegment(sKeys, "|", iKey)
Let sValue = StringTrim(IniReadString(sSection, sKey, "", sJkm))
Let sItem = StringPadRight(sValue, " ", 50) + StringTrim(sKey)
Let sItems = sItems + sItem + "|"
Let iKey = iKey + 1
EndWhile
Let sItems = StringchopRight(sItems, 1)

Let sItems = StringSegmentSort(sItems, "|")
Let iItemCount = StringcountSegment(sItems, "|")
Let iItem = 1
While iItem <= iItemCount
Let sItem = StringSegment(sItems, "|", iItem)
Let sValue = StringTrim(StringLeft(sItem, 50))
Let sKey = StringChopLeft(sItem, 50)
IniWriteString(sSection, sKey, sValue, sTempIni, True)
IniFlush(sTempIni)
Let iItem = iItem + 1
EndWhile

Let iSection = iSection + 1
EndWhile
IniFlush(sTempIni)
Pause()

Let sText = FileToString(sTempIni)
;Let sText = ConvertToUnixLineBreak(sText)
Let sText = RegExpReplaceCase(sText, "(\\r|\\n)\\s*(\\r|\\n+)", "\r\n")
;Let sText = ConvertToWinLineBreak(sText)
Return sText
EndFunction

String Function StringSortJss(String sJss)
;Sort routines alphabetically, placing functions before scripts
Var
Int iStart, Int i, Int iCount, Int iLeft, Int iRight, Int iLine, Int iLineCount, Int iParamCount, Int iParam,
String sStart, String sScripts, String sScript, String sKey, String s, String s1, String s2, String s3,
String sMatch, String sReplace, String sExp, String sTempJsd, String sJsd, String sText, String sLine, String sWord1, String sWord2, String sWord3, String sParamList, String sParam, String sOutput

If !FileExists(sJss) Then
Return False
EndIf

FileCopy(sJss, sJss + ".bak")
Let sText = FileToString(sJss)
Let sExp = "\\r\\n(( |\\t|\\w)* +)*((Function)|(Script)) +\\w+"
Let s = RegExpContainsEquiv(sText, sExp)
If StringIsBlank(s) Then
Return False
EndIf

Let iStart = StringToInt(StringSegment(s, "\7", 1))
Let sStart = StringLeft(sText, iStart - 1)
Let sText = Substring(sText, iStart, StringLength(sText))
Let sMatch = "(" + sExp + ")"
Let sReplace = "\r\n\f\r\n$1"
Let sText = RegExpReplaceEquiv(sText, sMatch, sReplace) + "\r\n\f\r\n"

Let sMatch = sMatch + "[^\f]*"
Let sScripts = RegExpExtractEquiv(sText, sMatch)
Let sText = ""
Let iCount = StringCountSegment(sScripts, "\7")
Let i = 1
While i <= iCount
Let sScript = StringSegment(sScripts, "\7", i)
Let s = StringLower(StringTrim(StringReplaceCase(sScript, "(", " (")))
Let s = StringReplaceAllCase(s, "  ", " ")
Let s1 = StringSegment(s, " ", 1)
Let s2 = StringSegment(s, " ", 2)
Let s3 = StringSegment(s, " ", 3)
If StringEqual(s1, "script") Then
Let sKey = "s" + s2
Else
Let sKey = "f" + s3
EndIf
Let sKey = StringPadRight(sKey, " ", 50)
Let sText = sText + sKey + sScript + "\7"
Let i = i + 1
EndWhile
Let sScripts = StringChopRight(sText, 1)
Let sScripts = StringSegmentSort(sScripts, "\7")

Let sText = ""
Let iCount = StringCountSegment(sScripts, "\7")
Let i = 1
While i <= iCount
Let sScript = StringSegment(sScripts, "\7", i)
Let sScript = StringChopLeft(sScript, 50)
Let sText = sText + sScript + "\r\n\r\n"
Let i = i + 1
EndWhile
Let sText = StringChopRight(sText, 1)
Let sMatch = "\f"
Let sReplace = ""
Let sText = RegExpReplaceCase(sText, sMatch, sReplace)
Let sMatch = "(\r\n){3,}"
Let sReplace = "\r\n\r\n"
Let sText = StringTrim(sStart) + "\r\n\r\n" + StringTrim(RegExpReplaceCase(sText, sMatch, sReplace)) + "\r\n"
;Return StringToFile(sText, sJss)
Return sText
EndFunction

String Function StringSwapListCase(String sText, String sMatchList, String sReplaceList)
Return StringSwapListEx(sText, sMatchList, sReplaceList, "\7", True)
EndFunction

String Function StringSwapListEquiv(String sText, String sMatchList, String sReplaceList)
Return StringSwapListEx(sText, sMatchList, sReplaceList, "\7", False)
EndFunction

String Function StringSwapListEx(String sText, String sMatchList, String sReplaceList, String sDelimiter, Int iCaseSensitive)
Var
Int i, Int iCount,
String sMatch, String sReplace, String sReturn

Let sReturn = sText
Let iCount = StringCountSegment(sMatchList, sDelimiter)
Let i = 1

If iCaseSensitive Then
While i <= iCount
Let sMatch = StringSegment(sMatchList, sDelimiter, i)
If StringContains(sText, sMatch) Then
Let sReplace = StringSegment(sReplaceList, sDelimiter, i)
Let sReturn = StringReplaceAllCase(sReturn, sMatch, sReplace)
EndIf
Let i = i + 1
EndWhile
Else
While i <= iCount
Let sMatch = StringSegment(sMatchList, sDelimiter, i)
If StringContainsEquiv(sText, sMatch) Then
Let sReplace = StringSegment(sReplaceList, sDelimiter, i)
Let sReturn = StringReplaceAllEquiv(sReturn, sMatch, sReplace)
EndIf
Let i = i + 1
EndWhile
EndIf
Return sReturn
EndFunction

Int Function StringToFile(String sText, String sFile)
;Saves string to text file, replacing if it exists

Var
Int iReturn, Int iReplace,
Object oSystem, Object oFile, Object oNull

If FileDelete(sFile) Then
Let oSystem =ObjectCreate("Scripting.FilesystemObject")
Let iReplace = True
Let oFile =oSystem.CreateTextFile(sFile, iReplace)
oFile.write(sText)
oFile.close()
Let iReturn = FileExists(sFile)
Let oFile = oNull
Let oSystem = oNull
Endif

Return iReturn
EndFunction

Int Function StringTrailCase(String s1, String s2)
;Case sensitive version of StringTrailEx
Return StringTrailEx(s1, s2, True)
EndFunction

Int Function StringTrailEquiv(String s1, String s2)
;Case equivalent version of StringTrailEx
Return StringTrailEx(s1, s2, False)
EndFunction

Int Function StringTrailEx(String s1, String s2, Int iCaseSensitive)
;Test whether second string matches Trailing part of first string

If iCaseSensitive Then
Return StringEqual(StringRight(s1, StringLength(s2)), s2)
Else
Return StringEquiv(StringRight(s1, StringLength(s2)), s2)
Endif
EndFunction

String Function StringTrim(String s)
;Remove leading and trailing blanks from a string
Return StringTrimLeadingBlanks(StringTrimTrailingBlanks(s))
EndFunction

String Function StringUnquote(String sText)
;Remove leading and trailing quotes from a string
Var
String sReturn

Let sReturn = sText
While StringLeft(sReturn, 1) == "\34"
Let sReturn = StringChopLeft(sReturn, 1)
EndWhile

While StringRight(sReturn, 1) == "\34"
Let sReturn = StringChopRight(sReturn, 1)
EndWhile
Return sReturn
EndFunction

Int Function TabPageGet(Handle h)
;Get index of the active tab page

Var
Int iMessage

If !h Then
Let h = GetRealWindow(GetFocus())
EndIf

Let iMessage = 4875 ;TCM_GETCURSEL
Return SendMessage(h, iMessage, 0, 0)
EndFunction

Int Function TabPageSet(Handle h, Int iPage)
;Activate a tab page by its index
Var
Int iMessage

If !h Then
Let h = GetRealWindow(GetFocus())
EndIf

Let iMessage = 4876 ;TCM_SETCURSEL
Return SendMessage(h, iMessage, iPage, 0)
EndFunction

Void Function UIElevateVersion(String sURL)
;Download and run an installer from a URL
Var
Int iChoice,
String sFile, String sText

SayString("Elevate Scripts")
Let sFile = PathGetName(sURL)
Let sFile = PathGetInternetCacheFolder() + "\\" + sfile
Let sText = "Download from\n"
Let sText = sText + sURL + "\n"
Let sText = sText + "To Internet cache, and run installer?"
If DialogConfirm("", sText, "Y") != "Y" Then
Return
EndIf

SayString("Please wait")
If !ShellUrlToFile(sURL, sFile) Then
SayString("Error downloading file!")
Return
EndIf
;Run("\34" + sFile + "\34")
ShellRun("\34" + sFile + "\34", 1, 1)
FileDelete(sFile)
EndFunction

Int Function UIGoToEditorCol(Int iCol)
;Go to a column in an editor
Var
Int i, Int iDelta, Int iLine, Int iReturn

SaveEditorPosition()
Let iLine = iHomerEditorLine
Let iDelta = iCol - iHomerEditorCol
If iDelta >= 0 Then
Let i = iDelta
While i
NextCharacter()
SaveEditorPosition()
If iHomerEditorLine != iLine Then
Let i = -2
ElIf iHomerEditorCol > iCol Then
Let i = -1
ElIf iHomerEditorCol == iCol Then
Let i = 0
Else
Let i = i - 1
EndIf
EndWhile
Else
Let i = -1 * iDelta
While i
PriorCharacter()
SaveEditorPosition()
If iHomerEditorLine != iLine Then
Let i = -2
ElIf iHomerEditorCol < iCol Then
Let i = -1
ElIf iHomerEditorCol == iCol Then
Let i = 0
Else
Let i = i - 1
EndIf
EndWhile
EndIf
If iHomerEditorLine == iLine && iHomerEditorCol == iCol Then
Let iReturn = True
EndIf
Return iReturn
EndFunction

Int Function UIGoToEditorLine(Int iLine)
;Go to a line in an editor
Var
Int iReturn

SaveEditorPosition()
If iHomerEditorLine == iLine Then
Return True
EndIf

If !KeyboardTypeAndWait("Control+G", "Edit", 10) Then
Return False
EndIf

TypeString(IntToString(iLine))
If !KeyboardTypeAndWait("Enter", "VsTextEditPane", 10) Then
Return False
EndIf

SaveEditorPosition()
If iHomerEditorLine == iLine Then
Let iReturn = True
EndIf

Return iReturn
EndFunction

Int Function UIGoToEditorPosition(Int iLine, Int iCol)
;Go to an editor line and column position
If !UIGoToEditorLine(iLine) Then
Return False
EndIf
If !UIGoToEditorCol(iCol) Then
Return False
EndIf
Return True
EndFunction

int Function UIHandleHomerWindows(handle hWnd)
;Suppress speech when returning from a dialog
Var
Handle hReal

if 0 then
return false
endif

Let hReal = GetRealWindow(hWnd)
If StringEqual(GetWindowName(hReal), "JAWS") Then
Return True
;ElIf StringEqual(GetWindowClass(hReal), "#32770") Then
ElIf StringEquiv(GetAppFileName(), "IniForm.exe") Then
If hWnd == hReal Then
Return True
Else
Return False
EndIf
ElIf InHJDialog() Then
If hWnd == hReal Then
Return False
ElIf StringEquiv(GetWindowClass(hWnd), "Static") Then
;Return True
Return False
;Else
ElIf StringEquiv(GetWindowClass(hWnd), "Edit") Then
SayString("Edit")
;SayObjectTypeAndText()
;SayString(EditGetText(hWnd))
;PerformScript SayLine()
;SayWindowTypeAndText(hWnd)
Return True
Else
;Return False
Return True
EndIf
Else
Return IniReadInteger("Internal", "SuppressTitle", False, "Homer.ini")
EndIf
EndFunction

Int Function UISelectToPosition(Int iLine, Int iCol)
;Select text from a saved position to the current one
Var
Int i, Int iDelta, Int iReturn

SaveEditorPosition()
If iHomerEditorLine > iLine || (iHomerEditorLine == iLine && iHomerEditorCol > iCol) Then
Return False
EndIf

Let iDelta = iLine - iHomerEditorLine
While iHomerEditorLine < iLine && iDelta > 0
SelectNextLine()
SaveEditorPosition()
Let iDelta = iDelta - 1
EndWhile
If iHomerEditorLine != iLine Then
Return False
EndIf

Let iDelta = iCol - iHomerEditorCol
While iHomerEditorLine == iLine && iHomerEditorCol < iCol && iDelta > 0
SelectNextCharacter()
SaveEditorPosition()
Let iDelta = iDelta - 1
EndWhile
If iHomerEditorLine == iLine && iHomerEditorCol == iCol Then
Let iReturn = True
EndIf
Return iReturn
EndFunction

Int Function UIWaitForControl(String sClasslist, String sIDlist, Int iTime)
;Wait until focus moves to one or more possible controls, based on lists of classes and IDs
;Return success/failure after the maximum milliseconds specified

Var
Int iItem, Int iReturn,
String sID, String sClass

While iTime >0 && !IsKeyWaiting()
Let iItem =1
While iItem >0
Let sClass =StringSegment(sClasslist, "|", iItem)
Let sID =StringSegment(sIDlist, "|", iItem)
If StringIsBlank(sClass) && StringIsBlank(sID) Then
Delay(1)
Let iTime =iTime -1
Let iItem =-1
Else
If StringIsBlank(sClass) Then
Let sClass =StringSegment(sClasslist, "|", 1)
ElIf StringIsBlank(sID) Then
Let sID =StringSegment(sIDlist, "|", 1)
EndIf
If (StringIsBlank(sClass) || StringEquiv(GetWindowClass(GetFocus()), sClass)) && (StringIsBlank(sID) || StringEquiv(IntToString(GetControlID(GetFocus())), sID)) Then
Let iReturn =True
Let iTime =-1
Let iItem =-1
Pause()
Delay(1)
Else
Let iItem =iItem +1
EndIf
EndIf
EndWhile
EndWhile
Return iReturn
EndFunction

Int Function UIWaitForTitleAndActivate(String sTitle, Int iTime)
;Wait for a title to appear and activate it
Var
Handle h

;While !iTime && !FindTopLevelWindow("", sTitle)
While iTime && !h 
Let h = FindTopLevelWindow("", sTitle)
Delay(1)
Let iTime = iTime - 1
EndWhile

If iTime && !IntEqual(h, GetRealWindow(GetFocus())) Then
if 1 then
Pause()
AppActivate(sTitle)
Pause()
Let h = FindWindow(h, "Edit", "")
If h Then
SetFocus(h)
EndIf
Else
;Pause()
SpeechOff()
AppActivateTitle("JAWS")
;WindowHide(h)
;Pause()
;WindowMaximize(h)
;Pause()
;WindowShow(h)
Pause()
;AppActivateTitle(sTitle)
SpeechOn()
SetFocus(h)
Pause()
EndIf
EndIf
Return StringEqual(GetWindowName(GetTopLevelWindow(GetFocus())), sTitle)
EndFunction

; Thanks to Doug Lee for much of the ideas and code related to variants
String Function VariantGetSubtype(Variant v)
; Get subtype of a variant

Var
String sReturn

Let sReturn = "Unknown"
If v Then
if !StringLength(v) then
; Only an object can have length 0 and remain boolean True.
Let sReturn = "Object"
ElIf StringLength(v) == StringLength(intToString(v)) then
; Looks like a number.
Let sReturn = "Int"
Else
; Looks like a string.
Let sReturn = "String"
EndIf
Else
; Looks like nothing at all.
Let sReturn = "Null"
EndIf
Return sReturn
EndFunction

Handle Function VariantToHandle(Variant v)
; Cast a variant to a Handle

Return v
EndFunction

Int Function VariantToInt(Variant v)
; Cast a variant to an integer

Return v
EndFunction

Object Function VariantToObject(Variant v)
; Cast a variant to an object

Return v
EndFunction

String Function VariantToString(Variant v)
; Cast a variant to a string

Return v
EndFunction

Variant Function VBSEval(String sCode, Variant v1, Variant v2, Variant v3, Variant v4)
Var
Object oWSC, Object oNull,
String sWSC,
Variant vReturn

Let sWSC = GetJAWSSettingsDirectory() + "\\Homer\\Homer.wsc"
Let oWSC = GetObject("script:" + sWSC)
Let vReturn = oWSC.VBSEval(sCode, v1, v2, v3, v4)
Let oWSC = oNull
Return vReturn
EndFunction

Variant Function VBSEvalFile(String sFile, Variant v1, Variant v2, Variant v3, Variant v4)
Var
String sCode,
Variant vReturn

If !StringContains(sFile, "\\") Then
Let sFile = GetJAWSSettingsDirectory() + "\\Homer\\" + sFile
EndIf

Let sCode = FileToString(sFile)
Let vReturn = VBSEval(sCode, v1, v2, v3, v4)
Return vReturn
EndFunction

Void Function VoiceSaveSetting(String sSetting, Int iLevel)
;Save an Eloquence voice setting
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

Int Function VoiceSetScreenEcho(Int iLevel)
;Set screen echo and return previous level

Var
Int iReturn

Let iReturn =GetScreenEcho()
While GetScreenEcho() !=iLevel
ScreenEcho()
EndWhile
Return iReturn
EndFunction

Int Function VoiceSetSpeech(Int iState)
;Set speech on or off and return previous state
Var
Int iReturn

Let iReturn = !IsSpeechOff()
If iState Then
SpeechOn()
Else
SpeechOff()
EndIf
Return iReturn
EndFunction

Int Function VoiceSetVerbosity(Int iLevel)
;Set verbosity and return previous level

Var
Int iReturn

Let iReturn =GetVerbosity()
While GetVerbosity() !=iLevel
VerbosityLevel()
EndWhile
Return iReturn
EndFunction

Int Function WebGetUrlToFile(String sUrl, String sFile)
;Download file from Internet to disk using WebGet.exe
Var
String sDir, String sText, String sCommand, String sTmp, String sExe

Let sExe = GetJAWSSettingsDirectory() + "\\Homer\\WebGet.exe"
If !FileExists(sExe) Then
Return
EndIf

Let sExe = PathGetShort(sExe)
Let sTmp = PathGetTempFile()
Let sDir = PathGetFolder(sFile)
If !FolderExists(sDir) Then
FolderCreate(sDir)
EndIf
If !FolderExists(sDir) Then
Return
EndIf

Let sText = sDir + "\r\n" + sUrl
StringToFile(sText, sTmp)
Let sCommand = sExe + " " + sTmp
ShellRun(sCommand, 0, True)
FileDelete(sTmp)
Return FileExists(sFile)
EndFunction

Int Function WindowActivate(Handle h, Int iState)
;Activate a window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetRealWindow(GetFocus())
EndIf

Let iMessage = 0x6 ;WM_ACTIVATE
;Istate options
;WA_INACTIVE    = 0
;WA_ACTIVE      = 1
;WA_CLICKACTIVE = 2
Return SendMessage(h, iMessage, iState, 0)
EndFunction

Int Function WindowClose(Handle h)
;Close a window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetRealWindow(GetFocus())
EndIf

;Let iMessage = 0x10 ;WM_CLOSE
Let iMessage = 0x112 ;WM_SYSCOMMAND
Let iParam = 0xF060 ;SC_CLOSE
Return SendMessage(h, iMessage, iParam, 0)
EndFunction

Int Function WindowHide(Handle h)
;Hide a window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x112 ;WM_SYSCOMMAND
Let iParam = 0 ;SW_HIDE
Return SendMessage(h, iMessage, iParam, 0)
EndFunction

Int Function WindowMaximize(Handle h)
;Maximize a window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x112 ;WM_SYSCOMMAND
Let iParam = 0xF030 ;SC_MAXIMIZE
;Let iParam = 3 ;SW_SHOWMAXIMIZED
;Let iParam = 11 ;SW_MAX
Return SendMessage(h, iMessage, iParam, 0)
EndFunction

Int Function WindowMinimize(Handle h)
;Minimize a window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x112 ;WM_SYSCOMMAND
Let iParam = 0xF020 ;SC_MINIMIZE
;Let iParam = 2 ;SW_SHOWMINIMIZED
;Let iParam = 6 ;SW_MINIMIZE
Return SendMessage(h, iMessage, iParam, 0)
EndFunction

Int Function WindowNext(Handle h)
;Cycle to next MDI window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x112 ;WM_SYSCOMMAND
Let iParam = 0xF040 ;SC_NEXTWINDOW
Return SendMessage(h, iMessage, iParam, 0)
EndFunction

Int Function WindowPrevious(Handle h)
;Cycle to previous MDI window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x112 ;WM_SYSCOMMAND
Let iParam = 0xF050 ;SC_PREVWINDOW
Return SendMessage(h, iMessage, iParam, 0)
EndFunction

Int Function WindowRestore(Handle h)
;Restore default size of window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x112 ;WM_SYSCOMMAND
Let iParam = 0xF120 ;SC_RESTORE
;Let iParam = 9 ;SW_RESTORE
Return SendMessage(h, iMessage, iParam, 0)
EndFunction

Int Function WindowSetState(Handle h, Int iState)
;Set state of a window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x18 ;WM_SHOWWINDOW
;Let iMessage = 0x112 ;WM_SYSCOMMAND
;Let iParam = 1 ;SW_SHOWNORMAL
;Let iParam = 5 ;SW_SHOW
;Let iParam = 10 ;SW_SHOWDEFAULT
Return SendMessage(h, iMessage, iState, 0)
EndFunction

Int Function WindowShow(Handle h)
;Show a window
Var
Int iMessage, Int iParam

If !h Then
Let h = GetFocus()
EndIf

Let iMessage = 0x18 ;WM_SHOWWINDOW
;Let iMessage = 0x112 ;WM_SYSCOMMAND
;Let iParam = 1 ;SW_SHOWNORMAL
;Let iParam = 5 ;SW_SHOW
;Let iParam = 10 ;SW_SHOWDEFAULT
Return SendMessage(h, iMessage, iParam, 0)
EndFunction

Script UIDeleteDown()
;Delete to bottom of file and say line
If CaretVisible() && IsPCCursor() && !IsVirtualPCCursor() Then
SayString("Delete down")
SpeechOff()
SelectToBottom()
Pause()
TypeKey("Delete")
Pause()
RefreshWindow(GetFocus())
Pause()
SpeechOn()
SayLine()
Else
TypeCurrentScriptKey()
EndIf
EndScript

Script UIDeleteLeft()
;Delete to beginning of line and say line

If CaretVisible() && IsPCCursor() && !IsVirtualPCCursor() Then
;SayString("Delete left")
SpeechOff()
SelectFromStartOfLine()
Pause()
TypeKey("Delete")
Pause()
RefreshWindow(GetFocus())
Pause()
SpeechOn()
SayLine()
Else
TypeCurrentScriptKey()
EndIf
EndScript

Script UIDeleteLine()
;Delete line and say new one
If CaretVisible() && IsPCCursor() && !IsVirtualPCCursor() Then
;SayString("Delete line")
SpeechOff()
Pause()
JawsEnd()
JawsHome()
;Pause()
SelectToEndOfLine()
Pause()
TypeKey("Delete")
;Pause()
RefreshWindow(GetFocus())
Pause()
SpeechOn()
SayLine()
Else
TypeCurrentScriptKey()
EndIf
EndScript

Script UIDeleteRight()
;Delete to end of line and say line

If CaretVisible() && IsPCCursor() && !IsVirtualPCCursor() Then
;SayString("Delete right")
SpeechOff()
SelectToEndOfLine()
Pause()
TypeKey("Delete")
Pause()
RefreshWindow(GetFocus())
Pause()
SpeechOn()
SayLine()
Else
TypeCurrentScriptKey()
EndIf
EndScript

Script UIDeleteUp()
;Delete to top of file and say line
If CaretVisible() && IsPCCursor() && !IsVirtualPCCursor() Then
SayString("Delete up")
SpeechOff()
SelectFromTop()
Pause()
TypeKey("Delete")
Pause()
RefreshWindow(GetFocus())
Pause()
SpeechOn()
SayLine()
Else
TypeCurrentScriptKey()
EndIf
EndScript

Script UISayAndTypeCurrentScriptKey()
;Say and type the key that invoked the current script
SayCurrentScriptKeyLabel()
TypeCurrentScriptKey()
EndScript

Script UITypeCurrentScriptKey()
TypeCurrentScriptKey()
EndScript
