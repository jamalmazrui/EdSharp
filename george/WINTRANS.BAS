' wintrans 1.1 windows skin for NfbTrans.exe
' Author: Anonymous John
' Begin date: 10/06/02
' Language: pb win 7.0

' metastatements
#Compile Exe
#Dim All

' included files
#Include "win32api.inc" ' windows api
#Include "comdlg32.inc" ' common windows dialogs
#Include "richedit.inc" ' for viewing stats

' constants
%Nfb_Input = %Wm_User + 1024 ' input file button
%Nfb_Output = %Nfb_Input + 1 ' output file button
%Nfb_BrfTxt = %Nfb_Output + 1 ' braille to text radio button
%Nfb_Txtbrf = %Nfb_brftxt + 1 ' text to braille radio button
%Nfb_Ok = %Nfb_txtbrf + 1 ' translate button
%Nfb_Cancel = 0 ' exit button
%Nfb_Stat = %Nfb_Ok + 1 ' status line
%NFB_MultiVol = %NFB_stat + 1 ' Multi volume checkstate
%nfb_pref = %NFB_MultiVol + 1 ' preferences button
%prpw = %Nfb_Pref + 1 ' page width textbox
%prpl = %prpw + 1 ' page length text box
%prps = %prpl + 1 ' starting page text box
%prpE = %prps + 1 ' ending page text box
%prff = %prpE + 1 ' formfeeds check box
%prvt = %prff + 1 ' vTab check box
%prbl = %prvt + 1 ' blank lines check box
%prIp = %prbl + 1 ' interpoint check box
%prlm = %prIp + 1 ' left margin text box
%prcO = %prlm + 1 ' copies text box
%prcf = %prcO + 1 ' config file text box
%prlE = %prcf + 1 ' format file text box
%prtf = %prlE + 1 ' table files check box
%prsO = %prtf + 1 ' sound on/off check box
%prcp = %prsO + 1 ' cmd line parms text box
%prpr = %prcp + 1 ' embosser port textbox
%PrApply = %Prpr + 1 ' apply button
%PrOk = %PrApply + 1 ' ok button
%PrCancel = 0 ' cancel button
%Nfb_prt = %prOk + 1 ' emboss button
%Nfb_Stats = %Nfb_Prt + 1 ' stats button
%StatsBox = %Nfb_Stats + 1 ' statistics text box
%StatsCancel = 0 ' statistics close button
%lList = %StatsBox + 1 ' volumes list
%NfbmTrans = %lList + 1 ' multivolume translate button
%NfbmEmboss = %NfbmTrans + 1 ' multivolume emboss button
%NfbmCancel = %NfbmEmboss + 1 ' multivolume cancel button

' declarations
Declare Function GetFile(Op As Long, title As String) As String ' filespec browse dialog
Declare CallBack Function InputButton() As Long
Declare CallBack Function OutputButton() As Long
Declare CallBack Function MvProc() As Long
Declare CallBack Function preferences() As Long
Declare CallBack Function StatsButton() As Long
Declare CallBack Function TransButton() As Long
Declare CallBack Function EmbossButton() As Long
Declare CallBack Function prApply() As Long
Declare CallBack Function PrOk() As Long
Declare CallBack Function CancelButton() As Long
Declare Sub StatLine
Declare Sub ShowFiles
Declare CallBack Function mTransButton() As Long
Declare CallBack Function mEmbossButton() As Long
Declare Function getFilePath$(fI As Asciiz)
Declare Function NfbShell(tProg As Asciiz * 255, tParms As Asciiz * 128) As Long
Declare Sub MultiVolume(op As Long)
Declare Sub ReadInit
Declare Sub WriteInit
Declare Function BuildParms() As Long
Declare Sub getPrefs(hdlgn As Long)

' globals
Global hdlg As Long ' dialog handle
Global lDialog As Long ' dialog handle
Global mv As Long ' multi volume checkstate
Global TransOpt As Long ' b2t or t2b
Global InFile As Asciiz * 255 ' input filespec
Global OutFile As Asciiz * 255 ' Output filespec
Global inPath As String ' default input path
Global OutPath As String ' default output path
Global HomePath As String ' home dir for wintrans.exe
Global pw As String ' page width
Global pl As String ' page length
Global ps As String ' starting page
Global pE As String ' ending page
Global ls As String ' page break
Global ip As String ' interpoint toggle
Global lm As String ' left margin
Global co As String ' copies
Global cf As String ' config file
Global le As String ' format file
Global tf As String ' table files
Global so As String ' sound on/off
Global cp As String ' extra command line parms
Global pr As String ' embosser port
Global parms As String ' command line parameters for nfbtrans
Global fls() As String ' files list
global debug as long ' %true show cmd line in msg box

Sub ReadInit ' reads wintrans.ini
Local f As Long ' marker
Local f1 As Long ' marker
Local nFile As Long ' file handle
Local buf As String ' file buffer
Local q As String ' = chr$(34)

q = Chr$(34)
nFile = FreeFile
Open HomePath + "WINTRANS.INI" For Binary As #nFile
Get$ #nFile, Lof(nFile), buf
Close #nFile
buf = LCase$(buf)
f = InStr(buf, "pw=") + 3
f1 = InStr(f, buf, " ")
pw = Mid$(buf, f, f1 - f)
f = InStr(buf, "pl=") + 3
f1 = InStr(f, buf, " ")
pl = Mid$(buf, f, f1 - f)
f = InStr(buf, "ps=") + 3
f1 = InStr(f, buf, " ")
ps = Mid$(buf, f, f1 - f)
f = InStr(buf, "pe=") + 3
f1 = InStr(f, buf, " ")
pE = Mid$(buf, f, f1 - f)
f = InStr(buf, "ls=") + 3
f1 = InStr(f, buf, " ")
ls = Mid$(buf, f, f1 - f)
f = InStr(buf, "ip=") + 3
f1 = InStr(f, buf, " ")
ip = Mid$(buf, f, f1 - f)
f = InStr(buf, "lm=") + 3
f1 = InStr(f, buf, " ")
lm = Mid$(buf, f, f1 - f)
f = InStr(buf, "co=") + 3
f1 = InStr(f, buf, " ")
co = Mid$(buf, f, f1 - f)
f = InStr(buf, "cf=") + 3
f1 = InStr(f, buf, " ")
cf = Mid$(buf, f, f1 - f)
f = InStr(buf, "le=") + 3
f1 = InStr(f, buf, " ")
le = Mid$(buf, f, f1 - f)
f = InStr(buf, "tf=") + 3
f1 = InStr(f, buf, " ")
tf = Mid$(buf, f, f1 - f)
f = InStr(buf, "so=") + 3
f1 = InStr(f, buf, " ")
so = Mid$(buf, f, f1 - f)
f = InStr(buf, "cp=") + 3
f1 = InStr(f, buf, " ")
cp = Mid$(buf, f, f1 - f)
f = InStr(buf, "pr=") + 3
f1 = InStr(f, buf, " ")
pr = Mid$(buf, f, f1 - f)
f = InStr(buf, "tr=") + 3
f1 = InStr(f, buf, " ")
TransOpt = Val(Mid$(buf, f, f1 - f))
f = InStr(buf, "inp=") + 5
f1 = InStr(f, buf, q)
InPath = Mid$(buf, f, f1 - f)
f = InStr(buf, "out=") + 5
f1 = InStr(f, buf, q)
OutPath = Mid$(buf, f, f1 - f)
f = instr(buf, "debug=") + 6
if mid$(buf, f, 1) = "y" then debug = %true
End Sub

Sub WriteInit ' writes wintrans.ini
Local nFile As Long ' file handle
Local q As String ' = chr$(34)

q = Chr$(34)
nFile = FreeFile
Open HomePath + "WINTRANS.INI" For Output As #nFile
Print #nFile, "lm=" + lm + " left margin"
Print #nFile, "co=" + co + " copies"
Print #nFile, "pw=" + pw + " page width"
Print #nFile, "pl=" + pl + " page length"
Print #nFile, "ps=" + ps + " starting page"
Print #nFile, "pE=" + pE + " ending page"
Print #nFile, "ls=" + ls + " page break"
Print #nFile, "ip=" + ip + " interpoint"
Print #nFile, "so=" + so + " sound"
Print #nFile, "cf=" + cf + " config file"
Print #nFile, "le=" + le + " format file"
Print #nFile, "tf=" + tf + " translation table file"
Print #nFile, "cp=" + cp + " command line parms"
Print #nFile, "pr=" + pr + " embosser port"
Print #nFile, "tr=" + Format$(transopt) + " translation mode"
Print #nFile, "inp=" + q + InPath + q + " input path"
Print #nFile, "out=" + q + OutPath + q + " output path"
print #nFile, "debug=";
if debug = %true then
  print #nFile, "yes ";
else
  print #nFile, "no ";
end if
print #nFile, "yes = show NfbTrans calls"
Close #nFile
End Sub

Function BuildParms() As Long  ' builds command line
Local f As Long ' marker
Local fi As String ' parsing file buffer
parms = "nc=1"
If lm <> "1" Then parms = parms + " lm=" + lm
If co <> "1" Then parms = parms + " co=" + co
If pw <> "40" Then parms = parms + " pw=" + pw
If pl <> "25" Then parms = parms + " pl=" + pl
If ps <> "1" Then parms = parms + " ps=" + ps
If pE <> "9999" Then parms = parms + " pE=" + pE
If ls <> "99" Then parms = parms + " ls=" + ls
If ip <> "0" Then parms = parms + " ip=" + ip
If so <> "0" Then parms = parms + " so=" + so
If cf <> "" Then
  ' does it exist?
  If Len(Dir$(cf)) Then
    FileCopy cf, HomePath + "TEMP.CNF"
    parms = parms + " cf=TEMP.CNF"
  Else
    MsgBox "Can't find " + cf + "using nfbtrans.cnf", %mb_IconHand, "Config file error"
    Function = %true
  End If
End If
If le <> "" Then
  If Len(Dir$(le)) Then
    parms = parms + " le=" + le
    parms = parms + " ef=1"
  Else
    MsgBox "Can't find " + le + " None used", %mb_IconHand, "Format file error"
    Function = %true
  End If
End If
If tf <> "" Then
  ' fill out the spec
  If Left$(tf, 1) = ";" Then tf = "BRAILLE.TAB" + tf
  If Right$(tf, 1) = ";" Then tf = tf + "BACK.TAB"
  If InStr(tf, ";") = 0 Then tf = tf + ";BACK.TAB"
  ' do they exist?
  f = InStr(tf, ";")
  fi = Left$(tf, f - 1) ' braille table
  If Dir$(fi) = "" Then
    MsgBox "Can't find " + fi + " Using BRAILLE.TAB", %mb_IconHand, "Table file error
    tf = "BRAILLE.TAB" + Mid$(tf, f)
    f = InStr(tf, ";")
    Function = %true
 Else
   FileCopy fi, HomePath + "TEMPBRL.TAB"
  End If
  fi = Mid$(tf, f + 1)
  If Dir$(fi) = "" Then
    MsgBox "Can't find " + fi + " using BACK.TAB", %mb_IconHand, "Table file error"
    tf = Left$(tf, f) + "BACK.TAB"
    Function = %true
  Else
    FileCopy fi, HomePath + "TEMPBAK.TAB"
  End If
  parms = parms + " tf=tempbrl.tab;tempbak.tab"
End If
If cp <> "" Then parms = parms + " cp=" + cp
parms = parms + " "
End Function

Sub getPrefs(hdlgn As Long)

Local lc As Long ' checkbox state
' get info from text boxes
Control Get Text hdlgn, %prlm To lm
Control Get Text hdlgn, %prcO To co
Control Get Text hdlgn, %prpw To pw
Control Get Text hdlgn, %prpl To pl
Control Get Text hdlgn, %prps To ps
Control Get Text hdlgn, %prpE To pE
Control Get Text hdlgn, %prcf To cf
Control Get Text hdlgn, %prlE To lE
Control Get Text hdlgn, %prtf To tf
Control Get Text hdlgn, %prcp To cp
Control Get Text hdlgn, %prpr To pr
' check our check boxes
Control Get Check hdlgn, %prIp To lc
ip = LTrim$(Str$(lc))
Control Get Check hdlgn, %prsO To lc
so = "0"
If lc = 1 Then sO = "15"
' radio buttons
Control Get Check hdlgn, %prff To lc
If lc = 1 Then ls = "99"
Control Get Check hdlgn, %prvt To lc
If lc = 1 Then ls = "999"
Control Get Check hdlgn, %prbl To lc
If lc = 1 Then ls = "0"
End Sub

Sub StatLine ' updates the status line

Local msg As String
msg = "Input: " + InFile + " Output: " + OutFile
Control Set Text hdlg, %NFB_Stat, msg
End Sub

Function GetFile(Op As Long, Title As String) As String
'
' op = operation: input/output
' title = string for dialog title

Local sFileName As String
Local fStyle As Long
Local a%
Local result As Long
Local f As Long
Local msg As Asciiz * 255
Local path As String ' starting path
Local CrtDir As String ' current directory

crtDir = Mid$(HomePath, 3, Len(HomePath) - 3)
sFileName = "*.*"
If op = 0 Then ' input file
  fStyle = %OFN_FILEMUSTEXIST Or %OFN_LONGNAMES
  If mv > 0 Then fStyle = %OFN_LONGNAMES
  a% = OpenFileDialog(hdlg, title, sFileName, InPath, _
    "All Files|*.*", "*", ByCopy fStyle)
Else ' output file
  sFileName = InputBox$("Enter output filespec ", title, OutPath)
End If
If sFileName = "*.*" Then sFileName = "" ' dialog canceled
Replace ".*" With "" In sFileName
If sFileName <> "" Then
  f = InStr(-1, sFileName, Any "\:")
  If f = 0 Then
    If op = 0 Then
      sFileName = InPath + sFileName
    Else
      sFileName = OutPath + sFileName
    End If
  Else
    path = UCase$(Left$(sFileName, f))
    If op = 0 Then
      InPath = path
    Else
      OutPath = path
    End If
  End If
End If
ChDir crtDir
Function = sFileName
End Function

Sub showFiles

Local style&, exStyle&

Style& = %WS_POPUP Or %DS_MODALFRAME Or %WS_CAPTION Or %WS_MINIMIZEBOX Or %WS_SYSMENU Or %WS_MAXIMIZEBOX Or %DS_CENTER
ExStyle& = 0

' show the dialog
Dialog New 0, "Multiple Volumes", ,, 350, 210, style&, exStyle& To lDialog
Control Add ListBox, lDialog, %lList, fls(), 25, 24, 300, 110, _
%WS_CHILD Or %WS_VISIBLE Or %LBS_notify Or %WS_VSCROLL Or %WS_TABSTOP Or _
%lbs_sort Or %lbs_MultipleSel, %WS_EX_CLIENTEDGE
Control Add Button, lDialog, %NfbmTrans, "&Translate", 100, 155, 40, 14 Call mTransButton
Control Add Button, lDialog, %NfbmEmboss, "&Emboss", 150, 155, 40, 14 Call mEmbossButton
Control Add Button, lDialog, %NfbmCancel, "&Cancel", 200, 155, 40, 14 Call CancelButton
Control Set Focus lDialog, %lList
Dialog Show Modal lDialog
End Sub

Function getFilePath$(fI As Asciiz)

Local f As Long ' string marker
Local fp As String ' path string
' get inFile path
f = InStr(-1, fI, "\")
If f > 0 Then fp = Left$(fI, f)
Function = fp
End Function

Function NfbShell(tProg As Asciiz * 255, tParms As Asciiz * 128) As Long

Local dos As Asciiz * 255
Local hPath As Asciiz * 255
Local w9x As Long ' w9x flag
Local cSpec As String ' comspec
Local f As Long ' marker

cSpec = LCase$(tProg)
f = InStr(cSpec, "command.com")
If f > 0 Then w9x = %true
w9x = %true
if debug = %true then msgbox trim$(tProg) + " " + trim$(tParms), , "NfbTrans Call"
If w9x = %true Then
  ' set tProg and tParms
  ' tParms = Mid$(tParms, 4)
  ' f = InStr(tParms, " tm=")
  ' tProg = Left$(tParms, f - 1)
  ' tParms = Mid$(tParms, f + 1)
  ' 9x shell
  ShellExecute 0, "open", tProg, tParms, hPath, %Sw_ShowMinNoActive
  Sleep 10000
Else ' xp shell
  Dos = tProg + " " + tParms
  Shell dos, 6
End If
End Function

Sub MultiVolume(op As Long)
' 0 = translate 1 = emboss
Local brf As String ' temp braille filespec
Local txt As String ' temp text file
Local dos As Asciiz * 255 ' shell commandSpec
Local tParms As Asciiz * 128 ' NfbTrans parms
Local x As Long ' counter
Local f As Long ' string marker
Local lf As Long ' file length
Local nFile As Long ' file handle
Local fs As String ' folder search

' setup temps
brf = HomePath + "TEMP.BRF"
txt = HomePath + "TEMP.TXT"
dos = Environ$("COMSPEC")
nFile = FreeFile
' process he input files
If TransOpt = 0 Then ' braille to text
  ' copy to short file names
  For x = 0 To UBound(fls)
    fs = fls(x)
    If Left$(fs, 1) = "|" Then
      fs = Mid$(fs, 2)
      brf = HomePath + "TEMP" + Right$("00" + LTrim$(Str$(x)), 3) + ".BRF"
      FileCopy fs, BRF
    End If
  Next x
  If op = 0 Then ' translate
    If mv = 1 Then ' combine output
      tParms = "/C " + HomePath + "NFBTRANS.EXE tm=13 " + parms + "TEMP*.BRF >" + txt
      NfbShell dos, tParms
      FileCopy txt, OutFile
    End If
    If mv = 2 Then ' separate output files
      For x = 0 To UBound(fls)
        tParms = "/C " + HomePath + "NFBTRANS.EXE tm=13 " + parms
        fs = fls(x)
        If Left$(fs, 1) = "|" Then
          fs = Mid$(fs, 2)
          fs = Left$(fs, Len(fs) - 4) + ".txt"
          f = InStr(-1, fs, "\")
          OutFile = OutPath + Mid$(fs, f + 1)
          brf = HomePath + "TEMP" + Right$("00" + LTrim$(Str$(x)), 3) + ".BRF"
          txt = HomePath + "TEMP" + Right$("00" + LTrim$(Str$(x)), 3) + ".TXT"
          tParms = tParms + brf + ">" + txt
          lf = 0
          ' msgbox "tParms = " + tParms + " OutFile = " + OutFile, , "multiple output files"
          If Len(Dir$(txt)) Then Kill txt
          NfbShell dos, tParms
          If Len(Dir$(txt)) Then
            FileCopy txt, OutFile
            Kill txt
          End If
        End If
      Next x
    End If
  End If
  If Op = 1 Then ' emboss
    For x = 0 To UBound(fls)
      brf = HomePath + "TEMP" + Right$("00" + LTrim$(Str$(x)), 3) + ".BRF"
      If Len(Dir$(brf)) Then FileCopy BRF, PR
    Next x
  End If
  StatLine
  ' kill the temp files
  For x = 0 To UBound(fls)
    brf = HomePath + "TEMP" + Right$("00" + LTrim$(Str$(x)), 3) + ".BRF"
    If Len(Dir$(brf)) Then Kill brf
  Next x
End If
If TransOpt = 1 Then ' text to braille
  ' copy to short file names
  For x = 0 To UBound(fls)
    fs = fls(x)
    If Left$(fs, 1) = "|" Then
      fs = Mid$(fs, 2)
      txt = HomePath + "TEMP" + Right$("00" + LTrim$(Str$(x)), 3) + ".TXT"
      FileCopy fs, txt
    End If
  Next x
  If op = 0 Then ' translate
    If mv = 1 Then ' combine output
      tParms = " /C " + HomePath + "nfbtrans.exe tm=11 " + parms
      tParms = tParms + "TEMP*.TXT >" + brf
      NfbShell dos, tParms
      FileCopy brf, OutFile
    End If
    If mv = 2 Then ' separate output files
      For x = 0 To UBound(fls)
        tParms = " /C " + HomePath + "nfbtrans.exe tm=11 " + parms
        fs = fls(x)
        If Left$(fs, 1) = "|" Then
          fs = Mid$(fs, 2)
          fs = Left$(fs, Len(fs) - 4) + ".brf"
          f = InStr(-1, fs, "\")
          OutFile = OutPath + Mid$(fs, f + 1)
          txt = HomePath + "TEMP" + Right$("00" + LTrim$(Str$(x)), 3) + ".TXT"
          tParms = tParms + txt + " >" + brf
          NfbShell dos, tParms
          FileCopy brf, OutFile
        End If
      Next x
    End If
  End If
  If Op = 1 Then ' emboss
    brf = HomePath + "EMBOSS.BRF"
    For x = 0 To UBound(fls)
      txt = HomePath + "TEMP" + Right$("00" + LTrim$(Str$(x)), 3) + ".TXT"
      If Len(Dir$(txt)) Then
        tParms = " /C " + HomePath + "nfbtrans.exe tm=11 " + parms + " " + TXT + " >" + brf
        NfbShell dos, tParms
        FileCopy BRF, PR
        Kill BRF
      End If
    Next x
  End If
  StatLine
  ' kill the temp files
  For x = 0 To UBound(fls)
    txt = HomePath + "TEMP" + Right$("00" + LTrim$(Str$(x)), 3) + ".TXT"
    If Len(Dir$(txt)) Then Kill txt
  Next x
End If
Control Set Focus hdlg, %Nfb_Input
End Sub

CallBack Function InputButton() As Long
Local fP As String ' old path
InFile = GetFile(0, "Input File")
fp = GetFilePath$(InFile)
If Len(fp) Then InPath = fp
StatLine
End Function

CallBack Function OutputButton() As Long
Local f As Long ' string marker
Local fln As String ' file spec
Local fp As String ' path spec
Dim FindHandle As Long
Dim FileData As WIN32_FIND_DATA
Do
  fln = GetFile(1, "Output file")
  If fln = "" Then
    Control Set Focus hdlg, %Nfb_Output
    Exit Function
  End If
  If InStr(fln, Any ":\") = 0 Then
    If Len(Dir$(fln)) Then
      f = MsgBox("Overwrite " + OutFile, %Mb_IconQuestion Or %Mb_YesNo, "File exists")
      If f <> 6 Then fln = ""
    End If
  End If
Loop Until Len(fln)
OutFile = fln
' is OutFile a folder?
If Right$(OutFile, 1) = "\" Then OutFile = Left$(OutFile, Len(OutFile) - 1)
FindHandle = FindFirstFile (OutFile, FileData)
f = FindClose (FindHandle)
If FileData.DwFileAttributes = %File_attribute_directory And FindHandle > 0 Then
  OutPath = OutFile
  OutFile = ""
Else
  fp = getFilePath$(OutFile)
  If Len(fp) Then OutPath = fp
End If
If Len(OutPath) And Right$(OutPath, 1) <> "\" Then OutPath = OutPath + "\"
StatLine
Control Set Focus hdlg, %Nfb_Output
End Function

CallBack Function statsButton() As Long
Local style&, exStyle& ' dialog styles
Local hdlgs As Long ' dialog handle
Local nFile As Long ' file handle
Local cs1 As Long ' text selection
Local cs2 As Long ' text selection
Local buffer As String ' text buffer

Style& = %WS_POPUP Or %DS_MODALFRAME Or %WS_CAPTION Or %WS_MINIMIZEBOX Or %WS_SYSMENU Or %WS_MAXIMIZEBOX Or %DS_CENTER
ExStyle& = 0
nFile = FreeFile
cs1 = 1
cs2 = 1

' get the text
If Dir$(HomePath + "stat.txt") = "" Then
  MsgBox "No statistics available.", %mb_IconAsterisk, "No stats"
  Exit Function
End If
Open HomePath + "stat.txt" For Binary As #nFile
Get$ #nFile, Lof(nFile), buffer
Close #nFile
' setup the dialog
Dialog New CbHndl, "NfbTrans Statistics", , , 350, 265, style&, exStyle& To hdlgs
' controls
Control Add TextBox, hdlgs, %Statsbox, buffer, 25, 15, 300, 235, _
  %WS_CHILD Or %WS_VISIBLE Or %es_WantReturn Or _
  %ES_READONLY Or %ES_MULTILINE Or _
  %ES_AUTOVSCROLL Or %ES_NOHIDESEL Or %Es_SaveSel, 0
Control Add Button, hdlgs, %StatsCancel, "&Close", 150, 250, 40, 14 Call CancelButton
Control Set Focus hdlgs, %StatsBox
Dialog Show Modal hdlgs
Control Set Focus hdlg, %Nfb_Input
End Function

CallBack Function mEmbossButton() As Long
Local x As Long
Local y As Long
Local z As Long
Local hResult As Long
Local fs As String

Control Send CbHndl, %lList, %LB_GETSELCOUNT, 0, 0, To x
If x = 0 Then
  MsgBox "No files selected", %Mb_IconAsterisk, "No files"
  Exit Function
End If
Dim f(x) As Long
Control Send CbHndl, %lList, %LB_GETSELITEMS, x, VarPtr(f(1))
For y = 1 To x
  z = f(y)
  fs = fls(z)
  fls(z) = "|" + fls(z)
  If OutFile = "" Then
    hResult = InStr(-1, fs, "\")
    fs = OutPath + Mid$(fs, hResult + 1)
    OutFile = Left$(fs, Len(fs) - 4) + ".TXT"
    If transopt = 1 Then OutFile = Left$(fs, Len(fs) - 4) + ".BRF"
  End If
Next y
mv = 3 ' emboss flag
Dialog End CbHndl, 0
End Function

CallBack Function mTransButton() As Long
Local x As Long
Local y As Long
Local z As Long
Local hResult As Long
Local fs As String

Control Send CbHndl, %lList, %LB_GETSELCOUNT, 0, 0, To x
Dim f(x) As Long
Control Send CbHndl, %lList, %LB_GETSELITEMS, x, VarPtr(f(1))
For y = 1 To x
  z = f(y)
  fs = fls(z)
  fls(z) = "|" + fls(z)
  If OutFile = "" Then
    hResult = InStr(-1, fs, "\")
    fs = OutPath + Mid$(fs, hResult + 1)
    OutFile = Left$(fs, Len(fs) - 4) + ".TXT"
    If transopt = 1 Then OutFile = Left$(fs, Len(fs) - 4) + ".BRF"
  End If
Next y
Dialog End CbHndl, 0
End Function

CallBack Function TransButton() As Long

Local f As Long ' string marker
Local f1 As Long ' string marker
Local dos As String' shell to NfbTrans
Local brf As String ' temp braille filespec
Local txt As String ' temp text filespec

' setup temp filespecs
txt = HomePath + "TEMP.TXT"
brf = HomePath + "TEMP.BRF"
' start a new stat.txt file
If Dir$(HomePath + "STAT.TXT") <> "" Then Kill HomePath + "STAT.TXT"
' Which radio button is checked?
Control Get Check hdlg, %Nfb_TxtBrf To TransOpt
' have we got everything?
If InFile = "" Then
  MsgBox "No Input file defined", %Mb_IconHand, "File error"
  Exit Function
End If
' don't translate .brf to braille or .txt to text
If TransOpt = 0 And LCase$(Right$(InFile, 4)) = ".txt" Then
  MsgBox InFile + " cannot be translated to text", %Mb_IconHand, "Entry error"
  Exit Function
End If
If TransOpt = 1 And LCase$(Right$(InFile, 4)) = ".brf" Then
  MsgBox InFile + " cannot be translated to braille", %Mb_IconHand, "Entry error"
  Exit Function
End If
f = InStr(-1, InFile, ".")
If f = 0 Then f = Len(InFile) + 1
f1 = InStr(-1, InFile, "\")
If OutFile = "" Then
  If TransOpt = 0 Then
    OutFile = Left$(InFile, f - 1) + ".TXT"
  Else
    OutFile = Left$(InFile, f - 1) + ".BRF"
  End If
  If OutPath <> "" Then OutFile = OutPath + Mid$(OutFile, f1 + 1)
End If
If LCase$(Infile) = LCase$(OutFile) Then
  MsgBox "file " + InFile + " cannot be translated onto itself", %Mb_IconHand, "Error"
  Exit Function
End If
' ensure short file names for NfbTrans and prepare shell statement
dos = Environ$("COMSPEC") + " /C "
If TransOpt = 0 Then
  FileCopy InFile, BRF
  Dos = dos + HomePath + "nfbtrans.exe tm=13 " + parms + brf + " >" + txt
Else
  FileCopy InFile, TXT
  Dos = dos + HomePath + "nfbtrans.exe tm=11 " + parms + txt + " >" + brf
End If
if debug = %true then msgbox dos$, , "NfbTrans Call"
' update status line
Control Set Text hdlg, %Nfb_Stat, "Translating"
Shell dos$, 6
StatLine
' check for errors
' kill empty files
Open txt For Binary As #1
f = Lof(1)
Close 1
If f = 0 Then Kill txt
Open brf For Binary As #1
f = Lof(1)
Close 1
If f = 0 Then Kill brf
If TransOpt = 0 Then
  If Dir$(txt) = "" Then
    MsgBox InFile + " to " + OutFile, %Mb_IconHand, "Translation error"
  Else
    MsgBox InFile + " to " + OutFile, %Mb_IconAsterisk Or %Mb_SetForeground, "Translated"
  End If
End If
If TransOpt = 1 Then
  If Dir$(brf) = "" Then
    MsgBox InFile + " to " + OutFile, %Mb_IconHand, "Translation error"
  Else
    MsgBox InFile + " to " + OutFile, %Mb_IconAsterisk, "Translated"
  End If
End If
' cleanup
If TransOpt = 0 And Dir$(txt) <> "" Then
  FileCopy txt, OutFile
  Kill txt
  Kill brf
End If
If TransOpt = 1 And Dir$(brf) <> "" Then
  FileCopy brf, OutFile
  Kill brf
  Kill txt
End If
OutFile = ""
Control Set Focus hdlg, %Nfb_Input
End Function

CallBack Function EmbossButton() As Long

Local f As Long ' string marker
Local dos As String' shell to NfbTrans
Local brf As String ' temp braille filespec
Local txt As String ' temp text filespec

' setup temp filespecs
txt = HomePath + "TEMP.TXT"
brf = HomePath + "TEMP.BRF"
OutFile = ""
' Which radio button is checked?
Control Get Check hdlg, %Nfb_TxtBrf To TransOpt
' have we got everything?
If InFile = "" Then
  MsgBox "No Input file defined", %Mb_IconHand, "File error"
  Exit Function
End If
' update status line
Control Set Text hdlg, %Nfb_Stat, "Embossing"
' ensure short file names for NfbTrans and prepare shell statement
dos = Environ$("COMSPEC") + " /C " + HomePath + "NFBTRANS.EXE "
If LCase$(Right$(InFile, 4)) = ".brf" Then
  FileCopy InFile, PR
End If
If LCase$(Right$(InFile, 4)) = ".txt" Then
  FileCopy InFile, TXT
  Dos = dos + "tm=11 " + parms + txt + " >" + brf
  if debug = %true then msgbox dos, , "NfbTrans Call"
  Shell dos$, 6
  FileCopy brf, PR
End If
StatLine
' kill temp files
If Len(Dir$(BRF)) Then Kill brf
If Len(Dir$(txt)) Then Kill txt
' cleanup
Control Set Focus hdlg, %Nfb_Input
End Function

CallBack Function CancelButton() As Long
InFile = ""
Dialog End CbHndl, 0
End Function

CallBack Function PrApply() As Long
GetPrefs(CbHndl)
BuildParms
End Function

CallBack Function PrOk() As Long
GetPrefs(CbHndl)
BuildParms
WriteInit
Dialog End CbHndl, 0
End Function

CallBack Function preferences() As Long

Local hdlgp As Long ' dialog handle
Local style&, exStyle&
Style& = %WS_POPUP Or %DS_MODALFRAME Or %WS_CAPTION Or %WS_MINIMIZEBOX Or %WS_SYSMENU Or %WS_MAXIMIZEBOX Or %DS_CENTER
ExStyle& = 0
' setup the dialog
Dialog New CbHndl, "Wintrans Preferences", , , 350, 130, style&, exStyle& To hdlgp
' controls
Control Add Label, hdlgp, -1, "Left &Margin", 10, 10, 40, 8
Control Add TextBox, hdlgp, %prlm, lm, 60, 10, 100, 12, %Es_autohscroll Or %Ws_Tabstop
Control Add Label, hdlgp, -1, "&Copies", 175, 10, 40, 8
Control Add TextBox, hdlgp, %prcO, co, 235, 10, 100, 12, %Es_Autohscroll Or %Ws_Tabstop
Control Add Label, hdlgp, -1, "Page &Width", 10, 25, 40, 8
Control Add TextBox, hdlgp, %prpw, pw, 60, 25, 100, 12, %es_Autohscroll Or %ws_tabstop
Control Add Label, hdlgp, -1, "Page &len", 175, 25, 40, 8
Control Add TextBox, hdlgp, %prpl, pl, 235, 25, 100, 12, %Es_Autohscroll Or %Ws_TabStop
Control Add Label, hdlgp, -1, "&Start page", 10, 40, 40, 8
Control Add TextBox, hdlgp, %prps, ps, 60, 40, 100, 12, %Es_Autohscroll Or %ws_Tabstop
Control Add Label, hdlgp, -1, "&End page", 175, 40, 40, 8
Control Add TextBox, hdlgp, %prpE, pE, 235, 40, 100, 12, %Es_Autohscroll Or %Ws_Tabstop
Control Add Label, hdlgp, -1, "Page Break", 25, 55, 40, 8
Control Add Option, hdlgp, %prff, "&FormFeed", 75, 55, 40, 14,_
  %ws_group Or %Bs_MultiLine, %Ws_ex_ClientEdge
Control Add Option, hdlgp, %prvt, "&V-Tab", 125, 55, 40, 14,_
  %BS_MultiLine, %Ws_ex_ClientEdge
Control Add Option, hdlgp, %prbl, "&Blanks", 175, 55, 40, 14,_
  %BS_MultiLine, %Ws_ex_ClientEdge
Control Add CheckBox, hdlgp, %PrIp, "&Interpoint", 225, 55, 40, 14
Control Add CheckBox, hdlgp, %prsO, "So&und", 275, 55, 40, 14
Control Add Label, hdlgp, -1, "Co&nfig file", 10, 70, 40, 8
Control Add TextBox, hdlgp, %prcf, cf, 60, 70, 100, 12, %Es_Autohscroll Or %Ws_Tabstop
Control Add Label, hdlgp, -1, "Fo&rmat file", 175, 70, 40, 8
Control Add TextBox, hdlgp, %prlE, le, 235, 70, 100, 12, %Es_Autohscroll Or %Ws_Tabstop
Control Add Label, hdlgp, -1, "&Table file", 10, 85, 40, 8
Control Add TextBox, hdlgp, %prtf, tf, 60, 85, 100, 12, %Es_Autohscroll Or %Ws_Tabstop
Control Add Label, hdlgp, -1, "Port", 175, 85, 40, 8
Control Add TextBox, hdlgp, %prpr, pr, 235, 85, 100, 12, %Es_Autohscroll Or %Ws_Tabstop
Control Add Label, hdlgp, -1, "Comman&ds", 10, 100, 40, 8
Control Add TextBox, hdlgp, %prcp, cp, 60, 100, 100, 12, %Es_Autohscroll Or %Ws_Tabstop
Control Add Button, hdlgp, %PrApply, "&Apply", 100, 115, 40, 14 Call PrApply
Control Add Button, hdlgp, %prOk, "&Ok", 150, 115, 40, 14 Call PrOk
Control Add Button, hdlgp, %PrCancel, "&Cancel", 200, 115, 40, 14 Call CancelButton
' set checkboxes and radio buttons
If so <> "0" Then Control Set Check hdlgp, %PrsO, 1
If ip = "1" Then Control Set Check hdlgp, %prIp, 1
' radio buttons
If ls = "99" Then Control  Set Check hdlgp, %prff, 1
If ls = "999" Then Control Set Check hdlgp, %prvt, 1
If ls = "0" Then Control Set Check hdlgp, %prbl, 1
' show the dialog
Dialog Show Modal hdlgp
End Function

CallBack Function MvProc() As Long
Local InFolder As Long ' inFile folder flag
Local OutFolder As Long ' outFile folder flag
Local result As Long ' combine or not msgBox
Local x As Long ' counter
Local fi As String ' brf FileSpec
Local fs As String ' folder search
Dim FindHandle As Long
Dim FileData As WIN32_FIND_DATA

InFile = InputBox$("Input volume spec", "Multiple Volumes",InPath)
' check for cancel button
If InFile = "" Then
  mv = %false
  Exit Function
End If
' is inFile a folder?
If Right$(InFile, 1) = "\" Then InFile = Left$(InFile, Len(InFile) - 1)
FindHandle = FindFirstFile (InFile, FileData)
x = FindClose (FindHandle)
If FileData.DwFileAttributes = %File_attribute_directory And FindHandle > 0 Then
  InPath = InFile
  InFolder = %true
  InFile = ""
Else
  fI = GetFilePath(InFile)
  If Len(fI) Then InPath = fI
End If
If Len(InPath) And Right$(InPath, 1) <> "\" Then InPath = InPath + "\"
If InFile = "" Or InStr(InFile, Any "\:") = 0 Then InFile = InPath + InFile
OutFile = InputBox$("Output volume spec", "Multiple volumes", OutPath)
' check for cancel button
If OutFile = "" Then
  mv = %false
  Exit Function
End If
' is OutFile a folder?
If Right$(OutFile, 1) = "\" Then OutFile = Left$(OutFile, Len(OutFile) - 1)
FindHandle = FindFirstFile (OutFile, FileData)
x = FindClose (FindHandle)
If FileData.DwFileAttributes = %File_attribute_directory And FindHandle > 0 Then
  OutPath = OutFile
  OutFolder = %true
  OutFile = ""
Else
  fI = getFilePath$(OutFile)
  If Len(fI) Then OutPath = fI
End If
If OutPath <> "" And Right$(OutPath, 1) <> "\" Then OutPath = OutPath + "\"
result = 6 ' default to combine output
If OutFolder = %true Then
  result = MsgBox("Combine volumes", %mb_IconQuestion Or %Mb_YesNo, " ")
 End If
If result = 6Then
  mv = 1
Else ' separate volumes
  mv = 2
End If
' check for wildcards
x = InStr(InFile, "*")
If x = %false Then
  fi = InFile + "*.BRF"
  If transOpt = 1 Then fi = InFile + "*.TXT"
Else
  InFolder = %true
End If
' how many volumes?
x = 0
fs = Dir$(fi)
Do While fs <> ""
  Incr x
  fs = Dir$
Loop
' setup volumes list
ReDim fls(0 to x) As Global String
x = 0
fs = Dir$(fi)
Do While fs <> ""
  fls(x) = UCase$(InPath + fs)
  If InFolder = %false Then ' select all files
    fls(x) = "|" + fls(x)
    If OutFile = "" Then
      Result = InStr(-1, fs, "\")
      fs = OutPath + Mid$(fs, Result + 1)
      OutFile = Left$(fs, Len(fs) - 4) + ".TXT"
      If transopt = 1 Then OutFile = Left$(fs, Len(fs) - 4) + ".BRF"
    End If
  End If
  Incr x
  fs = Dir$
Loop
Array Sort fls()
If fls(0) = "" Then Array Delete fls(0)
' show the files list if necessary
If InFolder = %true Then
  ShowFiles
  If InFile = "" Then
    mv = %false
    Exit Function
  End If
 End If
If mv = 3 Then ' emboss
  mv = 2
  MultiVolume 1
Else
  MultiVolume 0
End If
mv = %false
Control Set Focus hdlg, %Nfb_Input
End Function

Function PbMain() As Long

Local f As Long ' string marker
Local p As Long ' for progpath
Local style As Long, exStyle As Long ' dialog styles
Local hp As String ' for HomePath

' set defaults
Style& = %WS_POPUP Or %DS_MODALFRAME Or %WS_CAPTION Or %WS_MINIMIZEBOX Or %WS_SYSMENU Or %WS_MAXIMIZEBOX Or %DS_CENTER
ExStyle& = 0

' where's home?

p = GetCommandLine ()
hp$ = Peek$(p,255)
Replace Chr$(34) With "" In hp$
f& = InStr(hp$,Chr$(0))
hp$ = RTrim$(Left$(hp$, f&-1))
f = InStr(-1, hp$, "\")
HomePath$ = UCase$(Left$(hp$, f))

' read defaults
ReadInit
BuildParms
' create the primary gui

Dialog New 0, "WinTrans 2.0", , , 120, 115, style&, ExStyle& To hdlg

' add controls
Control Add Button, hdlg, %Nfb_Input, "&Input", 10, 10, 40, 14, &H10001 Call InputButton
Control Add Button, hdlg, %Nfb_Output, "&Output", 60, 10, 40, 14 Call OutputButton
Control Add Option, hdlg, %Nfb_BrfTxt, "&Braille to text", _
  20, 25, 80, 14,_
  %ws_group Or %Ws_Tabstop Or %BS_Multiline, %Ws_ex_ClientEdge
Control Add Option, hdlg, %Nfb_TxtBrf, "&Text to braille", _
  20, 40, 80, 14, _
  %BS_Multiline, %Ws_ex_ClientEdge
Control Add Button, hdlg, %Nfb_Ok, "T&ranslate", 10, 55, 40, 14 Call TransButton
Control Add Button, hdlg, %Nfb_Prt, "&Emboss", 60, 55, 40, 14 Call EmbossButton
Control Add Button, hdlg, %NFB_MultiVol, "&Volumes", 10, 70, 40, 14 Call mvProc
Control Add Button, hdlg, %NFB_Pref, "&Preferences", 60, 70, 40, 14 Call preferences
Control Add Button, hdlg, %Nfb_Stats, "&Stats", 10, 85, 40, 14 Call statsbutton
Control Add Button, hdlg, %Nfb_Cancel, "E&xit", 60, 85, 40, 14 Call CancelButton
Control Add Label, hdlg, %Nfb_Stat, "", 10, 100, 100, 8 ' status line
' set default radio button
If TransOpt = 0 Then Control Set Check hdlg, %Nfb_brftxt, 1
If TransOpt = 1 Then Control Set Check hdlg, %Nfb_txtbrf, 1
' focus on input button
Control Set Focus hdlg, %Nfb_Input
' display the dialog wait for exit
Dialog Show Modal hdlg
' cleanup
If Len(Dir$(HomePath + "TEMP.CNF")) Then Kill HomePath + "TEMP.CNF"
If Len(Dir$(HomePath + "TEMP.EFL")) Then Kill HomePath + "TEMP.EFL"
If Len(Dir$(HomePath + "TEMPBRL.TAB")) Then Kill HomePath + "TEMPBRL.TAB"
If Len(Dir$(HomePath + "TEMPBAK.TAB")) Then Kill HomePath + "TEMPBAK.TAB"
WriteInit
End Function
