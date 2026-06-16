' wintrans windows skin for NfbTrans.exe
' Author: Anonymous John
' Begin date: 10/06/02
' Language: pb win 7.0

' metastatements
#COMPILE EXE
#DIM ALL

' included files
#INCLUDE "win32api.inc" ' windows api
#INCLUDE "comdlg32.inc" ' common windows dialogs
#include "richedit.inc" ' for viewing stats

' constants
%Nfb_Input = %Wm_User + 1024 ' input file button
%Nfb_Output = %Nfb_Input + 1 ' output file button
%Nfb_BrfTxt = %Nfb_Output + 1 ' braille to text radio button
%Nfb_Txtbrf = %Nfb_brftxt + 1 ' text to braille radio button
%Nfb_Ok = %Nfb_txtbrf + 1 ' translate button
%Nfb_Cancel = %Nfb_Ok + 1 ' exit button
%Nfb_Stat = %Nfb_Cancel + 1 ' status line
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
%PrCancel = %PrOk + 1 ' cancel button
%Nfb_prt = %prCancel + 1 ' emboss button
%Nfb_Stats = %Nfb_Prt + 1 ' stats button
%StatsBox = %Nfb_Stats + 1 ' statistics text box
%StatsCancel = %StatsBox + 1 ' statistics close button

' declarations
DECLARE FUNCTION GetFile(Op AS LONG, title AS STRING) AS STRING ' filespec browse dialog
DECLARE CALLBACK FUNCTION InputButton() AS LONG
DECLARE CALLBACK FUNCTION OutputButton() AS LONG
DECLARE CALLBACK FUNCTION MvProc() AS LONG
DECLARE CALLBACK FUNCTION preferences() AS LONG
DECLARE CALLBACK FUNCTION StatsButton() AS LONG
DECLARE CALLBACK FUNCTION TransButton() AS LONG
DECLARE CALLBACK FUNCTION EmbossButton() AS LONG
DECLARE CALLBACK FUNCTION prApply() AS LONG
DECLARE CALLBACK FUNCTION PrOk() AS LONG
DECLARE CALLBACK FUNCTION CancelButton() AS LONG
DECLARE SUB StatLine
DECLARE SUB MultiVolume(op AS LONG)
DECLARE SUB ReadInit
DECLARE SUB WriteInit
DECLARE FUNCTION BuildParms() AS LONG
DECLARE SUB getPrefs(hdlgn AS LONG)

' globals
GLOBAL hdlg AS LONG ' dialog handle
GLOBAL mv AS LONG ' multi volume checkstate
global TransOpt AS LONG ' b2t or t2b
GLOBAL InFile AS STRING ' input filespec
GLOBAL OutFile AS STRING ' Output filespec
GLOBAL HomePath AS STRING ' home dir for wintrans.exe
GLOBAL pw AS STRING ' page width
GLOBAL pl AS STRING ' page length
GLOBAL ps AS STRING ' starting page
GLOBAL pE AS STRING ' ending page
GLOBAL ls AS STRING ' page break
GLOBAL ip AS STRING ' interpoint toggle
GLOBAL lm AS STRING ' left margin
GLOBAL co AS STRING ' copies
GLOBAL cf AS STRING ' config file
GLOBAL le AS STRING ' format file
GLOBAL tf AS STRING ' table files
GLOBAL so AS STRING ' sound on/off
GLOBAL cp AS STRING ' extra command line parms
global pr as string ' embosser port
GLOBAL parms AS STRING ' command line parameters for nfbtrans

SUB ReadInit ' reads wintrans.ini
LOCAL f AS LONG ' marker
LOCAL f1 AS LONG ' marker
LOCAL nFile AS LONG ' file handle
LOCAL buf AS STRING ' file buffer

nFile = FREEFILE
OPEN HomePath + "WINTRANS.INI" FOR BINARY AS #nFile
GET$ #nFile, LOF(nFile), buf
CLOSE #nFile
buf = LCASE$(buf)
f = INSTR(buf, "pw=") + 3
f1 = INSTR(f, buf, " ")
pw = MID$(buf, f, f1 - f)
f = INSTR(buf, "pl=") + 3
f1 = INSTR(f, buf, " ")
pl = MID$(buf, f, f1 - f)
f = INSTR(buf, "ps=") + 3
f1 = INSTR(f, buf, " ")
ps = MID$(buf, f, f1 - f)
f = INSTR(buf, "pe=") + 3
f1 = INSTR(f, buf, " ")
pE = MID$(buf, f, f1 - f)
f = INSTR(buf, "ls=") + 3
f1 = INSTR(f, buf, " ")
ls = MID$(buf, f, f1 - f)
f = INSTR(buf, "ip=") + 3
f1 = INSTR(f, buf, " ")
ip = MID$(buf, f, f1 - f)
f = INSTR(buf, "lm=") + 3
f1 = INSTR(f, buf, " ")
lm = MID$(buf, f, f1 - f)
f = INSTR(buf, "co=") + 3
f1 = INSTR(f, buf, " ")
co = MID$(buf, f, f1 - f)
f = INSTR(buf, "cf=") + 3
f1 = INSTR(f, buf, " ")
cf = MID$(buf, f, f1 - f)
f = INSTR(buf, "le=") + 3
f1 = INSTR(f, buf, " ")
le = MID$(buf, f, f1 - f)
f = INSTR(buf, "tf=") + 3
f1 = INSTR(f, buf, " ")
tf = MID$(buf, f, f1 - f)
f = INSTR(buf, "so=") + 3
f1 = INSTR(f, buf, " ")
so = MID$(buf, f, f1 - f)
f = INSTR(buf, "cp=") + 3
f1 = INSTR(f, buf, " ")
cp = MID$(buf, f, f1 - f)
f = INSTR(buf, "pr=") + 3
f1 = INSTR(f, buf, " ")
pr = MID$(buf, f, f1 - f)
END SUB

SUB WriteInit ' writes wintrans.ini
LOCAL nFile AS LONG ' file handle
nFile = FREEFILE
OPEN HomePath + "WINTRANS.INI" FOR OUTPUT AS #nFile
PRINT #nFile, "lm=" + lm + " left margin"
PRINT #nFile, "co=" + co + " copies"
PRINT #nFile, "pw=" + pw + " page width"
PRINT #nFile, "pl=" + pl + " page length"
PRINT #nFile, "ps=" + ps + " starting page"
PRINT #nFile, "pE=" + pE + " ending page"
PRINT #nFile, "ls=" + ls + " page break"
PRINT #nFile, "ip=" + ip + " interpoint"
PRINT #nFile, "so=" + so + " sound"
PRINT #nFile, "cf=" + cf + " config file"
PRINT #nFile, "le=" + le + " format file"
PRINT #nFile, "tf=" + tf + " translation table file"
PRINT #nFile, "cp=" + cp + " command line parms"
print #nFile, "pr=" + pr + " embosser port"
CLOSE #nFile
END SUB

FUNCTION BuildParms() AS LONG  ' builds command line
LOCAL f AS LONG ' marker
LOCAL fi AS STRING ' parsing file buffer
parms = "nc=1"
IF lm <> "1" THEN parms = parms + " lm=" + lm
IF co <> "1" THEN parms = parms + " co=" + co
IF pw <> "40" THEN parms = parms + " pw=" + pw
IF pl <> "25" THEN parms = parms + " pl=" + pl
IF ps <> "1" THEN parms = parms + " ps=" + ps
IF pE <> "9999" THEN parms = parms + " pE=" + pE
IF ls <> "99" THEN parms = parms + " ls=" + ls
IF ip <> "0" THEN parms = parms + " ip=" + ip
IF so <> "0" THEN parms = parms + " so=" + so
IF cf <> "" THEN
  ' does it exist?
  IF LEN(DIR$(cf)) THEN
    FILECOPY cf, HomePath + "TEMP.CNF"
    parms = parms + " cf=TEMP.CNF"
  ELSE
    MSGBOX "Can't find " + cf + "using nfbtrans.cnf", %mb_IconHand, "Config file error"
    FUNCTION = %true
  END IF
END IF
IF le <> "" THEN
  IF LEN(DIR$(le)) THEN
    parms = parms + " le=" + le
    parms = parms + " ef=1"
  ELSE
    MSGBOX "Can't find " + le + " None used", %mb_IconHand, "Format file error"
    FUNCTION = %true
  END IF
END IF
IF tf <> "" THEN
  ' fill out the spec
  IF LEFT$(tf, 1) = ";" THEN tf = "BRAILLE.TAB" + tf
  IF RIGHT$(tf, 1) = ";" THEN tf = tf + "BACK.TAB"
  IF INSTR(tf, ";") = 0 THEN tf = tf + ";BACK.TAB"
  ' do they exist?
  f = INSTR(tf, ";")
  fi = LEFT$(tf, f - 1) ' braille table
  IF DIR$(fi) = "" THEN
    MSGBOX "Can't find " + fi + " Using BRAILLE.TAB", %mb_IconHand, "Table file error
    tf = "BRAILLE.TAB" + MID$(tf, f)
    f = INSTR(tf, ";")
    FUNCTION = %true
 ELSE
   FILECOPY fi, HomePath + "TEMPBRL.TAB"
  END IF
  fi = MID$(tf, f + 1)
  IF DIR$(fi) = "" THEN
    MSGBOX "Can't find " + fi + " using BACK.TAB", %mb_IconHand, "Table file error"
    tf = LEFT$(tf, f) + "BACK.TAB"
    FUNCTION = %true
  ELSE
    FILECOPY fi, HomePath + "TEMPBAK.TAB"
  END IF
  parms = parms + " tf=tempbrl.tab;tempbak.tab"
END IF
IF cp <> "" THEN parms = parms + " cp=" + cp
parms = parms + " "
END FUNCTION

SUB getPrefs(hdlgn AS LONG)

LOCAL lc AS LONG ' checkbox state
' get info from text boxes
CONTROL GET TEXT hdlgn, %prlm TO lm
CONTROL GET TEXT hdlgn, %prcO TO co
CONTROL GET TEXT hdlgn, %prpw TO pw
CONTROL GET TEXT hdlgn, %prpl TO pl
CONTROL GET TEXT hdlgn, %prps TO ps
CONTROL GET TEXT hdlgn, %prpE TO pE
CONTROL GET TEXT hdlgn, %prcf TO cf
CONTROL GET TEXT hdlgn, %prlE TO lE
CONTROL GET TEXT hdlgn, %prtf TO tf
CONTROL GET TEXT hdlgn, %prcp TO cp
CONTROL GET TEXT hdlgn, %prpr TO pr
' check our check boxes
CONTROL GET CHECK hdlgn, %prIp TO lc
ip = LTRIM$(STR$(lc))
CONTROL GET CHECK hdlgn, %prsO TO lc
so = "0"
IF lc = 1 THEN sO = "15"
' radio buttons
CONTROL GET CHECK hdlgn, %prff TO lc
IF lc = 1 THEN ls = "99"
CONTROL GET CHECK hdlgn, %prvt TO lc
IF lc = 1 THEN ls = "999"
CONTROL GET CHECK hdlgn, %prbl TO lc
IF lc = 1 THEN ls = "0"
END SUB

SUB StatLine ' updates the status line

LOCAL msg AS STRING
msg = "Input: " + InFile + " Output: " + OutFile
CONTROL SET TEXT hdlg, %NFB_Stat, msg
END SUB

FUNCTION GetFile(Op AS LONG, Title AS STRING) AS STRING
'
' sFileName = File name to read
' title = string for dialog title

LOCAL sFileName AS STRING
LOCAL fStyle AS LONG
LOCAL a%
LOCAL result AS LONG
' LOCAL hctl AS LONG
LOCAL msg AS ASCIIZ * 255
LOCAL path AS STRING ' starting path
LOCAL CrtDir AS STRING ' current directory

ON ERROR RESUME NEXT
Path = HomePath
crtDir = MID$(HomePath, 3, LEN(HomePath) - 3)
sFileName = "*.*"
IF op = 0 THEN ' input file
  fStyle = %OFN_FILEMUSTEXIST OR %OFN_LONGNAMES
  IF mv > 0 THEN fStyle = %OFN_LONGNAMES
ELSE ' output file
  fStyle = %OFN_OVERWRITEPROMPT OR %OFN_LONGNAMES
END IF
a% = OpenFileDialog(hdlg, title, sFileName, Path, _
  "All Files|*.*", "*", BYCOPY fStyle)
IF sFileName = "*.*" THEN sFileName = "" ' dialog canceled
REPLACE ".*" WITH "" IN sFileName
IF sFileName <> "" THEN
  IF INSTR(sFileName, ANY "\:") = 0 THEN
    path = UCASE$(HomePath$)
    sFileName = UCASE$(path + sFilename)
  END IF
END IF
CHDIR crtDir
FUNCTION = sFileName
END FUNCTION

SUB MultiVolume(op AS LONG)
' 0 = translate 1 = emboss
LOCAL brf AS STRING ' temp braille filespec
LOCAL txt AS STRING ' temp text file
LOCAL dos AS STRING ' shell commandSpec
LOCAL fi AS STRING ' brf FileSpec
LOCAL fs AS STRING ' folder search
LOCAL fp AS STRING ' folder path
LOCAL fls() AS STRING ' volumes list
LOCAL x AS LONG ' counter

' setup temps
brf = HomePath + "TEMP.BRF"
txt = HomePath + "TEMP.TXT"
' get inFile path
x = INSTR(-1, InFile, "\")
IF x > 0 THEN fp = LEFT$(InFile, x)
x = 0
' get file mask
fi = InFile + "*.BRF"
if transOpt = 1 then fi = InFile + "*.TXT"
' how many volumes?
fs = DIR$(fi)
DO WHILE fs <> ""
  INCR x
  fs = DIR$
LOOP
' setup volumes list
DIM fls(1: x) AS STRING
x = 0
fs = DIR$(fi)
DO WHILE fs <> ""
  INCR x
  fls(x) = fp + fs
  fs = DIR$
LOOP
' process he input files
if TransOpt = 0 then ' braille to text
  ' copy to short file names
  FOR x = 1 TO UBOUND(fls)
    fs = fls(x)
    brf = HomePath + "TEMP" + RIGHT$("00" + LTRIM$(STR$(x)), 3) + ".BRF"
    FILECOPY fs, BRF
  NEXT x
  IF op = 0 THEN ' translate
    IF mv = 1 THEN ' combine output
      dos = ENVIRON$("COMSPEC") + " /C " + HomePath + "nfbtrans.exe tm=13 " + parms
      dos = dos + "TEMP*.BRF >" + txt
      SHELL dos$, 6
      FILECOPY txt, OutFile
    END IF
    IF mv = 2 THEN ' separate output files
      FOR x = 1 TO UBOUND(fls)
        dos = ENVIRON$("COMSPEC") + " /C " + HomePath + "nfbtrans.exe tm=13 " + parms
        fs = fls(x)
        brf = HomePath + "TEMP" + RIGHT$("00" + LTRIM$(STR$(x)), 3) + ".BRF"
        OutFile = LEFT$(fs, LEN(fs) - 4) + ".txt"
        Dos = dos + brf + " >" + txt
        SHELL dos$, 6
        FILECOPY txt, OutFile
      NEXT x
    END IF
  END IF
  IF Op = 1 THEN ' emboss
    FOR x = 1 TO UBOUND(fls)
      brf = HomePath + "TEMP" + RIGHT$("00" + LTRIM$(STR$(x)), 3) + ".BRF"
      FILECOPY BRF, PR
    NEXT x
  END IF
  ' kill the temp files
  FOR x = 1 TO UBOUND(fls)
    brf = HomePath + "TEMP" + RIGHT$("00" + LTRIM$(STR$(x)), 3) + ".BRF"
    KILL brf
  NEXT x
end if
if TransOpt = 1 then ' text to braille
  ' copy to short file names
  FOR x = 1 TO UBOUND(fls)
    fs = fls(x)
    txt = HomePath + "TEMP" + RIGHT$("00" + LTRIM$(STR$(x)), 3) + ".TXT"
    FILECOPY fs, txt
  NEXT x
  IF op = 0 THEN ' translate
    IF mv = 1 THEN ' combine output
      dos = ENVIRON$("COMSPEC") + " /C " + HomePath + "nfbtrans.exe tm=11 " + parms
      dos = dos + "TEMP*.TXT >" + brf
      SHELL dos$, 6
      FILECOPY brf, OutFile
    END IF
    IF mv = 2 THEN ' separate output files
      FOR x = 1 TO UBOUND(fls)
        dos = ENVIRON$("COMSPEC") + " /C " + HomePath + "nfbtrans.exe tm=11 " + parms
        fs = fls(x)
        txt = HomePath + "TEMP" + RIGHT$("00" + LTRIM$(STR$(x)), 3) + ".TXT"
        OutFile = LEFT$(fs, LEN(fs) - 4) + ".BRF"
        Dos = dos + txt + " >" + brf
        SHELL dos$, 6
        FILECOPY brf, OutFile
      NEXT x
    END IF
  END IF
  IF Op = 1 THEN ' emboss
    brf = HomePath + "EMBOSS.BRF"
    FOR x = 1 TO UBOUND(fls)
      txt = HomePath + "TEMP" + RIGHT$("00" + LTRIM$(STR$(x)), 3) + ".TXT"
      dos = ENVIRON$("COMSPEC") + " /C " + HomePath + "nfbtrans.exe tm=11 " + parms
      dos = dos + " " + TXT + " >" + brf
      SHELL dos$, 6
      FILECOPY BRF, PR
      kill BRF
    NEXT x
  END IF
  ' kill the temp files
  FOR x = 1 TO UBOUND(fls)
    txt = HomePath + "TEMP" + RIGHT$("00" + LTRIM$(STR$(x)), 3) + ".TXT"
    KILL txt
  NEXT x
end if
END SUB

CALLBACK FUNCTION InputButton() AS LONG
InFile = GetFile(0, "Input File")
StatLine
END FUNCTION

CALLBACK FUNCTION OutputButton() AS LONG
LOCAL f AS LONG
LOCAL fln AS STRING
DO
  fln = GetFile(1, "Output file")
  IF LEN(DIR$(fln)) THEN
    f = MSGBOX("Overwrite " + OutFile, %Mb_IconQuestion OR %Mb_YesNo, "File exists")
    IF f <> 6 THEN fln = ""
  END IF
LOOP UNTIL LEN(fln)
OutFile = fln
StatLine
CONTROL SET FOCUS hdlg, %Nfb_Output
END FUNCTION

CALLBACK FUNCTION statsButton() AS LONG
LOCAL style&, exStyle& ' dialog styles
LOCAL hdlgs AS LONG ' dialog handle
LOCAL nFile AS LONG ' file handle
local cs1 as long ' text selection
local cs2 as long ' text selection
local buffer as string ' text buffer

Style& = %WS_POPUP OR %DS_MODALFRAME OR %WS_CAPTION OR %WS_MINIMIZEBOX OR %WS_SYSMENU OR %WS_MAXIMIZEBOX OR %DS_CENTER
ExStyle& = 0
nFile = FREEFILE
cs1 = 1
cs2 = 1

' get the text
IF DIR$(HomePath + "stat.txt") = "" THEN
  MSGBOX "No statistics available.", %mb_IconAsterisk, "No stats"
  EXIT FUNCTION
END IF
OPEN HomePath + "stat.txt" FOR BINARY AS #nFile
GET$ #nFile, LOF(nFile), buffer
CLOSE #nFile
' setup the dialog
DIALOG NEW CBHNDL, "NfbTrans Statistics", , , 350, 265, style&, exStyle& TO hdlgs
' controls
CONTROL ADD TEXTBOX, hdlgs, %Statsbox, buffer, 25, 15, 300, 235, _
  %WS_CHILD OR %WS_VISIBLE OR %es_WantReturn OR _
  %ES_READONLY OR %ES_MULTILINE OR _
  %ES_AUTOVSCROLL OR %ES_NOHIDESEL or %Es_SaveSel, 0
CONTROL ADD BUTTON, hdlgs, %StatsCancel, "&Close", 150, 250, 40, 14 CALL CancelButton
control set focus hdlgs, %StatsBox
DIALOG SHOW MODAL hdlgs
END FUNCTION

CALLBACK FUNCTION TransButton() AS LONG

LOCAL f AS LONG ' string marker
LOCAL dos AS STRING' shell to NfbTrans
LOCAL brf AS STRING ' temp braille filespec
LOCAL txt AS STRING ' temp text filespec

' setup temp filespecs
txt = HomePath + "TEMP.TXT"
brf = HomePath + "TEMP.BRF"
' start a new stat.txt file
IF DIR$(HomePath + "STAT.TXT") <> "" THEN KILL HomePath + "STAT.TXT"
' Which radio button is checked?
CONTROL GET CHECK hdlg, %Nfb_TxtBrf TO TransOpt
' have we got everything?
IF InFile = "" THEN
  MSGBOX "No Input file defined", %Mb_IconHand, "File error"
  EXIT FUNCTION
END IF
' don't translate .brf to braille or .txt to text
IF TransOpt = 0 AND LCASE$(RIGHT$(InFile, 4)) = ".txt" THEN
  MSGBOX InFile + " cannot be translated to text", %Mb_IconHand, "Entry error"
  EXIT FUNCTION
END IF
IF TransOpt = 1 AND LCASE$(RIGHT$(InFile, 4)) = ".brf" THEN
  MSGBOX InFile + " cannot be translated to braille", %Mb_IconHand, "Entry error"
  EXIT FUNCTION
END IF
f = INSTR(InFile, ".")
IF f = 0 THEN f = LEN(InFile) + 1
IF OutFile = "" THEN
  IF TransOpt = 0 THEN
    OutFile = LEFT$(InFile, f - 1) + ".TXT"
  ELSE
    OutFile = LEFT$(InFile, f - 1) + ".BRF"
  END IF
END IF
IF LCASE$(Infile) = LCASE$(OutFile) THEN
  MSGBOX "file " + InFile + " cannot be translated onto itself", %Mb_IconHand, "Error"
  EXIT FUNCTION
END IF
' ensure short file names for NfbTrans and prepare shell statement
IF mv > 0 THEN
  MultiVolume 0
  mv = %false
  CONTROL ENABLE hdlg, %Nfb_Input
ELSE
  dos = ENVIRON$("COMSPEC") + " /C "
  IF TransOpt = 0 THEN
    FILECOPY InFile, BRF
    Dos = dos + HomePath + "nfbtrans.exe tm=13 " + parms + brf + " >" + txt
  ELSE
    FILECOPY InFile, TXT
    Dos = dos + HomePath + "nfbtrans.exe tm=11 " + parms + txt + " >" + brf
  END IF
END IF
' update status line
CONTROL SET TEXT hdlg, %Nfb_Stat, "Translating"
SHELL dos$, 6
StatLine
' check for errors
' kill empty files
OPEN txt FOR BINARY AS #1
f = LOF(1)
CLOSE 1
IF f = 0 THEN KILL txt
OPEN brf FOR BINARY AS #1
f = LOF(1)
CLOSE 1
IF f = 0 THEN KILL brf
IF TransOpt = 0 THEN
  IF DIR$(txt) = "" THEN
    MSGBOX InFile + " to " + OutFile, %Mb_IconHand, "Translation error"
  ELSE
    MSGBOX InFile + " to " + OutFile, %Mb_IconAsterisk OR %Mb_SetForeground, "Translated"
  END IF
END IF
IF TransOpt = 1 THEN
  IF DIR$(brf) = "" THEN
    MSGBOX InFile + " to " + OutFile, %Mb_IconHand, "Translation error"
  ELSE
    MSGBOX InFile + " to " + OutFile, %Mb_IconAsterisk, "Translated"
  END IF
END IF
' cleanup
IF TransOpt = 0 AND DIR$(txt) <> "" THEN
  FILECOPY txt, OutFile
  KILL txt
  KILL brf
END IF
IF TransOpt = 1 AND DIR$(brf) <> "" THEN
  FILECOPY brf, OutFile
  KILL brf
  KILL txt
END IF
OutFile = ""
CONTROL SET FOCUS hdlg, %Nfb_Input
END FUNCTION

CALLBACK FUNCTION EmbossButton() AS LONG

LOCAL f AS LONG ' string marker
LOCAL TransOpt AS LONG ' b2t or t2b
LOCAL dos AS STRING' shell to NfbTrans
LOCAL brf AS STRING ' temp braille filespec
LOCAL txt AS STRING ' temp text filespec

' setup temp filespecs
txt = HomePath + "TEMP.TXT"
brf = HomePath + "TEMP.BRF"
OutFile = ""
' Which radio button is checked?
CONTROL GET CHECK hdlg, %Nfb_TxtBrf TO TransOpt
' have we got everything?
IF InFile = "" THEN
  MSGBOX "No Input file defined", %Mb_IconHand, "File error"
  EXIT FUNCTION
END IF
f = INSTR(InFile, ".")
IF f = 0 THEN f = LEN(InFile) + 1
IF OutFile = "" THEN
  IF TransOpt = 0 THEN
    OutFile = LEFT$(InFile, f - 1) + ".TXT"
  ELSE
    OutFile = LEFT$(InFile, f - 1) + ".BRF"
  END IF
END IF
' update status line
CONTROL SET TEXT hdlg, %Nfb_Stat, "Embossing"
' ensure short file names for NfbTrans and prepare shell statement
IF mv > 0 THEN
  MultiVolume 1
  mv = %false
  CONTROL ENABLE hdlg, %Nfb_Input
ELSE
  dos = ENVIRON$("COMSPEC") + " /C " + HomePath + "NFBTRANS.EXE "
  IF LCASE$(RIGHT$(InFile, 4)) = ".brf" THEN
    FILECOPY InFile, PR
  end if
  IF LCASE$(RIGHT$(InFile, 4)) = ".txt" THEN
    FILECOPY InFile, TXT
    Dos = dos + "tm=13 " + parms + txt + " >" + brf
    SHELL dos$, 6
    filecopy brf, PR
  end if
END IF
StatLine
' kill temp files
if len(dir$(BRF)) then KILL brf
if len(dir$(txt)) then kill txt
' cleanup
CONTROL SET FOCUS hdlg, %Nfb_Input
END FUNCTION

CALLBACK FUNCTION CancelButton() AS LONG
DIALOG END CBHNDL, 0
END FUNCTION

CALLBACK FUNCTION PrApply() AS LONG
GetPrefs(CBHNDL)
BuildParms
END FUNCTION

CALLBACK FUNCTION PrOk() AS LONG
GetPrefs(CBHNDL)
BuildParms
WriteInit
END FUNCTION

CALLBACK FUNCTION preferences() AS LONG

LOCAL hdlgp AS LONG ' dialog handle
LOCAL style&, exStyle&
Style& = %WS_POPUP OR %DS_MODALFRAME OR %WS_CAPTION OR %WS_MINIMIZEBOX OR %WS_SYSMENU OR %WS_MAXIMIZEBOX OR %DS_CENTER
ExStyle& = 0
' setup the dialog
DIALOG NEW CBHNDL, "Wintrans Preferences", , , 350, 130, style&, exStyle& TO hdlgp
' controls
CONTROL ADD LABEL, hdlgp, -1, "Left &Margin", 10, 10, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prlm, lm, 60, 10, 100, 12, %Es_autohscroll OR %Ws_Tabstop
CONTROL ADD LABEL, hdlgp, -1, "&Copies", 175, 10, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prcO, co, 235, 10, 100, 12, %Es_Autohscroll OR %Ws_Tabstop
CONTROL ADD LABEL, hdlgp, -1, "Page &Width", 10, 25, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prpw, pw, 60, 25, 100, 12, %es_Autohscroll OR %ws_tabstop
CONTROL ADD LABEL, hdlgp, -1, "Page &len", 175, 25, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prpl, pl, 235, 25, 100, 12, %Es_Autohscroll OR %Ws_TabStop
CONTROL ADD LABEL, hdlgp, -1, "&Start page", 10, 40, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prps, ps, 60, 40, 100, 12, %Es_Autohscroll OR %ws_Tabstop
CONTROL ADD LABEL, hdlgp, -1, "&End page", 175, 40, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prpE, pE, 235, 40, 100, 12, %Es_Autohscroll OR %Ws_Tabstop
CONTROL ADD LABEL, hdlgp, -1, "Page Break", 25, 55, 40, 8
CONTROL ADD OPTION, hdlgp, %prff, "&FormFeed", 75, 55, 40, 14,_
  %ws_group OR %Bs_MultiLine, %Ws_ex_ClientEdge
CONTROL ADD OPTION, hdlgp, %prvt, "&V-Tab", 125, 55, 40, 14,_
  %BS_MultiLine, %Ws_ex_ClientEdge
CONTROL ADD OPTION, hdlgp, %prbl, "&Blanks", 175, 55, 40, 14,_
  %BS_MultiLine, %Ws_ex_ClientEdge
CONTROL ADD CHECKBOX, hdlgp, %PrIp, "&Interpoint", 225, 55, 40, 14
CONTROL ADD CHECKBOX, hdlgp, %prsO, "So&und", 275, 55, 40, 14
CONTROL ADD LABEL, hdlgp, -1, "Co&nfig file", 10, 70, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prcf, cf, 60, 70, 100, 12, %Es_Autohscroll OR %Ws_Tabstop
CONTROL ADD LABEL, hdlgp, -1, "Fo&rmat file", 175, 70, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prlE, le, 235, 70, 100, 12, %Es_Autohscroll OR %Ws_Tabstop
CONTROL ADD LABEL, hdlgp, -1, "&Table file", 10, 85, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prtf, tf, 60, 85, 100, 12, %Es_Autohscroll OR %Ws_Tabstop
CONTROL ADD LABEL, hdlgp, -1, "Port", 175, 85, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prpr, pr, 235, 85, 100, 12, %Es_Autohscroll OR %Ws_Tabstop
CONTROL ADD LABEL, hdlgp, -1, "Comman&ds", 10, 100, 40, 8
CONTROL ADD TEXTBOX, hdlgp, %prcp, cp, 60, 100, 100, 12, %Es_Autohscroll OR %Ws_Tabstop
CONTROL ADD BUTTON, hdlgp, %PrApply, "&Apply", 100, 115, 40, 14 CALL PrApply
CONTROL ADD BUTTON, hdlgp, %prOk, "&Ok", 150, 115, 40, 14 CALL PrOk
CONTROL ADD BUTTON, hdlgp, %PrCancel, "E&Xit", 200, 115, 40, 14 CALL CancelButton
' set checkboxes and radio buttons
IF so <> "0" THEN CONTROL SET CHECK hdlgp, %PrsO, 1
IF ip = "1" THEN CONTROL SET CHECK hdlgp, %prIp, 1
' radio buttons
IF ls = "99" THEN CONTROL  SET CHECK hdlgp, %prff, 1
IF ls = "999" THEN CONTROL SET CHECK hdlgp, %prvt, 1
IF ls = "0" THEN CONTROL SET CHECK hdlgp, %prbl, 1
' show the dialog
DIALOG SHOW MODAL hdlgp
END FUNCTION

CALLBACK FUNCTION MvProc() AS LONG
LOCAL result AS LONG ' combine or not msgBox
' toggle
IF mv  = %true THEN
  mv = %false
  CONTROL ENABLE CBHNDL, %nfb_Input
  MSGBOX "Multiple volumes off", %mb_IconExclamation, " "
  EXIT FUNCTION
END IF
InFile = INPUTBOX$("Volume spec", "Multiple Volumes",InFile)
' current dir if no path
IF INSTR(InFile, ANY "\:") = 0 THEN InFile = HomePath + InFile
mv = 1
result = MSGBOX("Combine output files", %mb_IconQuestion OR %Mb_YesNo, " ")
IF result <> 6 THEN mv = 2
CONTROL DISABLE hdlg, %nfb_Input
CONTROL SET FOCUS hdlg, %Nfb_Ok
END FUNCTION

FUNCTION PBMAIN() AS LONG

LOCAL f AS LONG ' string marker
LOCAL p AS LONG ' for progpath
LOCAL style AS LONG, exStyle AS LONG ' dialog styles
LOCAL hp AS STRING ' for HomePath

' set defaults
Style& = %WS_POPUP OR %DS_MODALFRAME OR %WS_CAPTION OR %WS_MINIMIZEBOX OR %WS_SYSMENU OR %WS_MAXIMIZEBOX OR %DS_CENTER
ExStyle& = 0

' where's home?

p = GetCommandLine ()
hp$ = PEEK$(p,255)
REPLACE CHR$(34) WITH "" IN hp$
f& = INSTR(hp$,CHR$(0))
hp$ = RTRIM$(LEFT$(hp$, f&-1))
f = INSTR(-1, hp$, "\")
HomePath$ = UCASE$(LEFT$(hp$, f))

' read defaults
ReadInit
BuildParms
' create the primary gui

DIALOG NEW 0, "WinTrans 1.0", , , 120, 115, style&, ExStyle& TO hdlg

' add controls
CONTROL ADD BUTTON, hdlg, %Nfb_Input, "&Input", 10, 10, 40, 14, &H10001 CALL InputButton
CONTROL ADD BUTTON, hdlg, %Nfb_Output, "&Output", 60, 10, 40, 14 CALL OutputButton
CONTROL ADD OPTION, hdlg, %Nfb_BrfTxt, "&Braille to text", _
  20, 25, 80, 14,_
  %ws_group OR %Ws_Tabstop OR %BS_Multiline, %Ws_ex_ClientEdge
CONTROL ADD OPTION, hdlg, %Nfb_TxtBrf, "&Text to braille", _
  20, 40, 80, 14, _
  %BS_Multiline, %Ws_ex_ClientEdge
CONTROL ADD BUTTON, hdlg, %Nfb_Ok, "T&ranslate", 10, 55, 40, 14 CALL TransButton
CONTROL ADD BUTTON, hdlg, %Nfb_Prt, "&Emboss", 60, 55, 40, 14 CALL EmbossButton
CONTROL ADD BUTTON, hdlg, %NFB_MultiVol, "&Volumes", 10, 70, 40, 14 CALL mvProc
CONTROL ADD BUTTON, hdlg, %NFB_Pref, "&Preferences", 60, 70, 40, 14 CALL preferences
CONTROL ADD BUTTON, hdlg, %Nfb_Stats, "&Stats", 10, 85, 40, 14 CALL statsbutton
CONTROL ADD BUTTON, hdlg, %Nfb_Cancel, "E&xit", 60, 85, 40, 14 CALL CancelButton
CONTROL ADD LABEL, hdlg, %Nfb_Stat, "", 10, 100, 100, 8 ' status line
' set default radio button
CONTROL SET CHECK hdlg, %Nfb_brftxt, 1
' focus on input button
CONTROL SET FOCUS hdlg, %Nfb_Input
' display the dialog wait for exit
DIALOG SHOW MODAL hdlg
' cleanup
IF LEN(DIR$(HomePath + "TEMP.CNF")) THEN KILL HomePath + "TEMP.CNF"
IF LEN(DIR$(HomePath + "TEMP.EFL")) THEN KILL HomePath + "TEMP.EFL"
IF LEN(DIR$(HomePath + "TEMPBRL.TAB")) THEN KILL HomePath + "TEMPBRL.TAB"
IF LEN(DIR$(HomePath + "TEMPBAK.TAB")) THEN KILL HomePath + "TEMPBAK.TAB"
END FUNCTION
