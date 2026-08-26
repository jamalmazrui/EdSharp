//EdSharp 5.0
// June 16, 2026
//Copyright 2007 - 2026 by Jamal Mazrui
// GNU Lesser General Public License (LGPL)

using Microsoft.VisualBasic.ApplicationServices;
using System.Windows.Automation.Provider;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Tektosyne.NetMail ;
using Tektosyne.Win32Api;
using Homer;

[assembly: AssemblyTitle("EdSharp")]
[assembly: AssemblyProduct("EdSharp")]
// Single-sourced from AppVersion in EdSharp_setup.iss: BuildEdSharp.cmd generates
// Version.cs (BuildVersion.Version) from it.  This replaces the wildcard "5.0.*",
// which made the assembly version a build TIMESTAMP unrelated to the release --
// so the program, the installer, and the release tag could never agree.
// AssemblyFileVersion also stamps the Win32 version resource, so the version is
// visible in Explorer's file properties and readable by tools.
[assembly: AssemblyVersion(BuildVersion.Version)]
[assembly: AssemblyFileVersion(BuildVersion.Version)]
[assembly: AssemblyDescription("EdSharp editor")]
[assembly: AssemblyCompany("EmpowermentZone.com")]
[assembly: AssemblyCopyright("Copyright 2007 - 2026 by Jamal Mazrui")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCulture("")]

namespace EdSharp {
public class App : WindowsFormsApplicationBase {
// Dotted-numeric version used by the Elevate Version command to compare with
// the latest GitHub release tag.  Bump this on each release; the About dialog
// shows the friendly "5.0 beta" label separately.
// The version number is NOT stored here.  It lives in exactly one place --
// AppVersion in EdSharp_Setup.iss -- and BuildEdSharp.cmd generates Version.cs
// from it at build time, defining BuildVersion.Version.  That makes it
// impossible for the running program and the installer (and therefore the
// release tag that F11 compares against) to disagree.
public const string VersionString = BuildVersion.Version;
public static App Shell;
public static MdiFrame Frame;
public static string ProgramName;
public static string NetDir;
public static string ProgramDir;
public static string DataDir;
public static string DefaultIniFile;
public static string HotkeyIniFile;
public static string IniFile;
public static string LogFile = "";
public static string SpellCheckError = "";
public static string IndentModeFile;
public static string TempFile;
public static List<string> TempFiles = new List<string>();
//public static object Word = null;
public static object Boo = null;
public static object JAWS = null;
public static bool WordCreated = false;
public static bool ExtraSpeech = true;
public static bool IndentChange = true;
public static bool CaptureOutput = false;
public static string SpeechLog;
public static string MatchChunk = @"\s+";
public static string MatchParagraph = @"\n(\s*\n)+\s*";
public static string MatchSentence = @"([.?!]\s+)|(" + MatchParagraph + ")";
public static Dictionary<string, int> BomDictionary = null;

[STAThread]
public static void Main(string[] cmdLineArgs) {
// Installer Finish-page option: "EdSharp.exe --install-jaws-settings" copies
// EdSharp's JAWS settings family into every installed JAWS version and
// compiles them there, then reports and exits without launching the editor.
// (DbDo-style: the compile logic is here in C#, invoked from the installer's
// [Run] section, not done silently in the installer script.)
foreach (string sArg in cmdLineArgs) {
if (sArg.Equals("--install-jaws-settings", StringComparison.OrdinalIgnoreCase)
 || sArg.Equals("/install-jaws-settings", StringComparison.OrdinalIgnoreCase)) {
int iCopied, iCompiled;
string sScriptsDir = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "Scripts");
string sReport = JawsScripts.install(sScriptsDir, out iCopied, out iCompiled);
MessageBox.Show(sReport, "EdSharp JAWS scripts: " + iCopied + " copied, " + iCompiled + " compiled");
return;
}
}
// Runtime log, in the Homer Tools convention: one file per session,
// EdSharp_<timestamp>.log, beside the setup log. It opens with the
// version and environment, and Util.Log adds a line for every outside
// command EdSharp runs, with its exit code -- so a failed conversion
// can be diagnosed from the log instead of guessed at. The newest
// thirty session logs are kept; older ones are pruned.
try {
string sLogDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"EdSharp\logs");
Directory.CreateDirectory(sLogDir);
App.LogFile = Path.Combine(sLogDir, "EdSharp_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log");
StringBuilder sbHeader = new StringBuilder();
sbHeader.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("  EdSharp ").Append(BuildVersion.Version).Append(" starting.\r\n");
sbHeader.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("  Program: ").Append(Application.ExecutablePath).Append("\r\n");
sbHeader.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("  Arguments: ").Append(String.Join(" ", cmdLineArgs)).Append("\r\n");
sbHeader.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("  Windows: ").Append(Environment.OSVersion.ToString()).Append(", 64-bit process: ").Append(Environment.Is64BitProcess).Append("\r\n");
sbHeader.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append("  Working directory: ").Append(Environment.CurrentDirectory).Append("\r\n");
File.AppendAllText(App.LogFile, sbHeader.ToString());
string[] aOldLogs = Directory.GetFiles(sLogDir, "EdSharp_2*.log");
Array.Sort(aOldLogs);
for (int iLog = 0; iLog < aOldLogs.Length - 30; iLog++) {
try { File.Delete(aOldLogs[iLog]); } catch (Exception) {}
}
}
catch (Exception) { App.LogFile = ""; }

// Multicore background JIT: record JIT decisions on first launch and, on
// later launches, compile methods in parallel on background cores. This
// shortens startup for a large single-assembly app. Wrapped so a failure
// (e.g. read-only profile folder) never blocks launch.
try {
string sProfileDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EdSharp");
Directory.CreateDirectory(sProfileDir);
System.Runtime.ProfileOptimization.SetProfileRoot(sProfileDir);
System.Runtime.ProfileOptimization.StartProfile("startup.jitprofile");
}
catch {}
// Environment.SetEnvironmentVariable("EdSharpIndent", "", EnvironmentVariableTarget.User);
if (System.IO.File.Exists(App.IndentModeFile)) System.IO.File.Delete(App.IndentModeFile);
Application.EnableVisualStyles();
//Application.SetCompatibleTextRenderingDefault(true);
Application.SetCompatibleTextRenderingDefault(false);
Application.OleRequired();

Shell = new App();
// The session's own bracket. Every log line is appended as it happens --
// nothing waits in memory for a tidy moment that a crash may never grant
// -- and this last line marks a clean finish, so a log without it says
// the program ended some other way, which is worth knowing when reading
// one after a problem.
try {
Shell.Run(cmdLineArgs);
Util.Log("EdSharp closed normally.");
}
catch (Exception ex) {
Util.Log("EdSharp ended with an error: " + ex.Message);
Util.Log(ex.ToString());
throw;
}
} // Main method

public App() {
base.IsSingleInstance = true;
/*
//this.IsSingleInstance = true;
base.IsSingleInstance = true;
App.ProgramName = GetAppName();
App.NetDir = RuntimeEnvironment.GetRuntimeDirectory();
// App.ProgramDir = GetProgramDir();
App.ProgramDir = GetProgramDir();
App.DataDir = GetDataDir();
App.TempFile = GetTempFile();
App.DefaultIniFile = GetDefaultIniFile();
App.HotkeyIniFile = Path.Combine(App.ProgramDir, "Hotkeys.ini");
App.IniFile = GetIniFile();

App.BomDictionary = Util.GetBomDictionary();
SetConfigurationValues();
App.SpeechLog = Path.Combine(App.DataDir, "Speech.log");
if (File.Exists(App.SpeechLog)) File.Delete(App.SpeechLog);
App.ExtraSpeech = (App.ReadOption("E&xtraSpeech", "Y").ToLower().Substring(0, 1) == "n") ? false : true;

InitNetSdk();
InitJFW();
*/

this.Shutdown += delegate(object o, EventArgs e) {
if (App.WordCreated) {
Util.Say("Exiting Microsoft Word");
COM.WordExit();
}
if (App.Boo != null) COM.Release(ref App.Boo);
if (App.JAWS != null) COM.Release(ref App.JAWS);

if (System.IO.File.Exists(App.IndentModeFile)) System.IO.File.Delete(App.IndentModeFile);
foreach (string sFile in App.TempFiles) if (File.Exists(sFile)) File.Delete(sFile);
};

this.UnhandledException += delegate(object sender, Microsoft.VisualBasic.ApplicationServices.UnhandledExceptionEventArgs e) {
Exception ex = (Exception) e.Exception;
string sMessage = ex.Message;
sMessage += "\n\nStack trace:\n" + ex.StackTrace;
// sMessage += "\nExit EdSharp?\n\nStack trace:\n" + ex.StackTrace;
// e.ExitApplication = Dialog.Confirm("Confirm", "Unexpected event!\n" + sMessage + ".\nExit EdSharp?", "N") == "Y";
string[] aButtons = {"&Mail to Developer", "Copy to Clipboard", "Exit EdSharp"};
string sButton = Dialog.Choose("Unexpected Event", sMessage, aButtons, 0);
switch (sButton) {
case "&Mail to Developer" :
Util.Say("Please add steps to reproduce the problem, if possible.");
string sSubject = "EdSharp error: " + ex.Message;
KeyValuePair<string, string>[] aAddresses = new KeyValuePair<string, string>[1];
string sName = "Jamal Mazrui";
string sAddress = "jamal@EmpowermentZone.com";
aAddresses[0] = new KeyValuePair<String, String>(sName, sAddress);
try {
MapiMail.SendMail(sSubject, sMessage, aAddresses, null);
}
catch {
Util.MailMessage(sAddress, sSubject, sMessage);
}
break;
case "Copy to Clipboard" :
Util.SetClipboardText(sMessage);
break;
case "Exit EdSharp" :
// Application.Exit();
e.ExitApplication = true;
return;
}
e.ExitApplication = false;
};

this.Startup += delegate(object sender, Microsoft.VisualBasic.ApplicationServices.StartupEventArgs e) {

//this.IsSingleInstance = true;
// base.IsSingleInstance = true;
App.ProgramName = GetAppName();
App.NetDir = RuntimeEnvironment.GetRuntimeDirectory();
App.ProgramDir = GetProgramDir();
App.DataDir = GetDataDir();
App.TempFile = GetTempFile();
App.DefaultIniFile = GetDefaultIniFile();
App.HotkeyIniFile = Path.Combine(App.ProgramDir, "Hotkeys.ini");
App.IniFile = GetIniFile();
App.IndentModeFile = Path.Combine(App.DataDir, "IndentMode.tmp");
// IniFile = GetIniFile();

App.BomDictionary = Util.GetBomDictionary();
SetConfigurationValues();
App.SpeechLog = Path.Combine(App.DataDir, "Speech.log");
if (File.Exists(App.SpeechLog)) File.Delete(App.SpeechLog);
App.ExtraSpeech = (App.ReadOption("E&xtraSpeech", "Y").ToLower().Substring(0, 1) == "n") ? false : true;
App.IndentChange = App.ReadOption("E&xtraSpeech", "Y").Contains("-") ? false : true;

InitNetSdk();
InitJFW();

Frame = new MdiFrame();
Homer.Say.attach(Frame);
this.MainForm = Frame;
MdiChild child = new MdiChild(Frame);
if (App.ReadOption("OpenPrevious", "Y").ToLower().Substring(0, 1) != "n") {
string[] aFiles = App.ReadSectionKeys("Previous");
int iCount = 0;
foreach (string s in aFiles) {
if (!File.Exists(s)) continue;
iCount ++;
int iIndex = Int32.Parse(App.ReadValue("Previous", s, "-1"));
// App.Frame.OpenOrActivateWindow(s, 1);
App.Frame.OpenOrActivateWindow(s, App.Frame.GetViewLevel(s));
if (App.Frame.Child.RTB.Index == 0) App.Frame.Child.RTB.Index = iIndex;
}
if (iCount > 0) App.Frame.AddMessage("Opened " + iCount + " previous file" + (iCount == 1 ? "" : "s"));
}
App.DeleteSection("Previous");

ReadOnlyCollection<string> cmdLineArgs = this.CommandLineArgs;
if (cmdLineArgs.Count > 0) {
string sFile = cmdLineArgs[0];
string sLine = "";
string sColumn = "";
if (cmdLineArgs.Count > 1) sLine = cmdLineArgs[1];
if (cmdLineArgs.Count > 2) sColumn = cmdLineArgs[2];
// Frame.OpenOrActivateWindow(sFile, 1, sLine, sColumn);
App.Frame.OpenOrActivateWindow(sFile, App.Frame.GetViewLevel(sFile), sLine, sColumn);
}
};

} // App constructor

protected override void OnStartupNextInstance(StartupNextInstanceEventArgs e) {
/*
Util.ActivatePid(Process.GetCurrentProcess().Id);
Microsoft.VisualBasic.Interaction.AppActivate(App.Frame.Text);
App.Frame.Activate();
Util.ActivateTitle(App.Frame.Text);
*/
//COM.ActivateTitle(App.Frame.Text);
//System.Threading.Thread.Sleep(1000);

/*
object oAutoIt = COM.CreateObject("AutoItX3.Control");
object[] aParams = {"WinTitleMatchMode", 4};
COM.CallMethod(oAutoIt, "AutoItSetOption", aParams);
string sParam = "handle=" + App.Frame.TopLevelControl.Handle.ToString();
COM.CallMethod(oAutoIt, "WinActivate", sParam);
Win32.SetForegroundWindow(App.Frame.TopLevelControl.Handle);
*/

//Process.Start(Path.Combine(App.ProgramDir, "ForceWin.exe"), App.Frame.TopLevelControl.Handle.ToString());
Win32.ForceWindow(App.Frame.TopLevelControl.Handle);

if (e.CommandLine.Count == 0) return;
string sFile = Util.Unquote(e.CommandLine[0]);
string sLine = "";
string sColumn = "";
if (e.CommandLine.Count > 1) sLine = e.CommandLine[1];
if (e.CommandLine.Count > 2) sColumn = e.CommandLine[2];
if (sFile != null && File.Exists(sFile)) App.Frame.OpenOrActivateWindow(sFile, App.Frame.GetViewLevel(sFile), sLine, sColumn);
} // OnStartUpNextInstance handler

public static bool InitJFW() {
string sDir = Win32.GetJFWDir();
if (sDir.Length == 0) return false;

string sPath = Environment.GetEnvironmentVariable("PATH");
sDir += ";";
if (!sPath.ToLower().Contains(sDir.ToLower())) {
sPath = sDir + sPath;
Environment.SetEnvironmentVariable("PATH", sPath);
}
return true;
} // InitJFW method

public static bool InitNetSdk() {
// Does not work
// string sDir = Win32.GetNetRuntimeDir();
// string sDir = Win32.GetNetSdkDir();
string sDir = RuntimeEnvironment.GetRuntimeDirectory();
// Dialog.Show("RuntimeEnvironment", RuntimeEnvironment.GetSystemVersion() + "\r\n" + RuntimeEnvironment.GetRuntimeDirectory() + "\r\n" + RuntimeEnvironment.SystemConfigurationFile);
// Dialog.Show(sDir);

if (sDir.Length == 0) return false;
if (sDir.EndsWith(@"\")) sDir = sDir.Substring(0, sDir.Length - 2);

string sPath = Environment.GetEnvironmentVariable("PATH");
sDir += ";";
if (!sPath.ToLower().Contains(sDir.ToLower())) {
sPath = sDir + sPath;
Environment.SetEnvironmentVariable("PATH", sPath);
// Clipboard.SetText(Environment.GetEnvironmentVariable("PATH"));
}
return true;
} // InitNetSdk method

public static string GetAppName() {
string sExe = Environment.GetCommandLineArgs()[0];
string sReturn = Path.GetFileNameWithoutExtension(sExe);
sReturn = Application.ProductName;
return sReturn;
} // GetAppName method

public static string GetProgramDir() {
//string sApp = System.Reflection.Assembly.GetExecutingAssembly().Location;
//string sApp = Application.ExecutablePath;
//string sReturn = Path.GetDirectoryName(sApp);
string sReturn = Application.StartupPath;
return sReturn;
} // GetProgramDir method

public static string GetDataDir() {
string sName = GetAppName();
//string sDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
//string sDir = Application.UserAppDataPath;
//string sDir = Application.LocalUserAppDataPath
string sDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
string sReturn = Path.Combine(sDir, sName);
if (!Directory.Exists(sReturn)) Directory.CreateDirectory(sReturn);
return sReturn;
} // GetDataDir method

public static string GetTempFile() {
string sName = GetAppName() + ".tmp";
string sDir = GetDataDir();
string sReturn = Path.Combine(sDir, sName);
App.TempFiles.Add(sReturn);
return sReturn;
} // GetTempFile method

public static string GetIniFile() {
string sName = GetAppName() + ".ini";
string sDir = GetDataDir();
string sReturn = Path.Combine(sDir, sName);
return sReturn;
} // GetIniFile method

public static string GetDefaultIniFile() {
string sName = GetAppName() + ".ini";
string sDir = GetProgramDir();
string sReturn = Path.Combine(sDir, sName);
return sReturn;
} // GetDefaultIniFile method

public static string ReadData(string sKey, string sDefault) {
string sSection = "Data";
return ReadValue(sSection, sKey, sDefault);
} // ReadData method

public static bool WriteData(string sKey, string sValue) {
string sSection = "Data";
//return WriteValue(sSection, sKey, sValue);
return Ini.WriteQuote(App.IniFile, sSection, sKey, sValue);
} // WriteData method

public static string ReadDefaultOption(string sKey, string sDefault) {
string sSection = "Options";
return Ini.ReadValue(App.DefaultIniFile, sSection, sKey, sDefault);
} // ReadDefaultOption method

public static string ReadOption(string sKey, string sDefault) {
string sSection = "Options";
return ReadValue(sSection, sKey, sDefault);
} // ReadOption method

public static string ReadValue(string sSection, string sKey, string sDefault) {
return Ini.ReadValue(App.IniFile, sSection, sKey, sDefault);
} // ReadValue method

public static bool WriteOption(string sKey, string sValue) {
string sSection = "Options";
return WriteValue(sSection, sKey, sValue);
} // WriteOption method

public static bool WriteValue(string sSection, string sKey, string sValue) {
// return Ini.WriteValue(App.IniFile, sSection, sKey, sValue);
return Ini.WriteQuote(App.IniFile, sSection, sKey, sValue);
} // WriteValue method

public static bool DeleteKey(string sSection, string sKey) {
return Ini.DeleteKey(App.IniFile, sSection, sKey);
} // DeleteKey method

public static bool DeleteSection(string sSection) {
return Ini.DeleteSection(App.IniFile, sSection);
} // DeleteSection method

public static string[] ReadDefaultOptions() {
string sSection = "Options";
return ReadDefaultSectionKeys(sSection);
} // ReadOptions method

public static string[] ReadDefaultSectionKeys(string sSection) {
bool bIncludeComments = false;
return ReadDefaultSectionKeys(sSection, bIncludeComments);
} // ReadDefaultSectionKeys method

public static string[] ReadDefaultSectionKeys(string sSection, bool bIncludeComments) {
return Ini.ReadSectionKeys(App.DefaultIniFile, sSection, bIncludeComments);
} // ReadDefaultSectionKeys method

public static string[] ReadSectionKeys(string sSection) {
return Ini.ReadSectionKeys(App.IniFile, sSection);
} // ReadSectionKeys method

public static string[] ReadSections() {
return Ini.ReadSections(App.IniFile);
} // ReadSections method

public static void SetConfigurationValues() {
string[] aSections = Ini.ReadSections(App.DefaultIniFile);
foreach (string sSection in aSections) {
string[] aCommands = Ini.ReadSectionKeys(App.DefaultIniFile, sSection, false);
string[] aKeys = new string[aCommands.Length];
for (int i = 0; i < aCommands.Length; i++) {
string sCommand = aCommands[i];
string sKey = Ini.ReadValue(App.DefaultIniFile, sSection, sCommand, "");
sKey = Ini.ReadValue(App.IniFile, sSection, sCommand, sKey);
aKeys[i] = sKey;
}

//Ini.DeleteSection(App.IniFile, sSection);
for (int i = 0; i < aCommands.Length; i++) {
string sCommand = aCommands[i];
string sKey = aKeys[i];
//if (sSection == "Import" || sSection == "Export") Ini.WriteValue(App.IniFile, sSection, sCommand, sKey);
//else Ini.WriteQuote(App.IniFile, sSection, sCommand, sKey);
Ini.WriteQuote(App.IniFile, sSection, sCommand, sKey);
}
}
} // SetConfigurationValues method

} // App class

public class MdiChild : Form {
public HomerRichTextBox RTB;
public Encoding YieldEncoding = null;
public bool IsUnixLineBreak = false;
public int AppendFromClipboard = 0;
public IntPtr NextClipboardViewer = (IntPtr) 0;
public int LastTickCount = 0;
public string LastClipboardText = "";
private string sFile = "";
public string File {
get {
return sFile;
}
set {
sFile = value;
}
} // File property

public DateTime FileTime;
public bool FileTimeChecked = false;
public MdiChild(MdiFrame frame) {
string sTitle = frame.GetNoNameTitle();
new MdiChild(frame, sTitle);
} // MdiChild constructor

public MdiChild(MdiFrame frame, string sTitle) {
this.SuspendLayout();
this.MdiParent = frame;
HomerRichTextBox rtb = new HomerRichTextBox();
rtb.GotFocus += CheckFileTime;
rtb.AccessibleRole = AccessibleRole.Text;
rtb.AutoWordSelection = false;
rtb.Dock = DockStyle.Fill;
rtb.Multiline = true;

string sFont = App.ReadOption("FontDefault", "");
if (sFont.Length > 0) {
string[] a = sFont.Split(',');
List<string> list = new List<string>(a);
int iCount = list.Count;
string sColor = list[iCount - 1];
try {
sColor = sColor.Split('=')[1];
rtb.ForeColor = Util.String2Color(sColor);
}
catch {}

list.RemoveAt(iCount - 1);
a = list.ToArray();
sFont = String.Join(",", a);
try {
//sFont = "Arial Unicode MS";
rtb.Font = Util.String2Font(sFont);
}
catch {}
}

string s = App.ReadOption("WordWrap", "Y").Trim().ToUpper();
if (s == "N" || s == "NO") rtb.SetWrap(false);
else rtb.SetWrap(true);
//rtb.ScrollBars = RichTextBoxScrollBars.Vertical;
rtb.ScrollBars = RichTextBoxScrollBars.Vertical | RichTextBoxScrollBars.Horizontal;
rtb.AcceptsTab = true;
rtb.FindText = "";
rtb.JumpLine = "";
rtb.GoPercent = "";
rtb.SearchTopic = "";
rtb.SelectionChanged += App.Frame.SetStatusAddress;
this.Controls.Add(rtb);
this.RTB = rtb;
//this.File = frame.GetNoNameTitle();
this.File = sTitle;
this.Text = System.IO.Path.GetFileName(this.File);
this.StartPosition = FormStartPosition.CenterParent;
this.AutoSize = true;
this.ResumeLayout();
this.KeyPreview = true;
this.Activated += delegate(object o, EventArgs e) {
frame.SetStatusAddress(this, null);
//this.WindowState = FormWindowState.Maximized;
//Win32.SetForegroundWindow(App.Frame.Handle);
//Win32.SetForegroundWindow(this.Handle);
//COM.ActivateTitle("EdSharp");
//int iPid = Process.GetCurrentProcess().Id;
//if (iPid > 0) Util.ActivatePid(iPid);

//Win32.ForceWindow(App.Frame.TopLevelControl.Handle);
};

string sText, sResult = "";
this.Shown += delegate(object o, EventArgs e) {
this.WindowState = FormWindowState.Maximized;
//Win32.ForceWindow(App.Frame.TopLevelControl.Handle);

sFile = this.File;
if (!sFile.Contains(@"\")) return;

sText = App.ReadValue("Favorites", sFile, "");
try {
string[] a = sText.Split('|');
sText = a[0];
rtb.Index = Int32.Parse(sText);
return;
}
catch {}

sText = App.ReadValue("Recent", sFile, "");
if (sText.Length == 0) return;

rtb = this.RTB;
HomerList hl = new HomerList(sText);
hl.KeepLike(@"^\d+$");
// hl.Remove("-1");
if (hl.Count == 0) {
return;
}

sResult = hl[0];
rtb.Index = Int32.Parse(sResult);
App.Frame.AddMessage("Previous percent " + rtb.Percent);
// Util.Say(rtb.RowText);
}; // Shown

this.Closing += delegate(object o, CancelEventArgs e) {
sFile = this.File;
if (!sFile.Contains(@"\")) return;
rtb = this.RTB;
int iIndex = rtb.Index;
if (iIndex == 0) return;

sText = App.ReadValue("Recent", sFile, "");
HomerList hl = new HomerList(sText);
hl.KeepLike(@"\d+");
hl.Remove("-1");
DateTime dt = DateTime.Now;
string sTime = dt.ToString("u");
sTime = sTime.Substring(0, sTime.Length - 1);
sText = sTime + "|" + iIndex + "|" + (rtb.ReadOnly ? "G" : "M") + "|" + (string) Util.If(rtb.WordWrap, "W", "U");
// hl.AddUniqueRange(sText);
// sText = hl.Segments;
App.WriteValue("Recent", sFile, sText);
}; // Closing

this.FileTime = System.IO.File.GetLastWriteTime(this.File);
this.Show();
} // child constructor

public void CheckFileTime(object sender, EventArgs e) {
// Moving between windows leaves Key Describer mode, since the mode
// belongs to the window it was switched on in. The status bar records
// it; it is not SPOKEN here, because a screen reader is already
// announcing the window switch and a second voice on top of that is
// exactly the chatter this handler used to add.
if (App.Frame.KeyDescriber) {
App.Frame.SetStatus("No Key Describer");
App.Frame.KeyDescriber = false;
}

// Jim's request of 19 August 2026, in the spirit of the word wrap
// announcement: a document that is collecting clipboard copies says so
// the moment focus arrives, so the mode is never a surprise.
if (this.AppendFromClipboard == 1) Util.Say("Append from clipboard");

string sFile = this.File;
// The mode-flag file may be held open by a screen reader script at
// this instant; a share collision must never take down a focus event.
try {
bool b = System.IO.File.Exists(App.IndentModeFile);
if (b && !this.RTB.IndentMode) System.IO.File.Delete(App.IndentModeFile);
else if (!b && this.RTB.IndentMode) System.IO.File.Create(App.IndentModeFile).Close();
}
catch (Exception) {}
if (this.FileTimeChecked || sFile.IndexOf(@"\") == -1 || !System.IO.File.Exists(sFile)) return;

DateTime dt = System.IO.File.GetLastWriteTime(sFile);
//if (this.FileTime >= dt || Util.File2String(sFile).Length == 0) return;
if (this.FileTime >= dt) return;
this.FileTimeChecked = true;
switch (Dialog.Confirm("Confirm", this.Text + " on disk is newer than the version opened in this window.  Open Again?", "Y")) {
case "Y" :
int iIndex = this.RTB.Index;
this.LoadTextOrRtfFile(sFile);
this.RTB.Index = iIndex;
break;
case "N":
break;
default :
this.FileTimeChecked = false;
return;
}
} // CheckFileTime handler

protected override void WndProc(ref Message m) {
base.WndProc(ref m);

const int WM_CHANGECBCHAIN = 0x30D;
const int WM_DRAWCLIPBOARD = 0x308;
//if (m.Msg == 776) {
switch (m.Msg) {
case WM_DRAWCLIPBOARD :
// Pass the notification along the clipboard-viewer chain.  Handles are
// pointer-sized: comparing or passing them as int truncates on a 64-bit build,
// and wParam/lParam were also being passed in the wrong order.  Both are fixed
// here.  The whole body is guarded, because an exception thrown inside a window
// procedure is fatal: that is how a clipboard hiccup crashed EdSharp.
try {
if (this.NextClipboardViewer != IntPtr.Zero) Win32.SendMessagePtr(this.NextClipboardViewer, m.Msg, m.WParam, m.LParam);

if (this.AppendFromClipboard == -1) {
this.AppendFromClipboard = 1;
}
else if (this.AppendFromClipboard == 1) {
string sClipboard = Util.GetClipboardText();
// Do not append a copy made in this same collecting window, which would copy the
// document back into itself.  this.ContainsFocus is true only when the input
// focus is in this document, so the copy originated here; a copy made in another
// EdSharp window or in another application is still collected.
if (this.ContainsFocus) sClipboard = "";
//if (sClipboard == this.LastClipboardText && ((Environment.TickCount - this.LastTickCount) < 100)) sClipboard = "";
if (sClipboard == this.LastClipboardText) sClipboard = "";
if (sClipboard.Length > 0) {
this.LastTickCount = Environment.TickCount;
this.LastClipboardText = sClipboard;
Console.Beep();

HomerRichTextBox rtb = this.RTB;
string sText = rtb.Text;
sText = sText.TrimEnd(new char[] {'\n'});
int iLength = sText.Length;
if (iLength >0) sText += "\f\n";
//if (iLength > 0 && sText.Substring(iLength - 1) != "\n") sText += "\n";
sClipboard = sClipboard.TrimEnd(new char[] {'\n'});
sText += sClipboard;
rtb.Text = sText;
rtb.Index = rtb.Text.Length - 1;
} // sClipboard.Length
} // this.AppendFromClipboard
}
catch (Exception) {
// Never let a clipboard problem take the program down.  The clipboard is a
// shared resource: another application can hold it locked, and content copied
// from a browser can fail to convert to text.  Skipping one clip is harmless;
// crashing is not.
}
break;
case WM_CHANGECBCHAIN :
try {
// Same pointer-sized handling here.  This message rebuilds the viewer chain, so
// truncated handles or swapped parameters corrupt the chain for every viewer --
// which is how a crash could surface later, on an unrelated copy.
IntPtr hNextClipboardViewer = m.WParam;
if (this.NextClipboardViewer == hNextClipboardViewer) this.NextClipboardViewer = m.LParam;
else if (this.NextClipboardViewer != IntPtr.Zero) Win32.SendMessagePtr(this.NextClipboardViewer, m.Msg, m.WParam, m.LParam);
}
catch (Exception) { }
break;
} // switch msg
} // WndProc event handler

public Encoding GetYieldEncoding() {
// The YieldEncoding option is the explicit encoding the user wants for both
// reading and writing a document. Blank returns null, which lets the open
// path auto-detect and the save path default to UTF-8 with BOM (utf8b).
// Friendly names are recognized in addition to a numeric code page or any
// .NET encoding name.
Encoding en = null;
string sEncoding = App.ReadOption("YieldEncoding", "").Trim();
string sKey = sEncoding.Replace("-", "").Replace("_", "").ToLower();
if (sKey == "utf8n") {
en = new UTF8Encoding(false);
this.IsUnixLineBreak = true;
}
else if (sKey == "utf8b" || sKey == "utf8") en = new UTF8Encoding(true);
else if (sKey == "utf16" || sKey == "utf16le" || sKey == "unicode") en = Encoding.Unicode;
else if (sKey == "utf16be") en = Encoding.BigEndianUnicode;
else if (sKey == "ansi" || sKey == "default") en = Encoding.Default;
else if (sEncoding.Length > 0 ) {
try {
if (Util.IsNumeric(sEncoding)) en = Encoding.GetEncoding(Int32.Parse(sEncoding));
else en = Encoding.GetEncoding(sEncoding);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
en = new UTF8Encoding(true);
}
}
return en;
} // GetYieldEncoding method

public void LoadTextOrRtfFile(string sFile) {
bool bLiteral = false;
LoadTextOrRtfFile(sFile, bLiteral);
} // LoadTextOrRtfFile method

public void LoadTextOrRtfFile(string sFile, bool bLiteral) {
this.FileTimeChecked = false;
this.FileTime = System.IO.File.GetLastWriteTime(sFile);
// Jim's request of 24 August 2026: the current directory used to follow
// only SAVES, so opening a file to read it left the next Open dialog
// pointing at wherever the last save happened, and reaching a sibling
// file meant navigating the whole way again. Every load now brings the
// current directory along, so opening a file makes its folder the
// starting point for the next open.
try {
if (sFile != null && sFile.IndexOf(@":\") > 0) Directory.SetCurrentDirectory(Path.GetDirectoryName(sFile));
}
catch (Exception) {}
try {
if (!bLiteral && Path.GetExtension(sFile).ToLower() == ".rtf") this.RTB.LoadFile(sFile, RichTextBoxStreamType.RichText);
//else this.RTB.LoadFile(sFile, RichTextBoxStreamType.UnicodePlainText);
else {
Encoding en = GetYieldEncoding();
string sText = Util.File2String(sFile, ref en);
this.RTB.Text = sText;
// Dialog.Show(sText.Length, this.RTB.TextLength);
if (sText.Length > 1 && this.RTB.TextLength == 1) {
en = Encoding.Unicode;
this.RTB.Text = Util.File2String(sFile, ref en);
}
this.YieldEncoding = en;
}
//else this.RTB.Text = Util.OldFile2String(sFile);
//else this.RTB.Text = System.IO.File.ReadAllText(sFile, System.Text.Encoding.UTF8);
//else this.RTB.Text = System.IO.File.ReadAllText(sFile, System.Text.Encoding.Default);
//else this.RTB.Text = System.IO.File.ReadAllText(sFile, System.Text.Encoding.GetEncoding(1252));
this.RTB.Modified = false;
this.Text = Path.GetFileName(sFile);
this.File = sFile;
}
catch {
App.Frame.AddMessage("Cannot open file!  Opening temporary copy.");
if (System.IO.File.Exists(App.TempFile)) System.IO.File.Delete(App.TempFile);
System.IO.File.Copy(sFile, App.TempFile);
App.Frame.OpenOrActivateWindow(App.TempFile);
}
//Dialog.Show(this.File);
// Stop double bookmark at message
// App.Frame.ApplyFileOptions(sFile);
} // LoadTextFile method

public void SaveTextOrRtfFile(string sFile) {
if (System.IO.File.Exists(sFile)) {
string sKeepBackup = App.ReadOption("KeepBackup", "N").Trim().ToLower();
if (sKeepBackup == "y" || sKeepBackup == "yes") {
string sBak = sFile + ".bak";
if (System.IO.File.Exists(sBak)) System.IO.File.Delete(sBak);
System.IO.File.Copy(sFile, sBak);
}
}

if (Path.GetExtension(sFile).ToLower() == ".rtf") this.RTB.SaveFile(sFile, RichTextBoxStreamType.RichText);
//else if (Util.IsUnicode(this.RTB.Text)) Util.String2File(this.RTB.Text, sFile);
//else this.RTB.SaveFile(sFile, RichTextBoxStreamType.PlainText);
else {
Encoding en = this.YieldEncoding;
if (en == null) en = GetYieldEncoding();
if (en == null) en = new UTF8Encoding(true);
string sText = this.RTB.Text;
if (!this.IsUnixLineBreak) sText = Util.Convert2WinLineBreak(sText);
Util.String2File(sText, sFile, ref en);
}
this.RTB.Modified = false;
App.Frame.SetRecent(sFile);
this.Text = Path.GetFileName(sFile);
this.File = sFile;
this.FileTime = System.IO.File.GetLastWriteTime(sFile);
this.FileTimeChecked = false;
} // SaveTextOrRtfFile method

protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
return App.Frame.ProcessCmdKey_Helper(ref msg, keyData);
} // ProcessCmdKey handler

} // MdiChild class

public class MdiFrame : Form {
public string LastDescription = "";
public bool KeyDescriber = false;
public string KeyString = "";
public int KeyRepeat = 0;
public int KeyIndex = -1;

public MdiChild Child {
get {
MdiChild child = this.ActiveMdiChild as MdiChild;
if (child == null) {
Form[] children = this.MdiChildren;
int iLength = children.Length;
if (children.Length > 0) child = (MdiChild) children[iLength - 1];
}
return child;
}
set {
value.Activate();
}
} // Child property

public bool FindWithRegExp = false;
public bool bCommandComplete = true;
public static string CR = "\r";
public static string LF = "\n";
public static string LB = LF;
public static string LineBreak = Environment.NewLine;
public static string FF = "\f";
public static string SB = FF + LB;
public static string DD = "----------";
public static string SectionBreak = LB + DD + LB + SB;
public static string EOD = LB + DD + LB + "End of Document" + LB;

public static Dictionary<Keys, ToolStripMenuItem> hashKey = new Dictionary<Keys, ToolStripMenuItem>();
public MenuStrip menuMain;
public ToolStripMenuItem menuFile, menuFileNew, menuFileNewFromClipboard, menuFileOpen, menuFileOpenOtherFormat, menuFileOpenAgain, menuFileRecent, menuFileSetFavorite, menuFileClearFavorite, menuFileListFavorites, menuFileFind, menuFileSave, menuFileSaveAs, menuFileSaveCopy, menuFileExport, menuFileRename, menuFileProperties, menuFileMailBody, menuFileMailAttach, menuFilePrint, menuFileRun, menuFileCurrentWindows, menuFileClose, menuFileCloseAllButCurrentWindow, menuFileExit;
public ToolStripMenuItem menuEdit, menuEditSelectAll, menuEditUnselectAll, menuEditCopy, menuEditCopyAppend, menuEditCopyRichText, menuEditCut, menuEditCutAppend, menuEditPaste, menuEditPasteFile, menuEditUndo, menuEditRedo, menuEditStartSelection, menuEditCompleteSelection, menuEditReselect, menuEditCopyAll, menuEditSelectChunk, menuEditAppendFromClipboard, menuEditQuote, menuEditUnquote, menuEditUpperCase, menuEditLowerCase, menuEditProperCase, menuEditSwapCase, menuEditYieldEncoding, menuEditJoinLines, menuEditHardLineBreak, menuEditEnterNewLine, menuEditIndentNewLine, menuEditIndentNewLinePrior, menuEditIndent, menuEditOutdent, menuEditAlign, menuEditIndentMode, menuEditJustify, menuEditStyle, menuEditBaseline, menuEditSetSelectionFont;
public ToolStripMenuItem menuDelete, menuDeleteReplaceRegular, menuDeleteReplaceWithRegExp, menuDeleteHardLine, menuDeleteParagraph, menuDeleteLine, menuDeleteRight, menuDeleteLeft, menuDeleteDown, menuDeleteUp, menuDeleteFile, menuDeleteTrimBlanks;
public ToolStripMenuItem menuNavigate, menuNavigateForwardFind, menuNavigateReverseFind, menuNavigateForwardFindWithRegExp, menuNavigateReverseFindWithRegExp,  menuNavigateForwardFindAtCursor, menuNavigateReverseFindAtCursor, menuNavigateForwardFindAgain, menuNavigateReverseFindAgain, menuNavigateJumpToLine, menuNavigateJumpToLineAgain, menuNavigateGoToPercent, menuNavigateGoToPercentAgain, menuNavigateGoToPart, menuNavigateSetBookmark, menuNavigateClearBookmark, menuNavigateGoToBookmark, menuNavigateHomeCharacter, menuNavigateEndCharacter, menuNavigateStartTag, menuNavigateEndTag,  menuNavigateNextJustify, menuNavigatePriorJustify, menuNavigateNextStyle, menuNavigatePriorStyle, menuNavigateNextBaseline, menuNavigatePriorBaseline, menuNavigateNextFont, menuNavigatePriorFont, menuNavigateRightBrace, menuNavigateNextBlock, menuNavigatePriorBlock, menuNavigateLeftBrace, menuNavigateNextIndent, menuNavigatePriorIndent, menuNavigateNextChunk,  menuNavigatePriorChunk, menuNavigateNextSentence, menuNavigatePriorSentence, menuNavigateNextParagraph, menuNavigatePriorParagraph, menuNavigateNextPart, menuNavigatePriorPart, menuNavigateNextSection, menuNavigatePriorSection, menuNavigateGoToSection, menuNavigateGoToContents, menuNavigateSearchForTopic, menuNavigateSearchForTopicAgain, menuNavigateGoToStartOfSelection;
public ToolStripMenuItem menuQuery, menuQueryAddress, menuQueryBraces, menuQueryBlock, menuQueryIndent, menuQueryPath, menuQueryTopic, menuQueryYield, menuQueryStatus, menuQueryCompiler, menuQuerySelected, menuQueryChunk, menuQueryReadAll, menuQueryWindowsOpen, menuQueryClipboard, menuQueryTime, menuQueryStyles, menuQueryFont;
public ToolStripMenuItem menuMisc, menuMiscSetDefaultFont, menuMiscConfigurationOptions, menuMiscManualOptions, menuMiscResetConfiguration, menuMiscGoToFolder, menuMiscGoToSpecialFolder, menuMiscWordWrap, menuMiscUnwrap, menuMiscExtraSpeechToggle, menuMiscExtraSpeechLog, menuMiscEnvironmentVariables, menuMiscSpellCheck, menuMiscThesaurus, menuMiscLookupTerm, menuMiscTranslateLanguage, menuMiscGuardDocument, menuMiscNoGuard, menuMiscPyBrace, menuMiscPyDent, menuMiscInferIndent, menuMiscFormatCode, menuMiscRepeatLine, menuMiscSectionBreak, menuMiscPathToClipboard, menuMiscPathList, menuMiscInsertTime, menuMiscCalculateDate, menuMiscHTMLFormat, menuMiscMarkdownToText, menuMiscHtmlToMarkdown, menuMiscHtmlToText, menuMiscPreviewMarkdown, menuMiscPreviewMarkdownBrowser, menuMiscCheckMarkdown, menuMiscRunCodeBlocks, menuMiscChatWithAI, menuMiscChatWithDocument, menuMiscTextConvert, menuMiscTextCombine, menuMiscTextContents, menuMiscYieldWithRegExp, menuMiscExtractWithRegExp, menuMiscRunAtCursor, menuMiscSpecialCharacter, menuMiscEvaluateExpression, menuMiscReplaceTokens, menuMiscTransformFiles, menuMiscGoToEnvironment, menuMiscCompile, menuMiscPickCompiler, menuMiscPromptCommand, menuMiscReviewOutput, menuMiscSaveSnippet, menuMiscInvokeSnippet, menuMiscViewSnippet, menuMiscKeepUniqueItems, menuMiscNumberItems, menuMiscOrderItems, menuMiscReverseItems, menuMiscListDifferentItems, menuMiscQueryCommonItems, menuMiscExplorerFolder, menuMiscCommandPrompt, menuMiscBurnToCD, menuMiscWebDownload;
public ToolStripMenuItem menuWindow, menuWindowNext, menuWindowPrior, menuWindowArrangeIcons, menuWindowCascade, menuWindowTileHorizontal, menuWindowTileVertical;
public ToolStripMenuItem menuHelp, menuHelpAbout, menuHelpDocumentation, menuHelpTutorial, menuHelpHistoryOfChanges, menuHelpSamplePrograms, menuHelpCopyLog, menuHelpKeyDescriber, menuHelpHotKeySummary, menuHelpAlternateMenu, menuHelpContextMenu, menuHelpSendToMenu, menuHelpElevateVersion;
public StatusStrip statusBar;
public ToolStripStatusLabel lblStatus;

public MdiFrame() {
SectionBreak = Util.Literalize(App.ReadOption("SectionBreak", SectionBreak));
this.SuspendLayout();
this.IsMdiContainer = true;
menuMain = CreateMainMenu();
//menuMain.ShowItemToolTips = true;
menuMain.AccessibleRole = AccessibleRole.MenuBar;
menuFile = CreateMenu("&File");
menuFileNew = CreateMenuItem("&New", "Control+N", menuItem_Click, "frame speak");
menuFileNewFromClipboard = CreateMenuItem("New from Clipboard", "Control+Shift+N", menuItem_Click, "frame speak");
menuFileOpen = CreateMenuItem("&Open ...", "Control+O", menuItem_Click, "frame speak");
menuFileOpenOtherFormat = CreateMenuItem("Open Other Format ...", "Control+Shift+O", menuItem_Click, "frame speak");
menuFileOpenAgain = CreateMenuItem("Open Again", "Alt+O", menuItem_Click, "child speak");
menuFileRecent = CreateMenuItem("Recent Files ...", "Alt+R", menuItem_Click, "frame silent");
menuFileSetFavorite = CreateMenuItem("Set Favorite", "Control+&L", menuItem_Click, "child speak");
//menuFileSetFavorite = CreateMenuItem("Set on Favorite &List", "Control+L", menuItem_Click, "child speak");
menuFileClearFavorite = CreateMenuItem("Clear Favorite", "Control+Shift+L", menuItem_Click, "child speak");
menuFileListFavorites = CreateMenuItem("List Favorites ...", "Alt+L", menuItem_Click, "frame silent");
menuFileFind = CreateMenuItem("File Find ...", "Alt+Shift+F", menuItem_Click, "frame speak");
menuFileSave = CreateMenuItem("&Save", "Control+S", menuItem_Click, "child speak");
menuFileSaveAs = CreateMenuItem("Save &As ...", "Control+Shift+S", menuItem_Click, "child silent");
menuFileSaveCopy = CreateMenuItem("Save Copy ...", "Alt+Shift+S", menuItem_Click, "child speak");
menuFileExport = CreateMenuItem("Export Format ...", "Alt+Shift+E", menuItem_Click, "child silent");
menuFileRename = CreateMenuItem("Rename ...", "Alt+Shift+R", menuItem_Click, "child silent");
menuFileProperties = CreateMenuItem("Properties", "Alt+Enter", menuItem_Click, "child speak");
menuFileMailBody = CreateMenuItem("&Mail Body ...", "Control+M", menuItem_Click, "child speak");
menuFileMailAttach = CreateMenuItem("Mail Attachment ...", "Control+Shift+M", menuItem_Click, "child speak");
menuFilePrint = CreateMenuItem("&Print", "Control+P", menuItem_Click, "child silent");
menuFileRun = CreateMenuItem("Run", "F5", menuItem_Click, "child speak");
menuFileCurrentWindows = CreateMenuItem("Current Windows ...", "F4", menuItem_Click, "frame silent");
menuFileClose = CreateMenuItem("&Close Window", "Control+F4", menuItem_Click, "child speak");
menuFileCloseAllButCurrentWindow = CreateMenuItem("Close All but Current Window", "Control+Shift+F4", menuItem_Click, "child speak");
menuFileExit = CreateMenuItem("&E&xit EdSharp", "Alt+F4", menuItem_Click, "frame speak");
menuFile.DropDownItems.AddRange(new ToolStripItem[] {menuFileNew, menuFileNewFromClipboard, menuFileOpen, menuFileOpenOtherFormat, menuFileOpenAgain, menuFileRecent, menuFileSetFavorite, menuFileClearFavorite, menuFileListFavorites, menuFileFind, menuFileSave, menuFileSaveAs, menuFileSaveCopy, menuFileExport, menuFileRename, menuFileProperties, menuFileMailBody, menuFileMailAttach, menuFilePrint, menuFileRun, menuFileCurrentWindows, menuFileClose, menuFileCloseAllButCurrentWindow, menuFileExit});
//Dialog.Show("File.", menuFile.DropDownItems.Count);

menuEdit = CreateMenu("&Edit");
menuEditSelectAll = CreateMenuItem("Select &All", "Control+A", menuItem_Click, "child speak");
menuEditUnselectAll = CreateMenuItem("Unselect All", "Control+Shift+A", menuItem_Click, "child speak");
menuEditCopy = CreateMenuItem("&Copy", "Control+C", menuItem_Click, "child speak");
menuEditCopyAppend = CreateMenuItem("Copy Append", "Alt+C", menuItem_Click, "child speak");
menuEditCopyRichText = CreateMenuItem("Copy Rich Text", "Control+Shift+C", menuItem_Click, "child speak");
menuEditCut = CreateMenuItem("Cut", "Control+&X", menuItem_Click, "child speak");
menuEditCutAppend = CreateMenuItem("Cut Append", "Alt+X", menuItem_Click, "child speak");
menuEditPaste = CreateMenuItem("Paste", "Control+&V", menuItem_Click, "child speak");
menuEditPasteFile = CreateMenuItem("Paste File ...", "Control+Shift+V", menuItem_Click, "child speak");
menuEditUndo = CreateMenuItem("Undo", "Control+&Z", menuItem_Click, "child speak");
menuEditRedo = CreateMenuItem("Redo", "Control+Shift+Z", menuItem_Click, "child speak");
menuEditStartSelection = CreateMenuItem("Start Selection", "F8", menuItem_Click, "child speak");
menuEditCompleteSelection = CreateMenuItem("Complete Selection", "Shift+F8", menuItem_Click, "child speak");
menuEditReselect = CreateMenuItem("Reselect", "Control+Shift+F8", menuItem_Click, "child speak");
menuEditCopyAll = CreateMenuItem("Copy All", "Control+F8", menuItem_Click, "child speak");
menuEditSelectChunk = CreateMenuItem("Select Chunk", "Control+Space", menuItem_Click, "child silent");
menuEditAppendFromClipboard = CreateMenuItem("Append from Clipboard", "Alt+D7", menuItem_Click, "child silent");
menuEditQuote = CreateMenuItem("&Quote", "Control+Q", menuItem_Click, "child speak");
menuEditUnquote = CreateMenuItem("Unquote", "Control+Shift+Q", menuItem_Click, "child speak");
menuEditUpperCase = CreateMenuItem("&Upper Case", "Control+U", menuItem_Click, "child speak");
menuEditLowerCase = CreateMenuItem("Lower Case", "Control+Shift+U", menuItem_Click, "child speak");
menuEditProperCase = CreateMenuItem("Proper Case", "Alt+U", menuItem_Click, "child speak");
menuEditSwapCase = CreateMenuItem("Swap Case", "Alt+Shift+U", menuItem_Click, "child speak");
menuEditYieldEncoding = CreateMenuItem("Yield Encoding", "Alt+Shift+Y", menuItem_Click, "child silent");
menuEditJoinLines = CreateMenuItem("Join Lines", "Control+Shift+J", menuItem_Click, "child speak");
menuEditHardLineBreak = CreateMenuItem("Hard Line Break ...", "Control+Shift+H", menuItem_Click, "child silent");
menuEditEnterNewLine = CreateMenuItem("Enter New Line", "Enter", menuItem_Click, "child silent");
menuEditIndentNewLine = CreateMenuItem("Indent New Line", "Shift+Enter", menuItem_Click, "child silent");
menuEditIndentNewLinePrior = CreateMenuItem("Indent New Line Prior", "Alt+Shift+Enter", menuItem_Click, "child speak");
menuEditIndent = CreateMenuItem("Indent", "Tab", menuItem_Click, "child silent");
menuEditOutdent = CreateMenuItem("Outdent", "Shift+Tab", menuItem_Click, "child silent");
menuEditAlign = CreateMenuItem("Align", "Alt+Shift+A", menuItem_Click, "child speak");
menuEditIndentMode = CreateMenuItem("Indent Mode", "Alt+Shift+I", menuItem_Click, "child speak");
menuEditJustify = CreateMenuItem("Justify ...", "Alt+Shift+J", menuItem_Click, "child silent");
menuEditStyle = CreateMenuItem("Style ...", "Alt+Shift+OemQuestion", menuItem_Click, "child silent");
menuEditBaseline = CreateMenuItem("Baseline ...", "Alt+Shift+D6", menuItem_Click, "child silent");
menuEditSetSelectionFont = CreateMenuItem("Set Selection Font ...", "Alt+Shift+OemMinus", menuItem_Click, "child speak");
menuEdit.DropDownItems.AddRange(new ToolStripItem[] {menuEditSelectAll, menuEditUnselectAll, menuEditCopy, menuEditCopyAppend, menuEditCopyRichText, menuEditCut, menuEditCutAppend, menuEditPaste, menuEditPasteFile, menuEditUndo, menuEditRedo, menuEditStartSelection, menuEditCompleteSelection, menuEditReselect, menuEditCopyAll, menuEditSelectChunk, menuEditAppendFromClipboard, menuEditQuote, menuEditUnquote, menuEditUpperCase, menuEditLowerCase, menuEditProperCase, menuEditSwapCase, menuEditYieldEncoding, menuEditJoinLines, menuEditHardLineBreak, menuEditEnterNewLine, menuEditIndentNewLine, menuEditIndentNewLinePrior, menuEditIndent, menuEditOutdent, menuEditAlign, menuEditIndentMode, menuEditJustify, menuEditStyle, menuEditBaseline, menuEditSetSelectionFont});
//Dialog.Show("Edit.", menuEdit.DropDownItems.Count);

menuDelete = CreateMenu("&Delete");
menuDeleteReplaceRegular = CreateMenuItem("&Replace ...", "Control+R", menuItem_Click, "child silent");
menuDeleteReplaceWithRegExp = CreateMenuItem("Replace with Regular Expression ...", "Control+Shift+R", menuItem_Click, "child silent");
menuDeleteHardLine = CreateMenuItem("Delete Hard Line", "Control+D", menuItem_Click, "child silent");
menuDeleteParagraph = CreateMenuItem("Delete Paragraph", "Control+Shift+D", menuItem_Click, "child silent");
menuDeleteLine = CreateMenuItem("Delete Line", "Alt+Back", menuItem_Click, "child silent");
menuDeleteRight = CreateMenuItem("Delete Right", "Control+Shift+Delete", menuItem_Click, "child silent");
menuDeleteLeft = CreateMenuItem("Delete Left", "Control+Shift+Back", menuItem_Click, "child silent");
menuDeleteDown = CreateMenuItem("Delete Down", "Alt+Shift+Delete", menuItem_Click, "child speak");
menuDeleteUp = CreateMenuItem("Delete Up", "Alt+Shift+Back", menuItem_Click, "child speak");
menuDeleteFile = CreateMenuItem("Delete File", "Alt+Shift+D", menuItem_Click, "child speak");
menuDeleteTrimBlanks = CreateMenuItem("Trim Blanks", "Control+Shift+Enter", menuItem_Click, "child speak");
menuDelete.DropDownItems.AddRange(new ToolStripMenuItem[] {menuDeleteReplaceRegular, menuDeleteReplaceWithRegExp, menuDeleteHardLine, menuDeleteParagraph, menuDeleteLine, menuDeleteRight, menuDeleteLeft, menuDeleteDown, menuDeleteUp, menuDeleteFile, menuDeleteTrimBlanks});
//Dialog.Show("Delete.", menuDelete.DropDownItems.Count);

menuNavigate = CreateMenu("&Navigate");
menuNavigateForwardFind = CreateMenuItem("Forward &Find ...", "Control+F", menuItem_Click, "child silent");
menuNavigateReverseFind = CreateMenuItem("Reverse Find ...", "Control+Shift+F", menuItem_Click, "child silent");
menuNavigateForwardFindWithRegExp = CreateMenuItem("Forward Find with Regular Expression ...", "Control+F3", menuItem_Click, "child silent");
menuNavigateReverseFindWithRegExp = CreateMenuItem("Reverse Find with Regular Expression ...", "Control+Shift+F3", menuItem_Click, "child silent");
menuNavigateForwardFindAtCursor = CreateMenuItem("Forward Find at Cursor", "Alt+F3", menuItem_Click, "child silent");
menuNavigateReverseFindAtCursor = CreateMenuItem("Reverse Find at Cursor", "Alt+Shift+F3", menuItem_Click, "child silent");
menuNavigateForwardFindAgain = CreateMenuItem("Forward Find Again", "F3", menuItem_Click, "child silent");
menuNavigateReverseFindAgain = CreateMenuItem("Reverse Find Again", "Shift+F3", menuItem_Click, "child silent");
menuNavigateJumpToLine = CreateMenuItem("&Jump to Line ...", "Control+J", menuItem_Click, "child silent");
menuNavigateJumpToLineAgain = CreateMenuItem("Jump to Line Again", "Alt+J", menuItem_Click, "child silent");
menuNavigateGoToPercent = CreateMenuItem("&Go to Percent ...", "Control+G", menuItem_Click, "child silent");
menuNavigateGoToPercentAgain = CreateMenuItem("Go to Percent Again", "Alt+G", menuItem_Click, "child silent");
menuNavigateGoToPart = CreateMenuItem("Go to Part", "Alt+Shift+G", menuItem_Click, "child silent");
menuNavigateSetBookmark = CreateMenuItem("Set Bookmar&k", "Control+K", menuItem_Click, "child speak");
menuNavigateClearBookmark = CreateMenuItem("Clear Bookmark", "Control+Shift+K", menuItem_Click, "child speak");
menuNavigateGoToBookmark = CreateMenuItem("Go to Bookmark", "Alt+K", menuItem_Click, "child speak");
menuNavigateHomeCharacter = CreateMenuItem("Home Character", "Alt+Home", menuItem_Click, "child silent");
menuNavigateEndCharacter = CreateMenuItem("End Character", "Alt+End", menuItem_Click, "child silent");
menuNavigateStartTag = CreateMenuItem("Start Tag", "Control+Shift+Oemcomma", menuItem_Click, "child silent");
menuNavigateEndTag = CreateMenuItem("End Tag", "Control+Shift+OemPeriod", menuItem_Click, "child silent");
menuNavigateNextJustify = CreateMenuItem("Next Alignment", "Control+OemCloseBrackets", menuItem_Click, "child silent");
menuNavigatePriorJustify = CreateMenuItem("Prior Alignment", "Control+OemOpenBrackets", menuItem_Click, "child silent");
menuNavigateNextStyle = CreateMenuItem("Next Style", "Control+OemQuestion", menuItem_Click, "child silent");
menuNavigatePriorStyle = CreateMenuItem("Prior Style", "Control+Shift+OemQuestion", menuItem_Click, "child silent");
menuNavigateNextBaseline = CreateMenuItem("Next Baseline", "Control+D6", menuItem_Click, "child silent");
menuNavigatePriorBaseline = CreateMenuItem("Prior Baseline", "Control+Shift+D6", menuItem_Click, "child silent");
menuNavigateNextFont = CreateMenuItem("Next Font", "Control+OemMinus", menuItem_Click, "child silent");
menuNavigatePriorFont = CreateMenuItem("Prior Font", "Control+Shift+OemMinus", menuItem_Click, "child silent");
menuNavigateRightBrace = CreateMenuItem("Right Brace", "Control+Shift+OemCloseBrackets", menuItem_Click, "child silent");
menuNavigateLeftBrace = CreateMenuItem("Left Brace", "Control+Shift+OemOpenBrackets", menuItem_Click, "child silent");
menuNavigateNextBlock = CreateMenuItem("Next Block", "Control+B", menuItem_Click, "child silent");
menuNavigatePriorBlock = CreateMenuItem("Prior Block", "Control+Shift+B", menuItem_Click, "child silent");
menuNavigateNextIndent = CreateMenuItem("Next Indent", "Control+I", menuItem_Click, "child silent");
menuNavigatePriorIndent = CreateMenuItem("Prior Indent", "Control+Shift+I", menuItem_Click, "child silent");
menuNavigateNextChunk = CreateMenuItem("Next Chunk", "Alt+Right", menuItem_Click, "child silent");
menuNavigatePriorChunk = CreateMenuItem("Prior Chunk", "Alt+Left", menuItem_Click, "child silent");
menuNavigateNextSentence = CreateMenuItem("Next Sentence", "Alt+Down", menuItem_Click, "child silent");
menuNavigatePriorSentence = CreateMenuItem("Prior Sentence", "Alt+Up", menuItem_Click, "child silent");
menuNavigateNextParagraph = CreateMenuItem("Next Paragraph", "Control+Down", menuItem_Click, "child silent");
menuNavigatePriorParagraph = CreateMenuItem("Prior Paragraph", "Control+Up", menuItem_Click, "child silent");
menuNavigateNextPart= CreateMenuItem("Next Part", "Alt+PageDown", menuItem_Click, "child silent");
menuNavigatePriorPart= CreateMenuItem("Prior Part", "Alt+PageUp", menuItem_Click, "child silent");
menuNavigateNextSection= CreateMenuItem("Next Section", "Control+PageDown", menuItem_Click, "child silent");
menuNavigatePriorSection= CreateMenuItem("Prior Section", "Control+PageUp", menuItem_Click, "child silent");
menuNavigateGoToSection= CreateMenuItem("Go to Section", "F6", menuItem_Click, "child speak");
menuNavigateGoToContents = CreateMenuItem("Go to Contents", "Shift+F6", menuItem_Click, "child speak");
menuNavigateSearchForTopic = CreateMenuItem("Search for Topic ...", "Control+F6", menuItem_Click, "child silent");
menuNavigateSearchForTopicAgain = CreateMenuItem("Search for Topic Again", "Alt+F6", menuItem_Click, "child silent");
menuNavigateGoToStartOfSelection = CreateMenuItem("Go to Start of Selection", "Alt+Shift+F8", menuItem_Click, "child speak");
menuNavigate.DropDownItems.AddRange(new ToolStripItem[] {menuNavigateForwardFind, menuNavigateReverseFind, menuNavigateForwardFindWithRegExp, menuNavigateReverseFindWithRegExp,  menuNavigateForwardFindAtCursor, menuNavigateReverseFindAtCursor, menuNavigateForwardFindAgain, menuNavigateReverseFindAgain, menuNavigateJumpToLine, menuNavigateJumpToLineAgain, menuNavigateGoToPercent, menuNavigateGoToPercentAgain, menuNavigateGoToPart, menuNavigateSetBookmark, menuNavigateClearBookmark, menuNavigateGoToBookmark, menuNavigateHomeCharacter, menuNavigateEndCharacter, menuNavigateStartTag, menuNavigateEndTag,  menuNavigateNextJustify, menuNavigatePriorJustify, menuNavigateNextStyle, menuNavigatePriorStyle, menuNavigateNextBaseline, menuNavigatePriorBaseline, menuNavigateNextFont, menuNavigatePriorFont, menuNavigateRightBrace, menuNavigateNextBlock, menuNavigatePriorBlock, menuNavigateLeftBrace, menuNavigateNextIndent, menuNavigatePriorIndent, menuNavigateNextChunk,  menuNavigatePriorChunk, menuNavigateNextSentence, menuNavigatePriorSentence, menuNavigateNextParagraph, menuNavigatePriorParagraph, menuNavigateNextPart, menuNavigatePriorPart, menuNavigateNextSection, menuNavigatePriorSection, menuNavigateGoToSection, menuNavigateGoToContents, menuNavigateSearchForTopic, menuNavigateSearchForTopicAgain, menuNavigateGoToStartOfSelection});
//Dialog.Show("Navigate.", menuNavigate.DropDownItems.Count);

menuQuery = CreateMenu("&Query");
menuQueryAddress = CreateMenuItem("Address", "Alt+A", menuItem_Click, "child silent");
menuQueryBraces = CreateMenuItem("Braces", "Alt+Shift+OemCloseBrackets", menuItem_Click, "child silent");
menuQueryBlock = CreateMenuItem("Block", "Alt+B", menuItem_Click, "child silent");
menuQueryIndent = CreateMenuItem("Indentation", "Alt+I", menuItem_Click, "child silent");
menuQueryPath = CreateMenuItem("Path", "Alt+P", menuItem_Click, "child silent");
menuQueryTopic = CreateMenuItem("Topic", "Alt+T", menuItem_Click, "child speak");
menuQueryYield = CreateMenuItem("Yield", "Alt+Y", menuItem_Click, "child speak");
menuQueryStatus = CreateMenuItem("Status", "Alt+Z", menuItem_Click, "child silent");
menuQueryCompiler = CreateMenuItem("Compiler", "Alt+D0", menuItem_Click, "frame silent");
menuQuerySelected = CreateMenuItem("Selected", "Shift+Space", menuItem_Click, "child silent");
menuQueryChunk = CreateMenuItem("Chunk", "Shift+Back", menuItem_Click, "child silent");
menuQueryReadAll = CreateMenuItem("Read All", "Alt+F8", menuItem_Click, "child speak");
menuQueryWindowsOpen = CreateMenuItem("Windows Open", "Shift+F4", menuItem_Click, "child speak");
menuQueryClipboard = CreateMenuItem("Clipboard", "Alt+OemQuotes", menuItem_Click, "frame silent");
menuQueryTime = CreateMenuItem("Time", "Alt+OemSemicolon", menuItem_Click, "frame silent");
menuQueryStyles = CreateMenuItem("Styles", "Alt+OemQuestion", menuItem_Click, "child silent");
menuQueryFont = CreateMenuItem("Font", "Alt+OemMinus", menuItem_Click, "child silent");
menuQuery.DropDownItems.AddRange(new ToolStripItem[] {menuQueryAddress, menuQueryBraces, menuQueryBlock, menuQueryIndent, menuQueryPath, menuQueryTopic, menuQueryYield, menuQueryStatus, menuQueryCompiler, menuQuerySelected, menuQueryChunk, menuQueryReadAll, menuQueryWindowsOpen, menuQueryClipboard, menuQueryTime, menuQueryStyles, menuQueryFont});
//Dialog.Show("Query.", menuQuery.DropDownItems.Count);

menuMisc = CreateMenu("&Misc");
menuMiscSetDefaultFont = CreateMenuItem("Set Default Font and Color ...", "Alt+Shift+Oemplus", menuItem_Click, "child speak");
menuMiscConfigurationOptions = CreateMenuItem("Configuration Options ...", "Alt+Shift+C", menuItem_Click, "frame silent");
menuMiscManualOptions = CreateMenuItem("Manual Options", "Alt+Shift+M", menuItem_Click, "frame silent");
menuMiscResetConfiguration = CreateMenuItem("Reset Configuration", "Alt+Shift+D0", menuItem_Click, "frame silent");
menuMiscGoToFolder = CreateMenuItem("Go to Folder", "Control+D0", menuItem_Click, "frame silent");
menuMiscGoToSpecialFolder = CreateMenuItem("Go to Special Folder", "Control+Shift+D0", menuItem_Click, "frame silent");
menuMiscWordWrap = CreateMenuItem("&Word Wrap", "Control+W", menuItem_Click, "child speak");
menuMiscUnwrap = CreateMenuItem("Unwrap", "Control+Shift+W", menuItem_Click, "child speak");
menuMiscExtraSpeechToggle = CreateMenuItem("Extra Speech Toggle", "Control+Shift+X", menuItem_Click, "frame silent");
menuMiscExtraSpeechLog = CreateMenuItem("Extra Speech Log", "Alt+Shift+X", menuItem_Click, "frame speak");
menuMiscEnvironmentVariables = CreateMenuItem("&Environment Variables ...", "Control+E", menuItem_Click, "frame speak");
menuMiscSpellCheck = CreateMenuItem("Spell Check", "F7", menuItem_Click, "child speak");
menuMiscThesaurus = CreateMenuItem("Thesaurus", "Shift+F7", menuItem_Click, "child speak");
menuMiscLookupTerm = CreateMenuItem("Lookup Term", "Alt+F7", menuItem_Click, "frame silent");
menuMiscTranslateLanguage = CreateMenuItem("Translate Language", "Alt+Shift+F7", menuItem_Click, "frame speak");
menuMiscGuardDocument = CreateMenuItem("Guard Document", "Control+F7", menuItem_Click, "child speak");
menuMiscNoGuard = CreateMenuItem("No Guard", "Control+Shift+F7", menuItem_Click, "child speak");
menuMiscPyBrace = CreateMenuItem("PyBrace", "Alt+Shift+OemOpenBrackets", menuItem_Click, "child speak");
menuMiscPyDent = CreateMenuItem("PyDent", "Alt+OemOpenBrackets", menuItem_Click, "child speak");
menuMiscInferIndent = CreateMenuItem("Infer Indent", "Alt+OemCloseBrackets", menuItem_Click, "child silent");
menuMiscFormatCode = CreateMenuItem("Format Code", "Control+D4", menuItem_Click, "child speak");
menuMiscRepeatLine = CreateMenuItem("Repeat Line", "Control+Y", menuItem_Click, "child speak");
menuMiscSectionBreak = CreateMenuItem("Section Break", "Control+Enter", menuItem_Click, "child speak");
menuMiscPathToClipboard = CreateMenuItem("Path to Clipboard", "Alt+Shift+P", menuItem_Click, "child speak");
menuMiscPathList = CreateMenuItem("Path List", "Control+Shift+P", menuItem_Click, "frame speak");
menuMiscInsertTime = CreateMenuItem("Insert Time", "Alt+Shift+OemSemicolon", menuItem_Click, "child speak");
menuMiscCalculateDate = CreateMenuItem("Calculate Date ...", "Control+Shift+OemSemicolon", menuItem_Click, "child silent");
menuMiscHTMLFormat = CreateMenuItem("HTML Format", "Control+H", menuItem_Click, "child speak");
menuMiscMarkdownToText = CreateMenuItem("Markdown to Plain Text", "", menuItem_Click, "child speak");
menuMiscHtmlToMarkdown = CreateMenuItem("HTML to Markdown", "", menuItem_Click, "child speak");
menuMiscHtmlToText = CreateMenuItem("HTML to Plain Text", "", menuItem_Click, "child speak");
menuMiscPreviewMarkdown = CreateMenuItem("Preview Markdown", "Control+F9", menuItem_Click, "child silent");
menuMiscPreviewMarkdownBrowser = CreateMenuItem("Preview Markdown in Web Browser", "", menuItem_Click, "child silent");
menuMiscCheckMarkdown = CreateMenuItem("Check Markdown", "Alt+F9", menuItem_Click, "child speak");
menuMiscRunCodeBlocks = CreateMenuItem("Run Code Blocks", "Alt+Shift+F9", menuItem_Click, "child speak");
menuMiscChatWithAI = CreateMenuItem("Chat with AI", "F12", menuItem_Click, "child speak");
menuMiscChatWithDocument = CreateMenuItem("Chat about Document", "Shift+F12", menuItem_Click, "child speak");
menuMiscTextConvert = CreateMenuItem("&Text Convert", "Control+T", menuItem_Click, "child speak");
menuMiscTextCombine = CreateMenuItem("Text Combine", "Control+Shift+T", menuItem_Click, "child speak");
menuMiscTextContents = CreateMenuItem("Text Contents", "Alt+Shift+T", menuItem_Click, "child speak");
menuMiscYieldWithRegExp = CreateMenuItem("Yield with Regular Expression ...", "Control+Shift+Y", menuItem_Click, "child silent");
//menuMiscYieldWithRegExp.ShortcutKeys = Util.String2Key("Control+Shift+Y");
menuMiscExtractWithRegExp = CreateMenuItem("Extract with Regular Expression ...", "Control+Shift+E", menuItem_Click, "child silent");
menuMiscRunAtCursor = CreateMenuItem("Run at Cursor ...", "Shift+F5", menuItem_Click, "child silent");
menuMiscSpecialCharacter = CreateMenuItem("Special Character ...", "F2", menuItem_Click, "child silent");
menuMiscEvaluateExpression = CreateMenuItem("Evaluate Expression", "Control+Oemplus", menuItem_Click, "child speak");
menuMiscReplaceTokens = CreateMenuItem("Replace Tokens", "Control+Shift+Oemplus", menuItem_Click, "child silent");
menuMiscTransformFiles = CreateMenuItem("Transform Files", "Alt+Oemplus", menuItem_Click, "child speak");
menuMiscGoToEnvironment = CreateMenuItem("Go to Environment", "Control+Shift+G", menuItem_Click, "frame speak");
menuMiscCompile = CreateMenuItem("Compile", "Control+F5", menuItem_Click, "child speak");
menuMiscPickCompiler = CreateMenuItem("Pick Compiler", "Control+Shift+F5", menuItem_Click, "frame silent");
menuMiscPromptCommand = CreateMenuItem("Prompt Command", "Alt+F5", menuItem_Click, "child silent");
menuMiscReviewOutput = CreateMenuItem("Review Output", "Alt+Shift+F5", menuItem_Click, "child speak");
menuMiscSaveSnippet = CreateMenuItem("Save Snippet", "Alt+S", menuItem_Click, "child speak");
menuMiscInvokeSnippet = CreateMenuItem("Invoke Snippet", "Alt+V", menuItem_Click, "child speak");
menuMiscViewSnippet = CreateMenuItem("View Snippet", "Alt+Shift+V", menuItem_Click, "frame speak");
menuMiscKeepUniqueItems = CreateMenuItem("Keep Unique Items", "Alt+Shift+K", menuItem_Click, "child speak");
menuMiscNumberItems = CreateMenuItem("Number Items ...", "Alt+Shift+N", menuItem_Click, "child silent");
menuMiscOrderItems = CreateMenuItem("Order Items", "Alt+Shift+O", menuItem_Click, "child speak");
menuMiscReverseItems = CreateMenuItem("Reverse Items", "Alt+Shift+Z", menuItem_Click, "child speak");
menuMiscListDifferentItems = CreateMenuItem("List Different Items", "Alt+Shift+L", menuItem_Click, "child speak");
menuMiscQueryCommonItems = CreateMenuItem("Query Common Items", "Alt+Shift+Q", menuItem_Click, "child speak");
menuMiscExplorerFolder = CreateMenuItem("Explorer Folder", "Alt+Oem5", menuItem_Click, "frame speak");
menuMiscCommandPrompt = CreateMenuItem("Command Prompt", "Control+Oem5", menuItem_Click, "frame speak");
menuMiscBurnToCD = CreateMenuItem("Burn to CD", "Alt+Shift+B", menuItem_Click, "child speak");
menuMiscWebDownload = CreateMenuItem("Web Download", "Alt+Shift+W", menuItem_Click, "frame speak");
menuMisc.DropDownItems.AddRange(new ToolStripItem[] {menuMiscSetDefaultFont, menuMiscConfigurationOptions, menuMiscManualOptions, menuMiscResetConfiguration, menuMiscGoToFolder, menuMiscGoToSpecialFolder, menuMiscWordWrap, menuMiscUnwrap, menuMiscExtraSpeechToggle, menuMiscExtraSpeechLog, menuMiscEnvironmentVariables, menuMiscSpellCheck, menuMiscThesaurus, menuMiscLookupTerm, menuMiscTranslateLanguage, menuMiscGuardDocument, menuMiscNoGuard, menuMiscPyBrace, menuMiscPyDent, menuMiscInferIndent, menuMiscFormatCode, menuMiscRepeatLine, menuMiscSectionBreak, menuMiscPathToClipboard, menuMiscPathList, menuMiscInsertTime, menuMiscCalculateDate, menuMiscHTMLFormat, menuMiscMarkdownToText, menuMiscHtmlToMarkdown, menuMiscHtmlToText, menuMiscPreviewMarkdown, menuMiscPreviewMarkdownBrowser, menuMiscCheckMarkdown, menuMiscRunCodeBlocks, menuMiscChatWithAI, menuMiscChatWithDocument, menuMiscTextConvert, menuMiscTextCombine, menuMiscTextContents, menuMiscYieldWithRegExp, menuMiscExtractWithRegExp, menuMiscRunAtCursor, menuMiscSpecialCharacter, menuMiscEvaluateExpression, menuMiscReplaceTokens, menuMiscTransformFiles, menuMiscGoToEnvironment, menuMiscCompile, menuMiscPickCompiler, menuMiscPromptCommand, menuMiscReviewOutput, menuMiscSaveSnippet, menuMiscInvokeSnippet, menuMiscViewSnippet, menuMiscKeepUniqueItems, menuMiscNumberItems, menuMiscOrderItems, menuMiscReverseItems, menuMiscListDifferentItems, menuMiscQueryCommonItems, menuMiscExplorerFolder, menuMiscCommandPrompt, menuMiscBurnToCD, menuMiscWebDownload});
//Dialog.Show("Misc.", menuMisc.DropDownItems.Count);

menuWindow = CreateMenu("&Window");
menuWindowNext = CreateMenuItem("Next Window", "Control+Tab", menuItem_Click, "child speak");
menuWindowPrior = CreateMenuItem("Prior Window", "Control+Shift+Tab", menuItem_Click, "child speak");
menuWindowArrangeIcons = CreateMenuItem("Arrange Icons", "Alt+F11", menuItem_Click, "child speak");
menuWindowCascade = CreateMenuItem("Cascade", "Control+F11", menuItem_Click, "child speak");
menuWindowTileHorizontal = CreateMenuItem("Tile Horizontal", "Alt+Shift+F11", menuItem_Click, "child speak");
menuWindowTileVertical = CreateMenuItem("Tile Vertical", "Control+Shift+F11", menuItem_Click, "child speak");
menuWindow.DropDownItems.AddRange(new ToolStripMenuItem[] {menuWindowNext, menuWindowPrior, menuWindowArrangeIcons, menuWindowCascade, menuWindowTileHorizontal, menuWindowTileVertical});
//Dialog.Show("Window.", menuWindow.DropDownItems.Count);

menuHelp = CreateMenu("&Help");
menuHelpAbout = CreateMenuItem("&About ...", "Alt+F1", menuItem_Click, "frame silent");
menuHelpDocumentation = CreateMenuItem("Documentation", "F1", menuItem_Click, "frame speak");
menuHelpTutorial = CreateMenuItem("Tutorial", "Control+Shift+F1", menuItem_Click, "frame speak");
menuHelpHistoryOfChanges = CreateMenuItem("History of Changes", "Shift+F1", menuItem_Click, "frame speak");
menuHelpSamplePrograms = CreateMenuItem("Sample Programs ...", "Control+Shift+F2", menuItem_Click, "frame speak");
menuHelpCopyLog = CreateMenuItem("Copy Log", "Control+F12", menuItem_Click, "frame speak");
menuHelpKeyDescriber = CreateMenuItem("Key Describer", "Control+F1", menuItem_Click, "frame silent");
menuHelpHotKeySummary = CreateMenuItem("Hotkey Summary", "Alt+Shift+H", menuItem_Click, "frame speak");
menuHelpAlternateMenu= CreateMenuItem("Alternate Menu ...", "Alt+F10", menuItem_Click, "frame silent");
menuHelpContextMenu= CreateMenuItem("Context Menu ...", "Shift+F10", menuItem_Click, "child silent");
menuHelpSendToMenu= CreateMenuItem("SendTo Menu ...", "Control+F10", menuItem_Click, "child silent");
menuHelpElevateVersion = CreateMenuItem("Elevate Version", "F11", menuItem_Click, "frame speak");
menuHelp.DropDownItems.AddRange(new ToolStripItem[] {menuHelpAbout, menuHelpDocumentation, menuHelpTutorial, menuHelpHistoryOfChanges, menuHelpSamplePrograms, menuHelpCopyLog, menuHelpKeyDescriber, menuHelpHotKeySummary, menuHelpAlternateMenu, menuHelpContextMenu, menuHelpSendToMenu, menuHelpElevateVersion});
//Dialog.Show("Help.", menuHelp.DropDownItems.Count);

menuMain.Items.AddRange(new ToolStripItem[] {menuFile, menuEdit, menuDelete, menuNavigate, menuQuery, menuMisc, menuWindow, menuHelp});
//menuMain.Items.AddRange(new ToolStripItem[] {menuFile, menuEdit, menuDelete, menuNavigate, menuQuery, menuMisc, menuHelp});
statusBar = CreateStatusBar();
// this.Controls.AddRange(new Control[] {menuMain, statusBar});
this.Controls.AddRange(new Control[] {statusBar, menuMain});
this.MainMenuStrip = menuMain;
menuMain.MdiWindowListItem = menuWindow;
//this.AutoSize = true;
this.Size = new Size(600, 600);
this.StartPosition = FormStartPosition.CenterScreen;
this.Text = "EdSharp";
this.ResumeLayout();
this.KeyPreview = true;
//this.MdiChildActivate += delegate(object o, EventArgs e) {this.Child = (MdiChild) this.ActiveMdiChild;};
string s = App.ReadOption("MaximizeWindow", "N").Trim().ToUpper();
if (s == "Y" || s == "YES") this.Shown += delegate(object o, EventArgs e) {
this.WindowState = FormWindowState.Maximized;
this.Activate();
Win32.SetForegroundWindow(this.Handle);
};
this.Shown += delegate(object o, EventArgs e) {
Util.ActivateTitle(this.Text);
};

string sDir = Directory.GetCurrentDirectory();
string sFile = Path.Combine(App.DataDir, App.ReadData("Compiler", "Default") + ".ini");
s = Ini.ReadValue(sFile, "Data", "Directory", "");
if (Directory.Exists(s) && !Util.Equiv(sDir, s)) {
//Dialog.Show(sDir, s);
//Directory.SetCurrentDirectory(s);
AddMessage("Folder " + Path.GetFileName(s));
Directory.SetCurrentDirectory(s);
}

} // MdiFrame constructor

public void SetStatus(object o) {
string sText = o.ToString();
this.statusBar.Items[0].Text = sText;
} // SetStatus method

public string GetNoNameTitle() {
object[] children = this.MdiChildren;
List<int> list = new List<int>();
foreach (object o in children) {
MdiChild child = (MdiChild) o;
string sTitle = child.Text;
if (sTitle.StartsWith("NoName") && Path.GetExtension(sTitle).Length == 0) {
string s = sTitle.Substring(6);
int i = Int32.Parse(s);
list.Add(i);
}
}

int iTitle = 0;
for (int i = 1; i <= children.Length; i++) {
if (!list.Contains(i)) {
iTitle = i;
break;
}
}

if (iTitle == 0) iTitle = children.Length + 1;
string sReturn = "NoName" + iTitle.ToString();
return sReturn;
} // GetNoNameTitle

public bool ProcessCmdKey_Helper(ref Message msg, Keys keyData) {
string sKey = keyData.ToString();
int iIndex = -1;
if (this.Child != null) iIndex = this.Child.RTB.Index;
if (sKey.StartsWith("Menu,") || sKey.StartsWith("ControlKey,") || sKey.StartsWith("ShiftKey,")) sKey = "";
else if (iIndex == this.KeyIndex && sKey == this.KeyString) this.KeyRepeat += 1;
else this.KeyRepeat = 0;
if (sKey.Length > 0) {
this.KeyString = sKey;
this.KeyIndex = iIndex;
}
// Util.Say(keyData.ToString());
//Clipboard.SetText(Clipboard.GetText() + keyData.ToString() + "\r\n");
// Util.Say("Repeat " + this.KeyRepeat);

ToolStripMenuItem menuItem;
if (keyData == Keys.F9) {
HomerRichTextBox rtb = App.Frame.Child.RTB;
string sText = rtb.GetRange(rtb.Index, rtb.TextLength);
//Dialog.Show(App.TempFile, sText);
/*
sText = Util.ConvertQuotes(sText);
Util.String2FileA(sText, App.TempFile);
*/
//File.WriteAllText(App.TempFile, sText, Encoding.GetEncoding(0, null, null));
File.WriteAllText(App.TempFile, sText, Encoding.GetEncoding(0));
//File.WriteAllText(App.TempFile, sText, Encoding.GetEncoding("US-ASCII", new EncoderReplacementFallback(" "), new DecoderReplacementFallback(" ")));
//File.WriteAllText(App.TempFile, sText, Encoding.GetEncoding("US-ASCII", null, null));
// Win32.JFWRunFunction("SayAllTempFile");
COM.JFWRunFunction("SayAllTempFile");
base.ProcessCmdKey (ref msg, keyData);
return true;
}
else if (keyData == (Keys.Shift | Keys.F9)) {
string s = Util.File2String(App.TempFile);
if (!Util.IsNumeric(s)) return true;
int i = Int32.Parse(s);
HomerRichTextBox rtb = App.Frame.Child.RTB;
rtb.Index += i;
base.ProcessCmdKey (ref msg, keyData);
return true;
}

else if (keyData == (Keys.Insert)) {
// Util.Say("Insert key now");
return true;
}

else if (hashKey.TryGetValue(keyData, out menuItem)) {
//if (this.Child != null && !this.Child.RTB.IndentMode && menuItem == menuEditEnterNewLine) return base.ProcessCmdKey (ref msg, keyData);
this.bCommandComplete = false;
menuItem.PerformClick();
this.bCommandComplete = true;
return true;
}
else return base.ProcessCmdKey (ref msg, keyData);
} // ProcessCmdKey_Helper method

protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
return this.ProcessCmdKey_Helper(ref msg, keyData);
} // ProcessCmdKey handler

static MenuStrip CreateMainMenu() {
MenuStrip menuMain = new MenuStrip();
menuMain.AccessibleRole = AccessibleRole.MenuBar;
//menuMain.AutoSize = true;
//menuMain.CanOverflow = false;
menuMain.Dock = DockStyle.Top;
//menuMain.LayoutStyle = ToolStripLayoutStyle.Flow;
//menuMain.Stretch = false;
return menuMain;
} // CreateMainMenu method

static ToolStripMenuItem CreateMenu(string sText) {
ToolStripMenuItem menuItem = new ToolStripMenuItem(sText);
menuItem.AccessibleRole = AccessibleRole.MenuItem;
return menuItem;
} // CreateMenu method

static ToolStripMenuItem CreateMenuItem(string sText, string sKey, EventHandler eh) {
bool bFrame = false;
return CreateMenuItem(sText, sKey, eh, bFrame);
} // CreateMenuItem method

static ToolStripMenuItem CreateMenuItem(string sText, string sKey, EventHandler eh, bool bFrame) {
string sOptions = "";
if (bFrame) sOptions += "frame ";
if (sText.EndsWith(" ...") || sText.EndsWith("Again")) sOptions += "silent ";
return CreateMenuItem(sText, sKey, eh, sOptions);
}  // CreateMenuItem method

static ToolStripMenuItem CreateMenuItem(string sText, string sKey, EventHandler eh, string sOptions) {
ToolStripMenuItem menuItem = new ToolStripMenuItem(sText, null, eh);
menuItem.AccessibleRole = AccessibleRole.MenuItem;
menuItem.Tag = sOptions;

string sCommand = sText.Replace("&", "").Replace("...", "").Trim();
menuItem.Name = sCommand;
// Keybindings are defined directly in code, via the sKey argument. The
// former per-item [Keys] override read from EdSharp.ini has been removed:
// it cost an INI read for every menu item at startup, and runtime
// rebinding is no longer supported (defaults are chosen to be optimal).
sKey = sKey.Replace("&", "");
Keys keyData = Util.String2Key(sKey);
// Keys.None means a menu-only command: no shortcut text and no entry in
// the shortcut table, where several unbound items would otherwise alert as
// duplicates of one another at startup. KeyMap.register below already
// handles unbound commands, so it still runs.
if (keyData == Keys.None) {
menuItem.Text = sText;
}
else if (hashKey.ContainsKey(keyData)) {
string s = hashKey[keyData].Name;
Dialog.Show("Alert", "Cannot assign " + sKey + " to " + sCommand + ",\nsince already assigned to " + s);
}
else {
string sFriendlyKey = Util.GetFriendlyKeyName(sKey);
menuItem.ShortcutKeyDisplayString = sFriendlyKey;
menuItem.AccessibleName = sText.Replace("&", "") + "   " + sFriendlyKey;
menuItem.Text = sText;
hashKey.Add(keyData, menuItem);
}

// Register the command in the central KeyMap (Homer): command -> key,
// owning menu item, and UI context (the sOptions string, e.g. "frame
// speak"). Additive -- hashKey above is unchanged; KeyMap becomes the
// single table the status bar, Key Describer, and Alternate Menu read.
KeyMap.register(sCommand, keyData, menuItem, sOptions);

menuItem.Paint += delegate(object oSender, PaintEventArgs e) {
// if (!App.Frame.KeyDescriber) return;
foreach (ToolStripMenuItem menu in App.Frame.menuMain.Items) {
foreach (object o in menu.DropDownItems) {
ToolStripMenuItem item = o as ToolStripMenuItem;
if (item == null) continue;
if (!item.Selected) continue;
string[] aSummary = App.Frame.GetKeySummary(item);
string sSummary = aSummary[0] + " = " + aSummary[1] + ", " + aSummary[2];
string sDescription = aSummary[2];
if (sDescription != App.Frame.LastDescription) {
// System.Threading.Thread.Sleep(1000);
// Util.Say(sDescription);
App.Frame.SetStatus(sDescription);
App.Frame.LastDescription = sDescription;
}
break;
}
}
};

return menuItem;
} // CreateMenuItem method

static StatusStrip CreateStatusBar() {
StatusStrip sb = new StatusStrip();
sb.AccessibleRole = AccessibleRole.StatusBar;
sb.SuspendLayout();
ToolStripStatusLabel lblStatus = new ToolStripStatusLabel("Ready");
lblStatus.AutoSize = true;
sb.Items.AddRange(new ToolStripItem[] {lblStatus});
sb.AutoSize = true;
sb.Dock = DockStyle.Bottom;
sb.ResumeLayout();
return sb;
} // CreateStatusBar method

public string GetPercentAddress(HomerRichTextBox rtb) {
return String.Format("Line {0}   Column {1}   Percent{2}", rtb.Line, rtb.Column, rtb.Percent);
} // GetPercentAddress method

public string GetPageAddress(HomerRichTextBox rtb) {
string sText = rtb.Text;
int iIndex = rtb.Index;
sText = sText.Substring(0, iIndex);
int iPage = sText.Length - sText.Replace("\f", "").Length + 1;
iIndex = sText.LastIndexOf("\f");
if (iIndex >= 0) sText = sText.Substring(iIndex);
if (sText.StartsWith("\f")) sText = sText.Remove(0, 1);
int iLine = sText.Length - sText.Replace("\n", "").Length + 1;
iIndex = sText.LastIndexOf("\n");
int iColumn = sText.Length - iIndex;
return String.Format("Page {0}   Line {1}   Column {2}", iPage, iLine, iColumn);
} // GetPageAddress method

public void SetStatusAddress(object sender, EventArgs e) {
if (sender != null && !this.bCommandComplete) return;
if (this.Child == null) {
SetStatus("");
return;
}

HomerRichTextBox rtb = this.Child.RTB;
//string sText = String.Format("Line {0}\tColumn {1}\tPercent{2}", rtb.Line, rtb.Column, rtb.Percent);
string sText = "";
int iIndex = -1;
bool bPageAddress = true;
if (App.ReadOption("HardPageAddress", "N").ToLower().Substring(0, 1) != "y") bPageAddress = false;
if (bPageAddress) sText = GetPageAddress(rtb);
else  sText = GetPercentAddress(rtb);

iIndex = rtb.Index;
char c = ' ';
if (iIndex < rtb.TextLength) c = rtb.Text[iIndex];

int iNewIndex = rtb.Index;
int iNewTextLength = rtb.TextLength;
int iDelta = Math.Abs(iNewIndex - rtb.OldIndex);
if (!bPageAddress || iDelta != 1 || iNewTextLength != rtb.OldTextLength) {} // Do nothing
else if (c == '\f') Util.Say("FormFeed");
else if (c == '\n') Util.Say("LineFeed");
else if (c == '\t') Util.Say("TabChar");
rtb.OldIndex = iNewIndex;
rtb.OldTextLength = iNewTextLength;

if (sender == null) Util.Say(sText);
this.SetStatus(sText);

if (!rtb.IndentMode) return;
string sLine = rtb.RowText.Trim();
if (sLine.Length == 0) return;
string sComment = App.ReadOption("QuotePrefix", "> ");
if (sLine.StartsWith(sComment)) return;
int iLevels = GetIndent();
// Indent Mode announcements ("In 1", "Out 2") are spoken directly by
// EdSharp through the Homer speech subsystem, which dispatches to the
// JAWS COM interface, the NVDA controller client, or a native UIA
// notification -- whichever screen reader is running. This retires the
// old synchronization protocol, which wrote each delta into
// IndentMode.tmp under an IndentChange key for a JAWS script to read
// back and speak. That design had three faults this history preserves:
// the file was truncated on EVERY cursor move where the level was
// unchanged, an unguarded write that could collide with the script's
// read; the script could read a stale delta and announce the previous
// line's change on the wrong line; and avoiding double speech required
// the obscure convention of a hyphen inside the ExtraSpeech option to
// mute EdSharp's own voice. (An even older mechanism, a per-user
// EdSharpIndent environment variable, had already been retired.)
// IndentMode.tmp itself remains -- created and deleted as a pure mode
// flag when Indent Mode toggles or a window gains focus -- so screen
// reader scripts can still adapt key behavior to the mode; it just no
// longer carries data. The hyphen convention still silences these
// announcements for anyone who prefers quiet.
if (rtb.IndentLevels == iLevels) return;
string sDelta = GetDelta(rtb.IndentLevels, iLevels);
if (App.IndentChange) Util.Say(sDelta);
rtb.IndentLevels = iLevels;
} // SetStatusAddress method

public void AddMessage(object oText) {
bool bGlobal = false;
AddMessage(oText, bGlobal);
} // AddMessage method

public void AddMessage(object oText, bool bGlobal) {
string sText = oText.ToString();
Util.Say(sText, bGlobal);
if (App.CaptureOutput) Util.StringAppend2File(sText + "\r\n", App.TempFile);
//sText = this.statusBar.Items[0].Text + "\t" + sText;
sText = this.statusBar.Items[0].Text + "   " + sText;
SetStatus(sText);
} // AddMessage method

public void SetMessage(object oText) {
SetStatus(oText);
Util.Say(oText);
} // SetMessage method

public void GetRowAndCol(out int iRow, out int iCol) {
HomerRichTextBox rtb = this.Child.RTB;
int iIndex = rtb.SelectionStart + rtb.SelectionLength;
iRow = rtb.GetLineFromCharIndex(iIndex);
iCol = iIndex - rtb.GetFirstCharIndexOfCurrentLine();
} // GetRowAndCol method

bool IsEmptyWindow() {
return !(this.Child == null || this.Child.RTB.Modified || this.Child.RTB.TextLength > 0);
} // IsEmptyWindow method

public bool IsCharacter() {
int iIndex;
return IsCharacter(out iIndex);
} // IsCharacter method

public bool IsCharacter(out int iIndex) {
HomerRichTextBox rtb = this.Child.RTB;
iIndex = rtb.Index;
if (iIndex >= rtb.TextLength) {
AddMessage("No character at cursor!");
return false;
}
else return true;
} // IsCharacter method

// The indentation unit that actually governs the current document: the
// leading whitespace of its first indented line whose prefix repeats a
// single character (four spaces, two spaces, or a tab, whatever the
// file uses), scanning at most a few hundred lines; when no indented
// line exists yet, the IndentUnit setting, and failing that two spaces.
// Every indentation command and announcement below uses this, so a
// four-space Python file reports level 2 where the old fixed setting
// of two spaces reported level 4 -- and Tab inserts what the file uses.
// The Infer Indent command still lets you configure the setting
// explicitly, which then governs documents with no indentation yet.
public string GetIndentUnit() {
if (this.Child != null) {
string[] aLines = this.Child.RTB.Text.Replace("\r\n", "\n").Split('\n');
int iLimit = Math.Min(aLines.Length, 400);
for (int iLine = 0; iLine < iLimit; iLine++) {
string sLine = aLines[iLine];
string sTrim = sLine.TrimStart();
if (sTrim.Length == 0) continue;
int iWs = sLine.Length - sTrim.Length;
if (iWs == 0) continue;
string sWs = sLine.Substring(0, iWs);
bool bUniform = true;
foreach (char c in sWs) if (c != sWs[0]) { bUniform = false; break; }
if (bUniform) return sWs;
}
}
string sUnit = Util.Literalize(App.ReadOption("IndentUnit", "\t"));
if (sUnit.Length == 0) sUnit = "\t";
return sUnit;
} // GetIndentUnit method

public int GetIndent() {
return GetIndent(this.Child.RTB.Row);
} // GetIndent method

public int GetIndent(int iRow) {
string sIndent = GetIndentUnit();
MdiChild child = this.Child;
HomerRichTextBox rtb = child.RTB;
string sLine = rtb.GetRowText(iRow);
int iLength = sIndent.Length;
int iLevels = 0;
while (sLine.StartsWith(sIndent)) {
iLevels++;
if (sLine.Length == iLength) sLine = "";
else sLine = sLine.Substring(iLength);
}
return iLevels;
} // GetIndent method

public string GetDelta(int iBefore, int iAfter) {
if (iBefore < iAfter) return "In " + (iAfter - iBefore);
else return "Out " + (iBefore - iAfter);
} // GetDelta method

public string GetStyleText() {
HomerRichTextBox rtb = this.Child.RTB;
string sText = "";
if (rtb.SelectionFont.Bold) sText += "Bold ";
if (rtb.SelectionFont.Italic) sText += "Italic ";
if (rtb.SelectionFont.Underline) sText += "Underline";
sText = sText.Trim();
if (sText.Length == 0) sText = "Regular";
return sText;
} // GetStyleText method

public string GetJustifyText() {
HomerRichTextBox rtb = this.Child.RTB;
string sText = "Left";
HorizontalAlignment ha = rtb.SelectionAlignment;
if (rtb.SelectionBullet) sText = "Bullet";
else if (ha == HorizontalAlignment.Center) sText = "Center";
else if (ha == HorizontalAlignment.Right) sText = "Right";
return sText;
} // GetJustifyText method

public string GetBaselineText() {
HomerRichTextBox rtb = this.Child.RTB;
string sText = "Flat";
int iOffset = rtb.SelectionCharOffset;
if (iOffset < 0) sText = "Down";
else if (iOffset > 0) sText = "Up";
return sText;
} // GetBaselineText method

public string GetFontText(Font font, Color color) {
string sFont = Util.Font2String(font);
string sColor = Util.Color2String(color);
sFont += ", Color=" + sColor;
return sFont;
} // GetFontText method

// Walk the Samples folder and its subfolders, building one row per
// sample: the file's name, its folder when it is in one, and a short
// description. Folders are visited in name order and files within them
// likewise, so the list reads the same way every time.
void collectSamples(string sDir, string sPrefix, List<object> lPaths, List<string> lDisplay) {
string[] aFiles = Directory.GetFiles(sDir);
Array.Sort(aFiles);
foreach (string sFile in aFiles) {
string sName = Path.GetFileName(sFile);
if (sName.ToLower() == "readme.md") continue;
string sWhat = describeSample(sFile);
string sRow = (sPrefix.Length > 0) ? sPrefix + ": " + sName : sName;
if (sWhat.Length > 0) sRow += " -- " + sWhat;
lPaths.Add(sFile);
lDisplay.Add(sRow);
}
string[] aDirs = Directory.GetDirectories(sDir);
Array.Sort(aDirs);
foreach (string sSubDir in aDirs) collectSamples(sSubDir, Path.GetFileName(sSubDir), lPaths, lDisplay);
} // collectSamples method

// A sample's first comment line, cleaned of its comment marks: the
// programs already explain themselves on their first line, so the list
// takes the description from the sample rather than from a table
// someone must remember to update.
string describeSample(string sFile) {
try {
foreach (string sLine in File.ReadLines(sFile)) {
string sTrim = sLine.Trim();
if (sTrim.Length == 0) continue;
if (sTrim.StartsWith("<!DOCTYPE") || sTrim.StartsWith("<html") || sTrim.StartsWith("<head") || sTrim.StartsWith("<meta")) continue;
if (sTrim.StartsWith("<title>")) return sTrim.Replace("<title>", "").Replace("</title>", "").Trim();
string sText = sTrim;
if (sText.StartsWith("//")) sText = sText.Substring(2);
else if (sText.StartsWith("rem ")) sText = sText.Substring(4);
else if (sText.StartsWith("\"\"\"")) sText = sText.Substring(3);
else if (sText.StartsWith("#")) sText = sText.TrimStart('#');
else continue;
sText = sText.Trim();
// "fruitBasket.py -- the fruit basket program in Python": the part
// after the dashes is the description.
int iDash = sText.IndexOf(" -- ");
if (iDash >= 0) sText = sText.Substring(iDash + 4);
if (sText.Length > 80) sText = sText.Substring(0, 77) + "...";
return sText.TrimEnd('.');
}
}
catch (Exception) {}
return "";
} // describeSample method

public string[] GetSnippetFiles(out string[] aValues) {
string sBaseDir = @"Snippets\" + App.ReadData("Compiler", "Default");
string sDir = Path.Combine(App.DataDir, sBaseDir);
if (!Directory.Exists(sDir)) Directory.CreateDirectory(sDir);
string[] aResults = Directory.GetFiles(sDir);

List<string> listResults = new List<string>(aResults);
List<string> listFiles = new List<string>();
foreach (string s in aResults) listFiles.Add(Path.GetFileName(s).ToLower());

sBaseDir = @"Snippets\Default";
sDir = Path.Combine(App.DataDir, sBaseDir);
if (!Directory.Exists(sDir)) Directory.CreateDirectory(sDir);
aResults = Directory.GetFiles(sDir);
foreach (string s in aResults) if (!listFiles.Contains(Path.GetFileName(s).ToLower())) listResults.Add(s);

// Snippets shipped with the program (the installer places the Snippets
// folder in the program directory) join the list last, so a same-named
// snippet in the data directory always wins: the user's copy overrides
// the shipped one, and shipped snippets appear without any copying.
foreach (string s in listResults) if (!listFiles.Contains(Path.GetFileName(s).ToLower())) listFiles.Add(Path.GetFileName(s).ToLower());
string[] aProgramDirs = new string[] {Path.Combine(App.ProgramDir, @"Snippets\" + App.ReadData("Compiler", "Default")), Path.Combine(App.ProgramDir, @"Snippets\Default")};
foreach (string sProgramDir in aProgramDirs) {
if (!Directory.Exists(sProgramDir)) continue;
foreach (string s in Directory.GetFiles(sProgramDir)) {
string sName = Path.GetFileName(s).ToLower();
if (listFiles.Contains(sName)) continue;
listFiles.Add(sName);
listResults.Add(s);
}
}
aResults = listResults.ToArray();

// Each row reads as the snippet's name and, when it did not come from
// the shared Default folder, the folder it did come from: "Mermaid
// Flowchart.txt (Python)" or "... (shipped)". A snippet collection
// grows over years, and a list that says where each entry lives is far
// easier to keep in order than a flat one -- while the name itself
// still leads, so typing to match a snippet works exactly as before.
aValues = new string[aResults.Length];
string sProgramSnippets = Path.Combine(App.ProgramDir, "Snippets").ToLower();
for (int i = 0; i < aResults.Length; i++) {
string sName = Path.GetFileName(aResults[i]);
string sFolder = Path.GetFileName(Path.GetDirectoryName(aResults[i]));
bool bShipped = aResults[i].ToLower().StartsWith(sProgramSnippets);
if (bShipped) aValues[i] = sName + " (shipped)";
else if (String.Compare(sFolder, "Default", true) != 0) aValues[i] = sName + " (" + sFolder + ")";
else aValues[i] = sName;
}
return aResults;
} // GetSnippetFiles method

public void GetDateAndTime(out string sDate, out string sTime) {
DateTime dt = DateTime.Now;
string sDateFormat = App.ReadOption("DateFormat", "");
string sTimeFormat = App.ReadOption("TimeFormat", "");
if (sDateFormat == "0") sDate = "";
else sDate = (sDateFormat.Length > 0) ? dt.ToString(sDateFormat) : dt.ToLongDateString();

if (sTimeFormat == "0") sTime = "";
else sTime = (sTimeFormat.Length > 0) ? dt.ToString(sTimeFormat) : dt.ToShortTimeString();
} // GetDateAndTime method

public string ReplaceTokens(string sText) {
string[] aTokens = App.ReadSectionKeys("Tokens");
foreach (string sToken in aTokens) {
if (sText.IndexOf(sToken) == -1) continue;
string s = App.ReadValue("Tokens", sToken, "");
string sFile = GetSnippetDir() + @"\" + s;
if (File.Exists(sFile)) s = Util.File2String(sFile);
//Dialog.Show(sFile, s);
string sResult = Script.run(s);
if (sResult == null) sResult = "";
sText = sText.Replace("%" + sToken + "%",sResult);
}

return sText;
} // ReplaceTokens method

public void TransFormFiles() {
// A transform job is now a Regexer-style .inix file: one [Section] per task,
// each with Find, Replace, Options, Extract, and Divider keys, and values that
// may span multiple lines. Replace is processed like Regexer (Regex.Unescape,
// so \n, \t, and \" work; $# substitutes the running match count); Extract
// collects the matched text to the clipboard. Options is a comma-separated list
// of .NET RegexOptions names (multiline, ignorecase, singleline, compiled, ...).
// The current document supplies the list of source files, one path per line.
if (File.Exists(App.TempFile)) File.Delete(App.TempFile);
App.CaptureOutput = true;
HomerRichTextBox rtb = this.Child.RTB;

string sJob = App.ReadData("Job", "");
string sTransformFile = Dialog.OpenFile("Open Job", sJob);
if (sTransformFile.Length == 0) { App.CaptureOutput = false; return; }
App.WriteData("Job", sTransformFile);

string sJobBody = Util.File2String(sTransformFile);
string[] aJobLines = sJobBody.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');

List<InixCodec.Section> lsAll = InixCodec.parseLines(aJobLines);
List<InixCodec.Section> lsTasks = new List<InixCodec.Section>();
foreach (InixCodec.Section section in lsAll) {
if (section.get("Find") != null) lsTasks.Add(section);
}
if (lsTasks.Count == 0) {
Dialog.Show("Transform Files", "No tasks found.  A job is now an .inix file with one [Section] per task, each having Find, Replace, Options, Extract, and Divider keys.");
App.CaptureOutput = false;
return;
}

string[] aChoices = {"&Test", "&Run", "&Verbose"};
string sChoice = Dialog.Choose("Choose Mode", "", aChoices, 0);
if (sChoice.Length == 0) { App.CaptureOutput = false; return; }
bool bApply = (sChoice != "&Test");
bool bVerbose = (sChoice == "&Verbose");

string sSourceList = rtb.Text.Trim();
string[] aSourceLines = sSourceList.Split('\n');
string sDir = Directory.GetCurrentDirectory();
string sExtractText = "";
string sClipboardDivider = "\f\n";
int iExtractTotal = 0;

int iFileCount = 0;
foreach (string sSourceLine in aSourceLines) {
string sSourceFile = sSourceLine.Trim();
if (sSourceFile.Length == 0) continue;
// A line that is not a usable path is skipped rather than throwing.  The current
// document is expected to hold a list of files, but if it holds ordinary text
// instead, Path.GetDirectoryName would raise an exception on the first line
// containing a character that cannot appear in a path.
string sSourceDir;
try {
sSourceDir = Path.GetDirectoryName(sSourceFile);
if (sSourceDir.Length > 0 && Directory.Exists(sSourceDir)) sDir = sSourceDir;
else sSourceFile = Path.Combine(sDir, Path.GetFileName(sSourceFile));
}
catch (Exception) { continue; }
if (!File.Exists(sSourceFile)) continue;
iFileCount++;

AddMessage(Path.GetFileName(sSourceFile));
Encoding en = App.Frame.Child.GetYieldEncoding();
string sSourceBody = Util.File2String(sSourceFile, ref en);
bool bWasCrlf = sSourceBody.Contains("\r\n");
sSourceBody = sSourceBody.Replace("\r\n", "\n");
bool bChanged = false;

foreach (InixCodec.Section section in lsTasks) {
string sFind = section.get("Find");
if (sFind == null || sFind.Length == 0) continue;
string sReplace = section.get("Replace");
sReplace = Regex.Unescape(sReplace == null ? "" : sReplace);
RegexOptions options = Util.RegexOptionsFromString(section.get("Options"));
bool bExtract = Util.ToBool(section.get("Extract"));
string sDivider = section.get("Divider");
sDivider = (sDivider == null || sDivider.Length == 0) ? "\f\n" : Regex.Unescape(sDivider);

Regex rex;
try { rex = new Regex(sFind, options); }
catch (Exception ex) {
Dialog.Show("Error", section.Name + ": " + ex.Message);
App.CaptureOutput = false;
return;
}

if (bVerbose || !bApply) AddMessage(section.Name);
int iCount = 0;
if (!bApply) {
iCount = rex.Matches(sSourceBody).Count;
}
else if (bExtract) {
// Extract tasks collect matches to the clipboard and never modify the file.
foreach (Match m in rex.Matches(sSourceBody)) {
iCount++;
sExtractText += (sExtractText.Length > 0 ? sDivider : "") + m.ToString();
}
if (iCount > 0) { iExtractTotal += iCount; sClipboardDivider = sDivider; }
}
else {
// Replace / delete tasks rewrite the file. An empty Replace deletes matches;
// either way, any match means the content changed and the file must be saved.
sSourceBody = rex.Replace(sSourceBody, delegate(Match m) {
iCount++;
return m.Result(sReplace).Replace("$#", iCount.ToString());
});
if (iCount > 0) bChanged = true;
}
if (bVerbose || !bApply) AddMessage(Util.Pluralize(iCount, "match", "matches"));
}

if (bApply && bChanged) {
if (bWasCrlf) sSourceBody = sSourceBody.Replace("\n", "\r\n");
Util.String2File(sSourceBody, sSourceFile, ref en);
}
}

// Nothing happened because no file was found to work on.  Say so plainly: the
// document that is open supplies the LIST of files to transform, one path per
// line -- it is not itself the file being transformed.  Silently doing nothing
// is what makes this look broken.
if (iFileCount == 0) {
Dialog.Show("Transform Files", "No files were transformed.\n\nThe document that is open should contain the list of files to transform, one full path per line.  It is not itself the file that gets transformed.\n\nTo transform a single file, open a new document, type or paste that file's full path on the first line, and run Transform Files from there.");
App.CaptureOutput = false;
return;
}

if (sExtractText.Length > 0) {
sExtractText = sExtractText.Replace("\n", "\r\n");
try {
string sClip = Util.GetClipboardText();
if (sClip.Length > 0) sClip += sClipboardDivider.Replace("\n", "\r\n");
Util.SetClipboardText(sClip + sExtractText);
AddMessage(Util.Pluralize(iExtractTotal, "match", "matches") + " to clipboard");
}
catch {}
}

AddMessage("Done", true);
App.CaptureOutput = false;
} // TransForm files method

public static string GetSnippetDir() {
string sBaseDir = @"Snippets\" + App.ReadData("Compiler", "Default");
string sDir = Path.Combine(App.DataDir, sBaseDir);
if (!Directory.Exists(sDir)) Directory.CreateDirectory(sDir);
return sDir;
} // GetSnippetDir method

public string GetDirChoice() {
string[] aButtons = {"&Current", "&Program", "&Data", "&Snippet", "&Other"};
string sButton = Dialog.Choose("Choose Directory", "", aButtons, 0);
if (sButton.Length == 0) return "";

string sDir = Directory.GetCurrentDirectory();
switch (sButton) {
case "&Current" :
if (this.Child != null && this.Child.File.IndexOf(@"\") >= 0) sDir = Path.GetDirectoryName(this.Child.File);
break;
case "&Program" :
sDir = App.ProgramDir;
break;
case "&Data" :
sDir = App.DataDir;
break;
case "&Snippet" :
sDir = App.DataDir + @"\Snippets\" + App.ReadData("Compiler", "Default");
if (!Directory.Exists(sDir)) Directory.CreateDirectory(sDir);
break;
case "&Other" :
sDir = Dialog.OpenFolder("Open Folder", "Name", Directory.GetCurrentDirectory());
if (sDir.Length == 0) return "";
break;
}
return sDir;
} // GetDirChoice method

public int GetViewLevel(string sFile) {
// Decide whether a file is converted when it is opened from outside the editor
// (Windows Explorer, "Open with", the command line, or Recent Files).  A return
// of 0 opens the file raw; 1 converts it through the Import table.  The ordinary
// Open command, Control+O, always opens raw regardless of this value.
// Precedence: an explicit ViewLevels entry wins, so the user can force any
// extension either way (e.g. "docx:0" to see a Word file raw, or "rst:1" to
// convert reStructuredText); otherwise binary / document formats convert,
// because their raw bytes are unreadable (and especially noisy for a screen
// reader); otherwise every text, markup, data, source, or unknown format opens
// raw.  This guarantees no text format is auto-converted -- now or as new
// converters are added -- unless it is explicitly listed.
const string sBinaryFormats = "doc docx xls xlsx ppt pptx pdf epub epub3 hlp wpd rtf";
string sExt = Path.GetExtension(sFile).TrimStart('.').ToLower();
string sViewLevels = App.ReadOption("ViewLevels", "");
foreach (string sViewLevel in sViewLevels.Split(' ')) {
string[] aViewLevel = sViewLevel.Split(':');
if (aViewLevel.Length < 2) continue;
if (sExt != aViewLevel[0].Trim().ToLower()) continue;
try { return Int32.Parse(aViewLevel[1]); }
catch (Exception ex) { Dialog.Show("Error", ex.Message); }
}
foreach (string sBinaryFormat in sBinaryFormats.Split(' ')) if (sExt == sBinaryFormat) return 1;
return 0;
} // GetViewLevel method

// ===== Command descriptions ================================================
// Every command's key and description, in the program itself. They used
// to live only in Hotkeys.ini beside the executable, which had two
// faults: a command added to the code but not to the file answered "No
// description available" in Key Describer mode, and -- worse -- the
// installer leaves an existing Hotkeys.ini alone, so a computer that had
// EdSharp before kept an old file and never saw new descriptions at all.
// That is what happened with Chat with AI on 26 August 2026.
//
// So the table below is the source of truth, shipped inside the binary
// where it cannot go stale. Hotkeys.ini is still read, and still wins
// where it has an entry, so anyone who has customized a description
// keeps it; anything the file lacks is answered from here.
//
// Each line is the command name, a tab, then the key and the
// description separated by a comma -- the same shape as the file, so the
// two can be compared, and the audit script does compare them.
static readonly string[] c_aCommandSummaries = new string[] {
"Launch EdSharp\tAlt+Control+E, Launch or activate the EdSharp application via a Windows desktop shortcut",
"Hotkey Summary\tAlt+Shift+H, Display this list of command names, hot keys, and descriptions in a new window",
"Documentation\tF1, Open Documentation in web browser",
"About\tAlt+F1, Display version and release date",
"History of Changes\tShift+F1, Display list of fixes and improvements",
"Key Describer\tControl+F1, Toggle a mode in which pressing a key describes its action",
"Alternate Menu\tAlt+F10, Present all commands in a single, alphabetized list",
"Context Menu\tShift+F10, Pick a command from those available to Windows Explorer for the current file extension",
"SendTo Menu\tControl+F10, Pick a command from those available as Windows \"Send To\" options",
"Select All\tControl+A, Select all text",
"Unselect All\tControl+Shift+A, Clear text selection",
"Select Chunk\tControl+Space, Select contiguous sequence of non-blank characters at cursor, or select the next chunk if a selection already exists",
"Say Selected\tShift+Space, or JAWSKey+Shift+DownArrow, Say selected text, or spell if repeated",
"Say Chunk\tShift+BackSpace, Say chunk at cursor",
"Start Selection\tF8, Mark starting point of text to be selected",
"Complete Selection\tShift+F8, Select text from starting point to cursor",
"Reselect\tControl+Shift+F8, Reselect between previous start and end positions",
"Go to Start of Selection\tAlt+Shift+F8, Return to start position of selection",
"Copy All\tControl+F8, Copy all text to clipboard",
"Read All\tAlt+F8, Say all text (without moving cursor)",
"Say Address\tAlt+A, Say line, column, and percent position of cursor",
"Say Block\tAlt+B, Say the rest of the current code block, or the whole block if repeated",
"Say Braces\tAlt+Shift+], Say number of braces on either side of cursor",
"Say Indentation\tAlt+I, Say the indentation level of the current line, or the preceding line with less indentation if repeated",
"Say Yield\tAlt+Y, Say number of characters, words, and lines in all or selected text",
"Say Status\tAlt+Z, Say whether current file has been modified since last save to disk, or say its character encoding if repeated",
"Say Clipboard\tAlt+Apostrophe, Say clipboard text, or spell if repeated",
"Say Time\tAlt+Semi-colon, Say current time and date",
"Insert Time\tAlt+Shift+Semi-colon, Insert current time and date",
"Calculate Date\tControl+Shift+Semi-colon, Calculate and insert date",
"Configuration Options\tAlt+Shift+C, Adjust configuration options through a dialog",
"Set Default Font and Color\tAlt+Shift+Equals, Set default font and color for editing window",
"Manual Options\tAlt+Shift+M, Adjust options by directly editing the main configuration file",
"Reset Configuration\tAlt+Shift+0, Revert to default options, or define a new compiler configuration",
"Copy\tControl+C, Copy selected text to clipboard, or copy current line if no selection",
"Copy Append\tAlt+C, Append selected text to clipboard, or append current line if no selection",
"Copy Rich Text\tControl+Shift+C, Copy selected text with formatting to clipboard",
"Cut\tControl+X, Cut selected text to clipboard, or cut current line if no selection",
"Cut Append\tAlt+X, Cut and append selected text to clipboard, or cut and append current line if no selection",
"Paste\tControl+V, Paste text from clipboard",
"Paste File\tControl+Shift+V, Insert another file at cursor position",
"Append from Clipboard\tAlt+7, Toggle a mode in which text copied to the clipboard is also saved to a file",
"Undo\tControl+Z, Undo the last editing action",
"Redo\tControl+Shift+Z, Redo the last action that was undone",
"Save Snippet\tAlt+S, Save all or selected text to a snippet file",
"Invoke Snippet\tAlt+V, Pick snippet file to paste or execute",
"View Snippet\tAlt+Shift+V, Pick snippet file to view or edit",
"Yield with Regular Expression\tControl+Shift+Y, Count parts of text matching a regular expression",
"Extract with Regular Expression\tControl+Shift+E, Extract text matching a regular expression, putting matches in a new window",
"Replace with Regular Expression\tControl+Shift+R, Search and replace regular expression in all or selected text",
"Replace\tControl+R, Search and replace string in all or selected text",
"File Find\tAlt+Shift+F, Open file from list of files containing a search string",
"Forward Find\tControl+F, Search forward for string in all or selected text",
"Reverse Find\tControl+Shift+F, Search backward for string",
"Forward Find with Regular Expression\tControl+F3, Search forward for regular expression in all or selected text",
"Reverse Find with Regular Expression\tControl+Shift+F3, Search forward for regular expression in all or selected text",
"Forward Find at Cursor\tAlt+F3, Search forward for chunk or selected text",
"Reverse Find at Cursor\tAlt+Shift+F3, Search backward for chunk or selected text",
"Forward Find Again\tF3, Search forward for next match",
"Reverse Find Again\tShift+F3, Search backward for previous match",
"Word Wrap\tControl+W, Word wrap lines",
"Unwrap\tControl+Shift+W, Unwrap lines",
"Guard Document\tControl+F7, Make document read-only",
"No Guard\tControl+Shift+F7, Clear read-only status",
"Toggle Punctuation\tJAWSKey+Grave, Accent, Toggle JAWS voice between all and no punctuation",
"Voice Louder\tAlt+Grave, Increase JAWS voice volume by 5%",
"Voice Softer\tAlt+Shift+Grave, Decrease JAWS voice volume by 5%",
"Voice Faster\tControl+Grave, Increase JAWS voice rate by 5%",
"Voice Slower\tControl+Shift+Grave, Decrease JAWS voice rate by 5%",
"Extra Speech Toggle\tControl+Shift+X, Toggle extra speech messages on or off, redirecting to Speech.log file",
"Extra Speech Log\tAlt+Shift+X, Open speech.log file in a new window",
"Go to Percent\tControl+G, Go to percentage point in document",
"Go to Percent Again\tAlt+G, Repeat Go command",
"Jump to Line\tControl+J, Jump to line number or to line, column position",
"Jump to Line Again\tAlt+J, Repeat Jump command",
"Set Bookmark\tControl+K, Set bookmark at cursor position",
"Clear Bookmark\tControl+Shift+K, Clear bookmark at cursor position",
"Go to Bookmark\tAlt+K, Go to bookmark in current file",
"Set Favorite\tControl+L, Add current file to the list of favorites",
"Clear Favorite\tControl+Shift+L, Clear current file from the list of favorites",
"List Favorites\tAlt+L, Open a file from the list of favorites",
"Recent Files\tAlt+R, Open a file from the list of those recently used",
"New\tControl+N, Open a new editing window",
"New from Clipboard\tControl+Shift+N, Open a new editing window containing clipboard text",
"Open\tControl+O, Open file",
"Open Other Format\tControl+Shift+O, Open file in another format and convert it to text",
"Open Again\tAlt+O, Reload the current file from disk",
"Properties\tAlt+Enter, display Windows properties dialog for current file",
"Save\tControl+S, Save",
"Save As\tControl+Shift+S, Save As",
"Save Copy\tAlt+Shift+S, Save copy of document using a different name",
"Export Format\tAlt+Shift+E, Export document to another format",
"Print\tControl+P, Print current file",
"Mail Body\tControl+M, Mail current file as body of an email message",
"Mail Attachment\tControl+Shift+M, Mail current file as an email attachment",
"Burn to CD\tAlt+Shift+B, Burn a list of files or folders to a CD",
"Web Download\tAlt+Shift+W, Pick files to download from a web page or the current document",
"Web Client Utilities\tAlt+Shift+Space, Pick a web client utility to run",
"Run\tF5, Execute current file, based on its extension",
"Run at Cursor\tShift+F5, Execute a web URL or email address at cursor position or in selected text",
"Prompt Command\tAlt+F5, Prompt for a command line to execute and say its standard output",
"Review Output\tAlt+Shift+F5, Open standard output of last prompt or compile command in a new editing window",
"Compile\tControl+F5, Compile source code, say output, and jump to error position",
"Pick Compiler\tControl+Shift+F5, Pick a compiler or interpreter from the list of those configured",
"Say Compiler\tAlt+0, Say current compiler and folder",
"Go to Folder\tControl+0, Go to folder containing recent or favorite files",
"Go to Special Folder\tControl+Shift+0, Go to special folder of Windows",
"Go to Environment\tControl+Shift+G, Go to interactive environment of current compiler",
"Spell Check\tF7, Spell check all or selected text",
"Thesaurus\tShift+F7, Look up synonyms for word at cursor",
"Lookup Term\tAlt+F7, Look up information from dictionary.com, thesaurus.com, and wikipedia.org",
"Translate Language\tAlt+Shift+F7, Translate all or selected text from one natural language to another",
"Say Path\tAlt+P, Say full path of current file",
"Path to Clipboard\tAlt+Shift+P, Copy full path of current file to clipboard",
"Path List\tControl+Shift+P, Generate a list of files in a new editing window",
"Special Character\tF2, Insert character indirectly by specifying its Unicode value",
"Quote\tControl+Q, Add prefix sequence to current or selected lines",
"Unquote\tControl+Shift+Q, Remove prefix sequence from current or selected lines",
"Join Lines\tControl+Shift+J, Word wrap lines in all or selected paragraphs",
"Hard Line Break\tControl+Shift+H, Set the maximum width of lines in all or selected text",
"Upper Case\tControl+U, Convert current or selected characters to upper case",
"Lower Case\tControl+Shift+U, Convert current or selected characters to lower case",
"Proper Case\tAlt+U, Convert current or selected characters to proper case",
"Swap Case\tAlt+Shift+U, Convert lower case characters to upper case, and vice versa",
"Yield Encoding\tAlt+Shift+Y, Render all or selected text based on a character encoding",
"Format Code\tControl+4, Arrange indentation and other stylistic conventions in a C-like language",
"Repeat Line\tControl+Y, Copy current line below it",
"Evaluate Expression\tControl+Equals, Evaluate current line or selected text as a JScript.NET expression and copy the result below",
"Replace Tokens\tControl+Shift+Equals, Swap user-defined tokens with their computed results in all or selected text",
"Transform Files\tAlt+Equals, Apply a set of search and replace tasks to a list of files in the current window",
"Trim Blanks\tControl+Shift+Enter, Trim leading and trailing blanks from the current or selected lines, and remove more than two consecutive blank lines",
"End Character\tAlt+End, Go to last non-blank character of line and read it",
"Home Character\tAlt+Home, Go to first non-blank character of line and read it",
"Next Word\tControl+RightArrow, Go to next word and read it",
"Prior Word\tControl+LeftArrow, Go to previous word and read it",
"Next Chunk\tAlt+RightArrow, Go to next chunk and read it",
"Prior Chunk\tAlt+LeftArrow, Go to previous chunk and read it",
"Next Sentence\tAlt+DownArrow, Go to next sentence and read it",
"Prior Sentence\tAlt+UpArrow, Go to previous sentence and read it",
"Next Paragraph\tControl+DownArrow, Go to next paragraph and read it",
"Prior Paragraph\tControl+UpArrow, Go to previous paragraph and read it",
"Delete Right\tControl+Shift+Delete, Delete from cursor to end of line",
"Delete Left\tControl+Shift+Backspace, Delete from cursor to start of line",
"Delete Down\tAlt+Shift+Delete, Delete from cursor to bottom of file",
"Delete Up\tAlt+Shift+Backspace, Delete from cursor to top of file",
"Delete Line\tAlt+Backspace, Delete current line",
"Delete Hard Line\tControl+D, Delete line ending in hard line break",
"Delete Paragraph\tControl+Shift+D, Delete past one or more blank lines",
"Delete File\tAlt+Shift+D, Delete current file on disk",
"Rename\tAlt+Shift+R, Rename current file on disk",
"Next Section\tControl+PageDown, Go to next section",
"Prior Section\tControl+PageUp, Go to Prior Section",
"Go to Section\tF6, Go to section in body from topic in table of contents",
"Go to Contents\tShift+F6, Go to topic in table of contents from section in body",
"Search for Topic\tControl+F6, Search for a topic based on text in its heading",
"Search for Topic Again\tAlt+F6, Search again for the next matching topic",
"Topic\tAlt+T, Say topic of current section",
"Text Contents\tAlt+Shift+T, Generate and prepend a table of contents to the current document",
"Section Break\tControl+Enter, Insert a section break at the cursor position",
"HTML Format\tControl+H, Convert current document to HTML in a new window",
"Text Convert\tControl+T, Convert other formats to text files with the same name except for a .txt extension",
"Text Combine\tControl+Shift+T, Convert other formats to text and combine them in a new editing window",
"Justify\tAlt+Shift+J, Set justification of cursor or selected text",
"Style\tAlt+Shift+Slash, Set style of cursor or selected text",
"Baseline\tAlt+Shift+6, Set vertical alignment of cursor or selected text",
"Set Selection Font\tAlt+Shift+Dash, Set font of cursor or selected text",
"Next Alignment\tControl+RightBracket, Go to next change in justification",
"Prior Alignment\tControl+LeftBracket, Go to previous change in justification",
"Next Style\tControl+Slash, Go to next change in style",
"Prior Style\tControl+Shift+Slash, Go to previous change in style",
"Next Baseline\tControl+6, Go to next change in baseline",
"Prior Baseline\tControl+Shift+6, Go to previous change in baseline",
"Next Font\tControl+Dash, Go to next change in font",
"Prior Font\tControl+Shift+Dash, Go to previous change in font",
"Say Font\tAlt+Dash, Say current font and color",
"Say Styles\tAlt+Slash, Say current justification and styles",
"Infer Indent\tAlt+RightBracket, Infer the indent unit of the current document, or configure EdSharp accordingly if repeated",
"Toggle Indentation\tWindows+Grave, Toggle announcement of indentation by JAWS",
"Indent Mode\tAlt+Shift+I, Toggle auto indent with Enter, and announcement of indentation changes",
"Enter New Line\tEnter, Start new line at left margin",
"Indent New Line\tShift+Enter, Start new line with same indentation as current one",
"Indent New Line Prior\tAlt+Shift+Enter, insert prior line with same indentation as current one",
"Indent\tTab, Indent current line or selected text by one unit",
"Outdent\tShift+Tab, Reduce indentation of current or selected lines by one unit",
"Align\tAlt+Shift+A, Adjust indentation of current or selected lines according to prior line",
"Next Block\tControl+B, Go to the next block of code, having the same or less indentation",
"Prior Block\tControl+Shift+B, Go to the previous block of code, having the same or less indentation",
"Next Indent\tControl+I, Go to the next change in indentation",
"Prior Indent\tControl+Shift+I, Go to the previous change in indentation",
"Right Brace\tControl+Shift+RightBracket, Search forward for matching right brace character",
"Left Brace\tControl+Shift+LeftBracket, Search backward for matching left brace character",
"End Tag\tControl+Shift+Period, go to closing tag of HTML element",
"Start Tag\tControl+Shift+Comma, Go to opening tag of HTML element",
"Next Part\tAlt+PageDown, Go to next match of NavigatePart setting",
"Prior Part\tAlt+PageUp, Go to previous match of NavigatePart setting",
"Go to Part\tAlt+Shift+G, Pick a part to go to",
"Order Items\tAlt+Shift+O, Sort items alphabetically in all or selected text",
"Reverse Items\tAlt+Shift+Z, Reverse order of all or selected items of text",
"Keep Unique Items\tAlt+Shift+K, Discard repetitive items in all or selected text",
"Number Items\tAlt+Shift+N, Insert numbers at the start of items in all or selected text",
"List Different Items\tAlt+Shift+L, Compare two lists and put non-overlapping items in a new window",
"Query Common Items\tAlt+Shift+Q, Compare two lists and put overlapping items in a new window",
"PyDent\tAlt+LeftBracket, Convert from PyBrace format, or reformat typical Python code, using the IndentUnit setting and adding comments at ends of blocks",
"PyBrace\tAlt+Shift+LeftBracket, Convert from PyDent format, or reformat typical Python code, using braces instead of indentation and adding comments at ends of blocks",
"Insert Script Path\tControl+I, Insert JAWS script path in Open or Save Dialog",
"Insert All Users Path\tControl+Shift+I, Insert JAWS All Users path in Open or Save Dialog",
"Explorer Folder\tAlt+Backslash, Open Windows Explorer in the EdSharp program folder, data folder, or current folder",
"Command Prompt\tControl+Backslash, Open a command prompt in the EdSharp program folder, data folder, or current folder",
"Environment Variables\tControl+E, Change Windows environment variables for the current process, user, or system",
"Next Window\tControl+Tab, Cycle to next editing window",
"Prior Window\tControl+Shift+Tab, Cycle to previous editing window",
"Windows Open\tShift+F4, Say titles of current editing windows",
"Current Windows\tF4, Activate an editing window from a list of those currently open",
"Close Window\tControl+F4, Close current editing window",
"Close All but Current Window\tControl+Shift+F4, Close all editing windows except the current one",
"Exit EdSharp\tAlt+F4, Exit the EdSharp application",
"Arrange Icons\tAlt+F11, Arrange open windows",
"Cascade\tControl+F11, Cascade open windows",
"Tile Horizontal\tAlt+Shift+F11, Tile open windows horizontally",
"Tile Vertical\tControl+Shift+F11, Tile open windows vertically",
"Elevate Version\tF11, Download latest EdSharp version and run installer (after confirming)",
"Set on Favorite List\tControl+L, Add or remove the current file on the favorites list",
"Markdown to Plain Text\t, Convert the Markdown in this window to plain text",
"HTML to Markdown\t, Convert the HTML in this window to Markdown",
"HTML to Plain Text\t, Convert the HTML in this window to plain text, keeping paragraphs and lists",
"Preview Markdown\tControl+F9, Show this Markdown as a formatted page in a preview window",
"Preview Markdown in Web Browser\t, Show this Markdown as a formatted page in your web browser, where diagrams are drawn",
"Check Markdown\tAlt+F9, Report problems in this Markdown: heading jumps, images without alt text, bare web addresses, unclosed code fences, mismatched table rows, and link references defined but never used",
"Run Code Blocks\tAlt+Shift+F9, Run this document's sql and jscript code blocks and put each block's results below it",
"Chat with AI\tF12, Ask an AI model on this computer a question; the answer opens in a new window, and the document travels with the question when your wording refers to it",
"Chat about Document\tShift+F12, Ask an AI model on this computer about the open text: the selection when text is selected, the whole document when it is not",
"Tutorial\tControl+Shift+F1, Open the EdSharp tutorial in your web browser",
"Sample Programs\tControl+Shift+F2, List the sample programs that ship with EdSharp and open one",
"Copy Log\tControl+F12, Copy this session's log path to the clipboard, as a file for pasting into a mail message and as text"
}; // c_aCommandSummaries

static Dictionary<string, string> dCommandSummaries = null;

// The table as a lookup, built once.
public static string builtInSummary(string sCommand) {
if (dCommandSummaries == null) {
dCommandSummaries = new Dictionary<string, string>();
foreach (string sLine in c_aCommandSummaries) {
int iTab = sLine.IndexOf('\t');
if (iTab <= 0) continue;
string sName = sLine.Substring(0, iTab);
if (!dCommandSummaries.ContainsKey(sName)) dCommandSummaries.Add(sName, sLine.Substring(iTab + 1));
}
}
string sValue;
if (dCommandSummaries.TryGetValue(sCommand, out sValue)) return sValue;
if (dCommandSummaries.TryGetValue("Say " + sCommand, out sValue)) return sValue;
return "";
} // builtInSummary method

public string[] GetKeySummary(ToolStripMenuItem item) {
string sCommand = item.Name;
// KeyMap (Homer) is the single source. The first time a command's summary
// is requested it is read once from Hotkeys.ini and cached here, so the
// status bar, Key Describer, and Alternate Menu all agree and later reads
// cost nothing. The Hotkeys.ini format and parse are unchanged.
string sValue = KeyMap.getSummary(sCommand);
if (sValue.Length == 0) {
// Hotkeys.ini first, so a description someone has edited still wins,
// then the table built into the program, which cannot be missing or
// out of date. Only a command in neither is left undescribed, and the
// audit script makes sure there is no such command.
sValue = Ini.ReadValue(App.HotkeyIniFile, "Hotkeys", sCommand, "");
if (sValue.Length == 0) sValue = Ini.ReadValue(App.HotkeyIniFile, "Hotkeys", "Say " + sCommand, "");
if (sValue.Length == 0) sValue = builtInSummary(sCommand);
if (sValue.Length == 0) sValue = "No description available";
KeyMap.setSummary(sCommand, sValue);
}
string sKey = "";
string sDescription = "";
int iComma = sValue.IndexOf(",");
if (iComma == -1) sDescription = sValue;
else {
sKey = sValue.Substring(0, iComma);
sDescription = sValue.Substring(iComma + 1);
}
return new string[] {sCommand, sKey, sDescription};
} // GetKeySummary method

public void menuItem_Click(object sender, EventArgs e) {
//Util.Beep();
HomerRichTextBox rtb = null;
string[] aLabels, aValues, aResults;
bool bSelected;
int iLength, iStart, iEnd, iResult, iIndex, iLine, iPercent, iCount;
string sFile, sMatch, sReplace, sPattern, sSubstitute, sLine, sTitle, sText, sResult, sLabel, sValue;

ToolStripMenuItem menuItem = (ToolStripMenuItem) sender;
string sOptions = (string) menuItem.Tag;
sOptions = " " + sOptions.Trim().ToLower() + " ";
//sLabel = menuItem.Text.Replace("&", "").Replace(" ...", "").Split('\t')[0];
sLabel = menuItem.Name;
// Exit EdSharp switches the mode OFF and then goes ahead. Turning it off
// matters as much as exiting: closing a document may ask whether to save
// it, and those questions must answer to ordinary keys rather than being
// caught and described. The change of mode is announced as usual, so
// nothing happens silently.
if (this.KeyDescriber && sLabel == "Exit EdSharp") {
this.KeyDescriber = false;
AddMessage("No Key Describer");
}
// Everything else is described instead of done -- except Key Describer
// itself, or the mode could not be switched off.
if (this.KeyDescriber && sLabel != "Key Describer") {
string[] aSummary = GetKeySummary(menuItem);
SetMessage(aSummary[0]);
AddMessage(aSummary[1]);
AddMessage(aSummary[2]);
return;
}

if (sOptions.Contains(" silent ")) SetStatus(sLabel);
else SetMessage(sLabel);

MdiChild child = this.Child;
if (child == null) {
if (!sOptions.Contains(" frame ")) return;
}
else {
rtb = Child.RTB;
}

if (menuItem == menuFileNew) {
child = new MdiChild(App.Frame);
}

if (menuItem == menuFileNewFromClipboard) {
new MdiChild(App.Frame);
child = App.Frame.Child;
Child.RTB.Text = Util.GetClipboardText();
child.RTB.Modified = true;
}

if (menuItem == menuFileOpen) {

string sDir = "";
/*
if (child != null) sDir = Path.GetDirectoryName(child.File);
*/
sFile = Dialog.OpenFile("", sDir);
if (sFile.Length == 0) return;

int iConvert = 0;
if (Path.GetExtension(sFile).ToLower() == ".rtf") {
switch (Dialog.Confirm("Confirm", "Treat as rich text?", "Y")) {
case "Y" :
iConvert = 2;
break;
case "" :
return;
}
}
OpenOrActivateWindow(sFile, iConvert);
}

if (menuItem == menuFileOpenOtherFormat) {
sFile = Dialog.OpenFile("", "");
if (sFile.Length == 0) return;

OpenOrActivateWindow(sFile, 2);
return;
/*
AddMessage("Converting");
try {
sText = COM.ConvertFile2String(sFile);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}

if (!IsEmptyWindow()) new MdiChild(this);
child = this.Child;
rtb = child.RTB;
rtb.Text = sText;
rtb.Index = 0;
child.Text = Path.GetFileNameWithoutExtension(sFile) + ".txt";
//rtb.Modified = true;
rtb.Modified = false;
//AddMessage("Done");
*/
}

if (menuItem == menuFileOpenAgain) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

iIndex = rtb.Index;
if (Dialog.Confirm("Confirm", "Existing text will be  replaced.  Sure?", "Y") != "Y") return;
child.LoadTextOrRtfFile(sFile);
rtb.Index = iIndex;
}

if (menuItem == menuFileRecent) {
aResults = App.ReadSectionKeys("Recent");
List<string> list = new List<string>(aResults);
for (int i = list.Count - 1; i >=0; i--) {
string s = list[i];
if (File.Exists(s)) continue;
App.DeleteKey("Recent", s);
list.RemoveAt(i);
}

aResults = list.ToArray();
if (aResults.Length == 0) {
AddMessage("No items!");
return;
}

string[] aTime = new string[aResults.Length];
for (int i = 0; i < aTime.Length; i++) aTime[i] = App.ReadValue("Recent", aResults[i], "");
Array.Sort(aTime, aResults);
//Array.Reverse(aTime);
Array.Reverse(aResults);

iLength = aResults.Length;
int iMax = Int32.Parse(App.ReadOption("RecentFiles", "30"));
if (iLength > iMax) {
List<string> listResults = new List<string>(aResults);
for (int i = iLength - 1; i >= iMax; i--) {
App.DeleteKey("Recent", listResults[i]);
listResults.RemoveAt(i);
}
aResults = listResults.ToArray();
}

string[] aDisplay = new string[aResults.Length];
for (int i = 0; i < aDisplay.Length; i++) aDisplay[i] = Path.GetFileName(aResults[i]);

sFile = Dialog.Pick("Recent Files", aResults, aDisplay, false, 0);
if (sFile.Length == 0) return;

OpenOrActivateWindow(sFile, GetViewLevel(sFile));
}

if (menuItem == menuFileFind) {
FileFind();
}

if (menuItem == menuFileSaveCopy) {
AddMessage("Save Copy");
sFile = child.File + ".bak";
sFile = Dialog.SaveFile("", sFile);
if (sFile.Length == 0) return;
if (sFile.ToLower().IndexOf(".rtf") >= 0) rtb.SaveFile(sFile, RichTextBoxStreamType.RichText);
//else if (Util.IsUnicode(rtb.Text)) Util.String2File(rtb.Text, sFile);
//else rtb.SaveFile(sFile, RichTextBoxStreamType.PlainText);
else Util.String2File(Util.Convert2WinLineBreak(rtb.Text), sFile);
this.SetRecent(sFile);
}

if ((menuItem == menuFileSave) || (menuItem == menuFileSaveAs)) {
sFile = child.File;
if ((menuItem == menuFileSave) && sFile.Contains(@"\")) sText = "";//AddMessage("Save");
else {
sFile = Dialog.SaveFile("", sFile);
if (sFile.Length == 0) return;
}

//Dialog.Show(sFile);
//if (Path.GetExtension(sFile).Length == 0) sFile += ".txt";
child.SaveTextOrRtfFile(sFile);
//this.SetRecent(sFile);
rtb.Modified = false;
}

if (menuItem == menuFileExport) {
sFile = child.File;
aValues = Ini.ReadSectionKeys(App.IniFile, "Export");
//aValues = Array.FindAll(aValues, delegate(string s) {return s.StartsWith("pdf2");} );
//aValues = aValues.ConvertAll(delegate(string s){s.ToLower();});
HomerList hl = new HomerList(aValues);
hl.ToLower();
//list = list.ConvertAll<string>(delegate(string s) { return s.ToLower(); });
string sExt = Path.GetExtension(child.File).ToLower().TrimStart('.');
sMatch = @"^\w+2\w+$";
HomerList hl2 = hl.FindLike(sMatch);
hl.RemoveLike(sMatch);
sMatch = "^" + sExt + @"2\w+";
hl2 = hl2.FindLike(sMatch);
hl.AddRange(hl2);
hl.AddRange("asc|doc|htm|mac|rtf|unx|xml");
// do not offer original format, since already available with Control+O
if (sExt.Length > 0 && !hl.Contains(sExt)) hl.Add(sExt + "2" + sExt);
// hl.Add("Other");
hl.KeepUnique();
aValues = hl.ToArray();
hl.ReplaceLike(@"^\w+2(\w+)$", "$1");
string[] aDisplay = hl.ToArray();
Array.Sort(aDisplay, aValues);

hl.Clear();
hl.AddRange(aDisplay);
hl.Add("Other");
aDisplay = hl.ToArray();
hl.Clear();
hl.AddRange(aValues);
hl.Add("Other");
aValues = hl.ToArray();

sTitle = "Export " + sExt + " to ";
sResult = Dialog.Pick(sTitle, aValues, aDisplay, false, 0);
//Dialog.Show(sResult);
if (sResult.Length == 0) return;

int iCodePage = -1;
string sCodePage = "";
if (sResult == "Other") {
iCodePage = Dialog.PickEncoding("", 0);
if (iCodePage == -1) return;
sCodePage = iCodePage.ToString();
sResult = sCodePage;
}

string sTargetExt = Util.RegExpReplaceCase(sResult, @"^\w+2", "");
sFile = Path.ChangeExtension(sFile, sTargetExt);
sFile = Dialog.SaveFile("", sFile);
if (sFile.Length == 0) return;

if (sCodePage.Length > 0) sExt = "Other";
else sExt = sResult.ToLower();
switch (sExt) {
case "Other" :
Encoding en = Encoding.GetEncoding(iCodePage);
sText = rtb.Text;
sText = Util.Convert2WinLineBreak(sText);
File.WriteAllText(sFile, sText, en);
break;
case "asc" :
case "mac" :
case "unx" :
sText = rtb.Text;
if (sExt == "asc") sText = Util.Convert2Ascii(sText);
else if (sExt == "mac") sText = Util.Convert2MacLineBreak(sText);
else if (sExt == "unx") sText = Util.Convert2UnixLineBreak(sText);
// Util.String2File(sText, sFile);
File.WriteAllText(sFile, sText, Encoding.Default);
break;
case "rtf" :
rtb.SaveFile(sFile, RichTextBoxStreamType.RichText);
break;
default :
string s = Path.GetExtension(child.File);
if (Util.Equiv(s, ".rtf")) sText = rtb.Rtf;
else sText = rtb.Text;
Util.ConvertString2FileFormat(sText, s, sFile, sExt);
break;
}
if (File.Exists(sFile)) AddMessage("Done");
else AddMessage("Error!");
}

if (menuItem == menuFileRename) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

string sName = Dialog.Input("Rename", "File Name", child.Text, "Rename").Trim();
if (sName.Length == 0) return;

string sNewFile = Path.Combine(Path.GetDirectoryName(sFile), sName);
try {
File.Move(sFile, sNewFile);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}
child.Text = sName;
child.File = sNewFile;
SetRecent(sNewFile);
}

if (menuItem == menuFileMailBody) {
string sSubject = Path.GetFileNameWithoutExtension(child.Text);
string sBody = rtb.Text;
// string sRecipient = "";
try {
MapiMail.SendMail(sSubject, sBody, null, null);
}
catch {
Util.MailMessage("", sSubject, sBody);
}
return;

/*
try {
sText = rtb.Text;
sText = Util.Convert2WinLineBreak(sText);
sText = Util.RegExpReplaceCase(sText, "\r\n", "%0D%0A");
sText = Util.RegExpReplaceCase(sText, " ", "%20");
sText = Util.RegExpReplaceCase(sText, "\t", "%09");
sText = Util.RegExpReplaceCase(sText, "\"", "%22");
sText = Util.RegExpReplaceCase(sText, "'", "%27");
sText = Util.RegExpReplaceCase(sText, "\\\\", "%5C");
string sCommand = "mailto:?BODY=" + sText;
Process.Start(sCommand);
}
catch {
Mail(false);
}
*/
}

if (menuItem == menuFileMailAttach) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

string sSubject = Path.GetFileName(child.Text) + " attached";
KeyValuePair<string, string>[] aAttachments = {new KeyValuePair<String, String>(Path.GetFileName(sFile), sFile)};
try {
MapiMail.SendMail(sSubject, "", null, aAttachments);
}
catch {
}
return;

//Mail(true);
}

if (menuItem == menuFileRun) {
sFile = child.File;
if (!sFile.Contains(@"\") || rtb.Modified) {
sFile = Path.Combine(Path.GetTempPath(), Path.GetFileName(sFile));
sText = rtb.Text;
Util.String2File(sText, sFile);
}

Process.Start(sFile);
}

if (menuItem == menuFilePrint) {
sFile = child.File;
sFile = Path.GetFileName(sFile);
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

if (Dialog.Confirm("Confirm", "Print " + sFile + "?", "Y") != "Y") return;

if (Path.GetExtension(sFile).ToLower() == ".rtf") COM.InvokeVerb(sFile, "Print");
else Util.RunHideWait("Notepad.exe /p " + Util.Quote(sFile));
/*
sFile = Path.Combine(Path.GetTempPath(), sFile);
sText = rtb.Text;
sText = Util.Convert2WinLineBreak(sText);
Util.String2File(sText, sFile);
string sExe;
if (Path.GetExtension(sFile).ToLower() == ".rtf") sExe = "cmd.exe /c WordPad.exe";
else sExe = "Notepad.exe";
string sCommand = sExe + " /P " + Util.Quote(sFile);
Util.RunHideWait(sCommand);
File.Delete(sFile);
*/
}

if (menuItem == menuFileProperties) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

Dialog.Properties(sFile);
}

if (menuItem == menuFileCurrentWindows) {
CurrentWindows();
}

if (menuItem == menuFileClose) {
CloseWindow(child);
}

if (menuItem == menuFileCloseAllButCurrentWindow) {
CloseAllButCurrentWindow();
}

if (menuItem == menuFileExit) {
ExitApp();
}

if (menuItem == menuEditSelectAll) {
rtb.SelectAll();
iCount = rtb.SelectedText.Length;
AddMessage(Util.Pluralize(iCount, "character"));
}

if (menuItem == menuEditUnselectAll) {
rtb.DeselectAll();
iIndex = rtb.Index;
if (iIndex >= rtb.TextLength) return;
sText = rtb.GetRange(rtb.Index, rtb.Index + 1);
AddMessage(sText);
}

if (menuItem == menuEditCopy) {
if (rtb.SelectionLength == 0) {
AddMessage("Line");
sText = rtb.RowText + LineBreak;
}
else {
AddMessage("Selected");
sText = rtb.SelectedText;
rtb.StoreSelection();
}
sText = Util.Convert2WinLineBreak(sText);
Util.SetClipboardText(sText);
}

if (menuItem == menuEditCopyAppend) {
sText = Util.GetClipboardText();
sText = Util.Convert2UnixLineBreak(sText);
if (sText.Length > 0 && !sText.EndsWith(LB)) sText += LB;
if (rtb.SelectionLength == 0) {
AddMessage("Line");
sText += rtb.RowText + LB;
}
else {
AddMessage("Selected");
sText += rtb.SelectedText;
rtb.StoreSelection();
}
sText = Util.Convert2WinLineBreak(sText);
Util.SetClipboardText(sText);
}

if (menuItem == menuEditCopyRichText) {
rtb.Copy();
}

if (menuItem == menuEditCut) {
if (rtb.SelectionLength == 0) {
AddMessage("Line");
sText = rtb.RowText + LineBreak;
rtb.Select(rtb.RowStart, rtb.RowLength);
}
else {
AddMessage("Selected");
sText = rtb.SelectedText;
}
rtb.Cut();
sText = Util.Convert2WinLineBreak(sText);
Util.SetClipboardText(sText);
Util.Say(rtb.RowText);
}

if (menuItem == menuEditCutAppend) {
sText = Util.GetClipboardText();
sText = Util.Convert2UnixLineBreak(sText);
if (sText.Length > 0 && !sText.EndsWith(LB)) sText += LB;
if (rtb.SelectionLength == 0) {
AddMessage("Line");
sText += rtb.RowText + LB;
rtb.Select(rtb.RowStart, rtb.RowLength);
}
else {
AddMessage("Selected");
sText += rtb.SelectedText;
}
rtb.Cut();
sText = Util.Convert2WinLineBreak(sText);
Util.SetClipboardText(sText);
Util.Say(rtb.RowText);
}

if (menuItem == menuEditPaste) {
rtb.Paste();
sText = Util.GetClipboardText();
sText = Util.Convert2UnixLineBreak(sText);
aResults = Util.RegExpExtractCase(sText, @"\s+\Z");
if (aResults.Length > 0) {
iIndex = rtb.Index;
sText = aResults[0];
rtb.ReplaceRange(iIndex, iIndex, sText);
rtb.Index = iIndex + sText.Length;
}
}

if (menuItem == menuEditPasteFile) {
sFile = Dialog.OpenFile("", "");
if (sFile.Length == 0) return;

sText = Util.File2String(sFile);
string sChoice = "N";
if (Path.GetExtension(sFile).ToLower() == ".rtf") sChoice = Dialog.Confirm("Confirm", "Treat as rich text?", "Y");
if (sChoice.Length == 0) return;

rtb.Index = rtb.SelectionStart + rtb.SelectionLength;
if (sChoice == "Y") rtb.SelectedRtf = sText;
else rtb.SelectedText = sText;
Util.Say(rtb.RowText);
}

if (menuItem == menuEditUndo) {
rtb.Undo();
}

if (menuItem == menuEditRedo) {
rtb.Redo();
}

if (menuItem == menuEditStartSelection) {
//if (!IsCharacter(out iIndex)) return;
iIndex = rtb.Index;
rtb.StartSelection = iIndex;
if (iIndex >=0 && iIndex < rtb.TextLength) {
sText = rtb.GetRange(rtb.Index, rtb.Index + 1);
AddMessage(sText);
}
}

if (menuItem == menuEditCompleteSelection) {
iStart = rtb.StartSelection;
iEnd = rtb.Index;
if (iStart > iEnd) Util.Swap(ref iStart, ref iEnd);
rtb.SelectRange(iStart, iEnd);
iCount = rtb.SelectedText.Length;
AddMessage(Util.Pluralize(iCount, "character"));
rtb.OldSelectionStart = rtb.SelectionStart;
rtb.OldSelectionLength = rtb.SelectionLength;
}

if (menuItem == menuEditReselect) {
rtb.Reselect();
}

if (menuItem == menuEditCopyAll) {
sText = rtb.Text;
sText = Util.Convert2WinLineBreak(sText);
Util.SetClipboardText(sText);
}

if (menuItem == menuEditSelectChunk) {
bool bLoop = false;
string c = "";
object[] a = GetChunk();
iStart = (int) a[0];
iIndex = iStart;
sText = rtb.Text;
if (rtb.SelectionLength == 0) {
AddMessage("Select Chunk");
}
else {
iStart = rtb.SelectionStart;
bLoop = iIndex < sText.Length;
while (bLoop) {
c = sText.Substring(iIndex, 1);
bLoop = (c.Trim().Length == 0);
iIndex++;
bLoop = (bLoop && iIndex < sText.Length);
}
}

bLoop = iIndex < sText.Length;
while (bLoop) {
c = sText.Substring(iIndex, 1);
bLoop = (c.Trim().Length > 0);
iIndex++;
bLoop = (bLoop && iIndex < sText.Length);
}
iEnd = iIndex;
rtb.SelectRange(iStart, iEnd);
}

if (menuItem == menuEditAppendFromClipboard) {
if (child.AppendFromClipboard == 0) {
AddMessage("Append from Clipboard On");
child.AppendFromClipboard = -1;
child.NextClipboardViewer =  Util.SetClipboardViewer(child.Handle);
}
else {
AddMessage("No Append from Clipboard");
child.AppendFromClipboard = 0;
Util.ChangeClipboardChain(child.Handle, child.NextClipboardViewer);
child.NextClipboardViewer = (IntPtr) 0;
}
}

if (menuItem == menuEditQuote) {
//if (!IsCharacter(out iIndex)) return;

string sPrefix = App.ReadOption("QuotePrefix", "> ");
if (rtb.SelectionLength == 0) {
AddMessage("Line");
iStart = rtb.RowStart;
iEnd = iStart + rtb.RowText.Length;
//Dialog.Show(iStart, iEnd);
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sText = rtb.GetRange(iStart, iEnd);
sText = Util.RegExpReplaceCase(sText, "^", sPrefix);
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
Util.Say(rtb.RowText);
}

if (menuItem == menuEditUnquote) {
string sPrefix = App.ReadOption("QuotePrefix", "> ");
if (rtb.SelectionLength == 0) {
AddMessage("Line");
iStart = rtb.RowStart;
iEnd = iStart + rtb.RowText.Length;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sText = rtb.GetRange(iStart, iEnd);
//sText = Util.RegExpReplaceCase(sText, @"^( |\t|\>)+", "");
sText = Util.RegExpReplaceCase(sText, @"^" + sPrefix, "");
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
Util.Say(rtb.RowText);
}

if (menuItem == menuEditUpperCase) {
if (rtb.SelectionLength == 0) {
AddMessage("Character");
iStart = rtb.Index;
iEnd = iStart + 1;
bSelected = false;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
bSelected = true;
}

sText = rtb.GetRange(iStart, iEnd);
sText = sText.ToUpper();
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
if (bSelected) sText = rtb.RowText;
else sText = rtb.GetRange(rtb.Index, rtb.Index + 1);
AddMessage(sText);
}

if (menuItem == menuEditLowerCase) {
if (rtb.SelectionLength == 0) {
AddMessage("Character");
iStart = rtb.Index;
iEnd = iStart + 1;
bSelected = false;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
bSelected = true;
}

sText = rtb.GetRange(iStart, iEnd);
sText = sText.ToLower();
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
if (bSelected) sText = rtb.RowText;
else sText = rtb.GetRange(rtb.Index, rtb.Index + 1);
AddMessage(sText);
}

if (menuItem == menuEditProperCase) {
if (rtb.SelectionLength == 0) {
AddMessage("Character");
iStart = rtb.Index;
iEnd = iStart + 1;
bSelected = false;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
bSelected = true;
}

sText = rtb.GetRange(iStart, iEnd);
sText = Util.ProperCase(sText);
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
if (bSelected) sText = rtb.RowText;
else sText = rtb.GetRange(rtb.Index, rtb.Index + 1);
AddMessage(sText);
}

if (menuItem == menuEditSwapCase) {
if (rtb.SelectionLength == 0) {
AddMessage("Character");
iStart = rtb.Index;
iEnd = iStart + 1;
bSelected = false;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
bSelected = true;
}

sText = rtb.GetRange(iStart, iEnd);
sText = Util.SwapCase(sText);
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
if (bSelected) sText = rtb.RowText;
else sText = rtb.GetRange(rtb.Index, rtb.Index + 1);
AddMessage(sText);
}

if (menuItem == menuEditYieldEncoding) {
if (rtb.SelectionLength == 0) {
sTitle = "Yield Encoding All";
iStart = 0;
iEnd = rtb.TextLength;
}
else {
sTitle = "Yield Encoding Selected";
iStart = rtb.SelectionStart;
// Dialog.Show(iStart, rtb.SelectionLength);
iEnd = iStart + rtb.SelectionLength;
}

sText = rtb.GetRange(iStart, iEnd);

// string[] aButtons = {"&Default", "&ASCII", "UTF-&7", "UTF-&8", "&Unicode", "UTF-&32", "&Latin1", "&Other"};
string[] aButtons = {"&ASCII", "&Latin1", "UTF-&8", "&UTF-16", "UTF-&7", "UTF-&32", "&Other", "&Codes"};
string sButton = Dialog.Choose(sTitle, "", aButtons, 0);
if (sButton.Length == 0) return;

Encoding def = Encoding.Default;
byte[] aBytes = new byte[def.GetByteCount(sText)];
def.GetBytes(sText, 0, sText.Length, aBytes, 0);
switch (sButton) {
case "&Default" :
sText = def.GetString(aBytes);
break;
case "&ASCII" :
Encoding asc = Encoding.GetEncoding("us-ascii", new EncoderReplacementFallback(""), new DecoderReplacementFallback(""));
// aBytes = new byte[asc.GetByteCount(sText)];
// asc.GetBytes(sText, 0, sText.Length, aBytes, 0);
sText = asc.GetString(aBytes);
break;
case "UTF-&7" :
sText = Encoding.UTF7.GetString(aBytes);
break;
case "UTF-&8" :
sText = Encoding.UTF8.GetString(aBytes);
break;
case "&UTF-16" :
sText = Encoding.Unicode.GetString(aBytes);
break;
case "UTF-&32" :
sText = Encoding.UTF32.GetString(aBytes);
break;
case "&Latin1" :
Encoding latin1 = Encoding.GetEncoding(1252, new EncoderReplacementFallback(""), new DecoderReplacementFallback(""));
// aBytes = new byte[latin1.GetByteCount(sText)];
// latin1.GetBytes(sText, 0, sText.Length, aBytes, 0);
sText = latin1.GetString(aBytes);
break;
case "&Codes" :
sResult = "\n";
foreach (char c in sText) {
sResult+= ((int) c).ToString() + "\n";
}
sText = sResult;
break;

case "&Other" :
/*
string sCodePage = Dialog.Input("Input", "Code Page:", "", "CodePage");
if (sCodePage.Length == 0) return;
Encoding other;
try {
if (Util.IsNumeric(sCodePage)) other = Encoding.GetEncoding(Int32.Parse(sCodePage), new EncoderReplacementFallback(""), new DecoderReplacementFallback(""));
else other = Encoding.GetEncoding(sCodePage, new EncoderReplacementFallback(""), new DecoderReplacementFallback(""));
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}
*/

int iCodePage = Dialog.PickEncoding("", 0);
if (iCodePage == -1) return;

Encoding other = Encoding.GetEncoding(iCodePage, new EncoderReplacementFallback(""), new DecoderReplacementFallback(""));

// aBytes = new byte[other.GetByteCount(sText)];
// other.GetBytes(sText, 0, sText.Length, aBytes, 0);
sText = other.GetString(aBytes);
break;
}

rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart + sText.Length;
Util.Say(rtb.RowText);
}

if (menuItem == menuEditJoinLines) {
if (rtb.SelectionLength == 0) {
AddMessage("All");
iStart = 0;
iEnd = rtb.TextLength;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sText = rtb.GetRange(iStart, iEnd);
sText = Util.RegExpReplaceCase(sText, @" +\n", "\n");
sText = Util.RegExpReplaceCase(sText, "([^\n])\n([^\n])", "$1 $2");
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart + sText.Length;
Util.Say(rtb.RowText);
}

if (menuItem == menuEditHardLineBreak) {
HardLineBreak();
}

if (menuItem == menuEditEnterNewLine|| menuItem == menuEditIndentNewLine) {
SetStatus("");
if ((!rtb.IndentMode && menuItem == menuEditEnterNewLine) || (rtb.IndentMode && menuItem == menuEditIndentNewLine)) {
// Reduce verbosity
// if (rtb.IndentMode) AddMessage("Enter New Line");
// else SetStatus("Enter New Line");
SetStatus("Indent New Line");
sText = "\n";
iIndex = rtb.Index;
rtb.ReplaceRange(iIndex, iIndex, sText);
rtb.Index = iIndex + sText.Length;
}
else {
SetStatus("Indent New Line");
// Reduce verbosity
// AddMessage("Indent New Line");
sText = rtb.RowText;
iIndex = rtb.RowStart + sText.Length;
sMatch = @"^(\s*).*";
sReplace = "$1";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
sText = "\n" + sText;
rtb.ReplaceRange(iIndex, iIndex, sText);
rtb.Index = iIndex + sText.Length;
AddMessage("Level " + this.GetIndent());
// Reduce verbosity
// Util.Say(rtb.RowText);
}
}

if (menuItem == menuEditIndentNewLinePrior) {
sText = rtb.RowText;
iIndex = rtb.RowStart;
sMatch = @"^(\s*).*";
sReplace = "$1";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
sText = sText + "\n";
rtb.ReplaceRange(iIndex, iIndex, sText);
rtb.Index--;
AddMessage("Level " + this.GetIndent());
// Reduce verbosity
// Util.Say(rtb.RowText);
}

if (menuItem == menuEditIndent) {
string sIndent = GetIndentUnit();
iIndex = rtb.Index;
bool bLine;
if (rtb.SelectionLength == 0) {
// AddMessage("Line");
iStart = rtb.RowStart;
iEnd = iStart + rtb.RowText.Length;
bLine = true;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
bLine = false;
}

sText = rtb.GetRange(iStart, iEnd);
sText = Util.RegExpReplaceCase(sText, "^", sIndent);
rtb.ReplaceRange(iStart, iEnd, sText);

//if (bLine) rtb.Index = iIndex;
if (bLine) rtb.Index = iIndex + sIndent.Length;
else rtb.Index  = iIndex + sText.Length;
AddMessage("Level " + this.GetIndent());
//Util.Say(rtb.RowText);
}

if (menuItem == menuEditOutdent) {
string sIndent = GetIndentUnit();
iIndex = rtb.Index;
if (rtb.SelectionLength == 0) {
// AddMessage("Line");
iStart = rtb.RowStart;
iEnd = iStart + rtb.RowText.Length;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sText = rtb.GetRange(iStart, iEnd);
iLength = sText.Length;
sText = Util.RegExpReplaceCase(sText, "^" + sIndent, "");
rtb.ReplaceRange(iStart, iEnd, sText);
if (sText.Length < iLength) rtb.Index = iIndex - sIndent.Length;
AddMessage("Level " + this.GetIndent());
//Util.Say(rtb.RowText);
}

if (menuItem == menuEditAlign) {
string sIndent = GetIndentUnit();
iIndex = rtb.Index;
bool bLine;
if (rtb.SelectionLength == 0) {
AddMessage("Line");
iStart = rtb.RowStart;
iEnd = iStart + rtb.RowText.Length;
bLine = true;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
bLine = false;
}

char[] a = {' ', '\t'};
sLine = "";
string sComment = App.ReadOption("QuotePrefix", "> ");
int iRow = rtb.GetIndexRow(iStart);
// Dialog.Show("iStart " + iStart, "iRow " + iRow);
int iLevels = GetIndent(iRow);
int i = iLevels;
int iTop = 0;
while (iRow > iTop) {
iRow--;
// Dialog.Show("row " + iRow);
sLine = rtb.GetRowText(iRow).Trim(a);
if (sLine.Length == 0 || sLine.StartsWith(sComment)) continue;
i = GetIndent(iRow);
if (iLevels != i) break;
}

int iDelta = i - iLevels;
// Dialog.Show(iDelta);

sText = rtb.GetRange(iStart, iEnd);
if (iDelta > 0) {
for (i = 0; i < iDelta; i++) {
iLength = sText.Length;
sText = Util.RegExpReplaceCase(sText, "^", sIndent);
if (sText.Length > iLength) iIndex += sIndent.Length;
}
}
else {
iDelta = iDelta * -1;
for (i = 0; i < iDelta; i++) {
iLength = sText.Length;
sText = Util.RegExpReplaceCase(sText, "^" + sIndent, "");
if (sText.Length < iLength) iIndex -= sIndent.Length;
}
}

rtb.ReplaceRange(iStart, iEnd, sText);
bLine = !bLine;
rtb.Index = iIndex;
AddMessage("Level " + this.GetIndent());
}

if (menuItem == menuEditIndentMode) {
rtb.IndentMode = !rtb.IndentMode;
AddMessage(rtb.IndentMode ? "On" : "Off");
// The same share-collision guard as in CheckFileTime: the flag file
// may be open in a screen reader script when the mode toggles.
try {
bool b = System.IO.File.Exists(App.IndentModeFile);
if (b && !rtb.IndentMode) System.IO.File.Delete(App.IndentModeFile);
else if (!b && rtb.IndentMode) System.IO.File.Create(App.IndentModeFile).Close();
}
catch (Exception) {}
//return;
}

if (menuItem == menuEditJustify) {
if (rtb.SelectionLength == 0) {
sTitle = "Justify Cursor";
}
else {
sTitle = "Justify Selected";
}

aValues = new string[] {"&Left", "&Bullet", "&Center", "&Right"};
int i = 0;
if (rtb.SelectionBullet) i = 1;
if (rtb.SelectionAlignment == HorizontalAlignment.Center) i = 2;
else if (rtb.SelectionAlignment == HorizontalAlignment.Right) i = 3;
sResult = Dialog.Choose(sTitle, "", aValues, i);
if (sResult.Length == 0) return;

rtb.SelectionBullet = false;
switch (sResult) {
case "&Left" :
rtb.SelectionAlignment = HorizontalAlignment.Left;
break;
case "&Bullet" :
rtb.SelectionBullet = true;
break;
case "&Center" :
rtb.SelectionAlignment = HorizontalAlignment.Center;
break;
case "&Right" :
rtb.SelectionAlignment = HorizontalAlignment.Right;
break;
}
}

if (menuItem == menuEditStyle) {
if (rtb.SelectionLength == 0) {
sTitle = "Style Cursor";
}
else {
sTitle = "Style Selected";
}

aValues = new string[] {"Bold", "Italic", "Underline"};
List<int> listSelect = new List<int>();
if (rtb.SelectionFont.Bold) listSelect.Add(0);
if (rtb.SelectionFont.Italic) listSelect.Add(1);
if (rtb.SelectionFont.Underline) listSelect.Add(2);
int[] aSelect = listSelect.ToArray();

//aResults = Dialog.MultiPick(sTitle, aValues, aSelect, false);
aResults = Dialog.MultiCheck(sTitle, aValues, aSelect, false, 0);
if (aResults.Length == 0) return;

if (!listSelect.Contains(0) && Array.IndexOf(aResults, "Bold") >= 0) rtb.SelectionFont = Util.SetBold(rtb.SelectionFont, true);
if (listSelect.Contains(0) && Array.IndexOf(aResults, "Bold") < 0) rtb.SelectionFont = Util.SetBold(rtb.SelectionFont, false);
if (!listSelect.Contains(0) && Array.IndexOf(aResults, "Italic") >= 0) rtb.SelectionFont = Util.SetItalic(rtb.SelectionFont, true);
if (listSelect.Contains(0) && Array.IndexOf(aResults, "Italic") < 0) rtb.SelectionFont = Util.SetItalic(rtb.SelectionFont, false);
if (!listSelect.Contains(0) && Array.IndexOf(aResults, "Underline") >= 0) rtb.SelectionFont = Util.SetUnderline(rtb.SelectionFont, true);
if (listSelect.Contains(0) && Array.IndexOf(aResults, "Underline") < 0) rtb.SelectionFont = Util.SetUnderline(rtb.SelectionFont, false);
}

if (menuItem == menuEditBaseline) {
if (rtb.SelectionLength == 0) {
sTitle = "Baseline Cursor";
}
else {
sTitle = "Baseline Selected";
}

aValues = new string[] {"&Down", "&Flat", "&Up"};
int i = 1;
if (rtb.SelectionCharOffset < 0) i = 0;
else if (rtb.SelectionCharOffset > 0) i = 2;
sResult = Dialog.Choose(sTitle, "", aValues, i);
if (sResult.Length == 0) return;

switch (sResult) {
case "&Down" :
rtb.SelectionCharOffset = -4;
break;
case "&Flat" :
rtb.SelectionCharOffset = 0;
break;
case "&Up" :
rtb.SelectionCharOffset = 4;
break;
}
}

if (menuItem == menuEditSetSelectionFont) {
if (rtb.SelectionLength == 0) {
AddMessage("Cursor");
}
else {
AddMessage("Selected");
}

object[] a = Dialog.GetFont(rtb.SelectionFont, rtb.SelectionColor);
if (a.Length == 0) return;

rtb.SelectionFont = (Font) a[0];
rtb.SelectionColor = (Color) a[1];
}

if (menuItem == menuMiscEnvironmentVariables) {
string sChoice = Dialog.Choose("Target", "", new string[] {"&Process", "&User", "&Machine"}, 0);
if (sChoice.Length == 0) return;

EnvironmentVariableTarget target = EnvironmentVariableTarget.Process;
if (sChoice == "&User") target = EnvironmentVariableTarget.User;
else if (sChoice == "&Machine") target = EnvironmentVariableTarget.Machine;

IDictionary dic = Environment.GetEnvironmentVariables(target);
iCount = dic.Count;
aLabels = new string[iCount];
aValues = new string[iCount];
string[] aKeys = new string[iCount];

int iKey = 0;
foreach (DictionaryEntry de in dic) {
aKeys[iKey] = ((string) de.Key).ToLower();
aLabels[iKey] = "&" + (string) de.Key;
iKey++;
}
Array.Sort(aKeys, aLabels);

iKey = 0;
foreach (DictionaryEntry de in dic) {
aKeys[iKey] = ((string) de.Key).ToLower();
aValues[iKey] = (string) de.Value;
iKey++;
}
Array.Sort(aKeys, aValues);

sTitle = "Variables for ";
if (sChoice == "&Process") sTitle += "Process " + Process.GetCurrentProcess().ProcessName;
else if (sChoice == "&User") sTitle += "User " + Environment.UserName;
else if (sChoice == "&Machine") sTitle += "Machine " + Environment.MachineName;
aResults = Dialog.MultiInput(sTitle, aLabels, aValues);
if (aResults.Length == 0) return;

try {
for (int i = 0; i < iCount; i++) {
if (aResults[i] == aValues[i]) continue;

if (Dialog.Confirm("Confirm", "Change " + aKeys[i] + "?", "Y") != "Y") continue;
Environment.SetEnvironmentVariable(aKeys[i], aResults[i], target);
}
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}
AddMessage("Done");
}

if (menuItem == menuMiscSpellCheck) {
SpellCheck();
}

if (menuItem == menuMiscExtraSpeechToggle) {
SetStatus("");
bool b = !App.ExtraSpeech;
App.ExtraSpeech = true;
AddMessage("Extra Speech");
AddMessage(b ? "On" : "Off");
App.ExtraSpeech = b;
App.WriteOption("E&xtraSpeech", (b ? "Y" : "N"));
}

if (menuItem == menuMiscExtraSpeechLog) {
OpenOrActivateWindow(App.SpeechLog, 0);
}

if (menuItem == menuMiscThesaurus) {
Thesaurus();
}

if (menuItem == menuMiscLookupTerm) {
if (this.Child != null) {
if (rtb.SelectionLength == 0) {
//AddMessage("Chunk");
object[] a = GetChunk();
iStart = (int) a[0];
sText = (string) a[1];
}
else {
//AddMessage("Selected");
iStart = rtb.SelectionStart;
sText = rtb.SelectedText;
iEnd = iStart + sText.Length;
}

sText = sText.TrimEnd();
}
else sText = "";

if (sText.Length == 0) sText = App.ReadData("Term", "");
sResult = Dialog.Input("Lookup", "Term", sText, "Lookup").Trim();
if (sResult.Length == 0) return;

App.WriteData("Term", sResult);
//AddMessage("Please wait");
AddMessage("Connecting");
sText = VB.LookupTerm(sResult);
if (!IsEmptyWindow()) new MdiChild(this);
child = this.Child;
child.Text = sResult + ".txt";
child.File = child.Text;
rtb = child.RTB;
rtb.Text = sText;
}

if (menuItem == menuMiscTranslateLanguage) {
if (rtb.SelectionLength == 0) {
iStart = 0;
iEnd = rtb.TextLength;
}
else {
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}
sText = rtb.GetRange(iStart, iEnd);

string[] aLanguageNames, aLanguageAbbreviations;
Util.GetGoogleLanguages(out aLanguageNames, out aLanguageAbbreviations);
string sSourceLanguage = Dialog.Pick("Source Language", aLanguageAbbreviations, aLanguageNames, false, 0);
if (sSourceLanguage.Length == 0) return;

string sTargetLanguage = Dialog.Pick("Target Language", aLanguageAbbreviations, aLanguageNames, false, 0);
if (sTargetLanguage.Length == 0) return;

string sExe = App.ProgramDir + @"\Convert\TranLang.exe";
string sSourceFile = App.TempFile;
Encoding en = Encoding.UTF8;
en = null;
Util.String2File(sText, sSourceFile, ref en);
string sTargetFile = sSourceFile;
string sCommand = Util.Quote(sExe) + " " + sSourceLanguage + " " + Util.Quote(sSourceFile) + " " + sTargetLanguage + " " + Util.Quote(sTargetFile);
Util.RunHideWait(sCommand);
en = Encoding.UTF8;
// en = null;
sText = Util.File2String(sTargetFile, ref en);
File.Delete(sSourceFile);
File.Delete(sTargetFile);

if (!IsEmptyWindow()) new MdiChild(this);
child = this.Child;
// child.Text = sResult + ".txt";
child.File = child.Text;
rtb = child.RTB;
rtb.Text = sText;

}

if (menuItem == menuMiscGuardDocument) {
rtb.SetGuard(true);
SetRecent(child.File);
}

if (menuItem == menuMiscNoGuard) {
rtb.SetGuard(false);
SetRecent(child.File);
}

if (menuItem == menuMiscPyBrace) {
if (rtb.SelectionLength == 0) {
AddMessage("All");
iStart = 0;
iEnd = rtb.TextLength;
sFile = Path.GetFileNameWithoutExtension(child.Text);
if (Path.GetExtension(child.Text).ToLower() == ".boo") sFile += ".bob";
else sFile += ".pyb";
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
sFile = "";
}
sText = rtb.GetRange(iStart, iEnd);
sText = PyDent2Brace(sText);

if (sFile.Length == 0) {
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
}
else {
child = new MdiChild(App.Frame, sFile);
Child.RTB.Text = sText;
child.RTB.Modified = true;
}
AddMessage("Done");
}

if (menuItem == menuMiscPyDent) {
sFile = child.File;
string sExt = Path.GetExtension(sFile).ToLower();
sFile = Path.GetFileNameWithoutExtension(sFile);
if (sExt == ".bob") sFile += ".boo";
else sFile += ".py";

if (rtb.SelectionLength == 0) {
AddMessage("All");
iStart = 0;
iEnd = rtb.TextLength;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
sFile = "";
}
sText = rtb.GetRange(iStart, iEnd);

//if (sExt != ".bob" && sExt != ".pyb") sText = PyDent2Brace(sText);
if (sExt == ".boo" || sExt == ".py") sText = PyDent2Brace(sText);
sText = PyBrace2Dent(sText);

if (sFile.Length == 0) {
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
}
else {
child = new MdiChild(App.Frame, sFile);
Child.RTB.Text = sText;
child.RTB.Modified = true;
}
AddMessage("Done");
}

if (menuItem == menuMiscInferIndent) {
sText = rtb.Text;
aResults = Util.RegExpExtractCase(sText, @"^( |\t)+");
if (aResults.Length == 0) {
AddMessage("No indentation found!");
return;
}

string sIndent = aResults[0];
if (this.KeyRepeat % 2 == 0) {
AddMessage("Infer Indent");
if (sIndent.Contains(" ") && sIndent.Contains("\t")) {
foreach (char c in sIndent) {
if (c == ' ') AddMessage("Space");
else AddMessage("Tab");
}
}
else {
if (sIndent.StartsWith(" ")) sText = "space";
else sText = "tab";
AddMessage(Util.Pluralize(sIndent.Length, sText));
}
}
else {
sIndent = sIndent.Replace(" ", @"\040");
sIndent = sIndent.Replace("\t", @"\t");
App.WriteOption("IndentUnit", sIndent);
AddMessage("IndentUnit configured");
}
}

if (menuItem == menuMiscFormatCode) {
// Route the document to the right formatter by extension. The former
// test here, sExt.Contains(sExt), was always true, so EVERY file --
// including Python source -- went through the HTML tidy tool, which
// rewrote it as if it were markup. Now: web markup goes to tidy;
// Python is normalized by the structure engine (indentation rebuilt
// from the IndentUnit setting, spoken "# end" markers refreshed at
// block ends); the C-family goes to astyle; anything else is declined
// with its name instead of being damaged.
sFile = App.Frame.Child.File;
string sExt = Path.GetExtension(sFile).ToLower().TrimStart('.');
string sCommand = "";
sText = "";
if (sExt == "htm" || sExt == "html" || sExt == "xhtml" || sExt == "xml") {
sCommand = "%ProgDir%\\Convert\\Tidy\\tidy.exe -config %ProgDir%\\Convert\\Tidy\\tidy.cfg -m \"%SourceLong%\"";
sCommand = Util.ExpandCommandLine(sCommand, sFile, sFile);
Util.RunHideWait(sCommand);
sText = File.ReadAllText(sFile);
}
else if (sExt == "py" || sExt == "pyw") {
sText = PyBrace2Dent(PyDent2Brace(App.Frame.Child.RTB.Text));
}
else if (sExt == "c" || sExt == "cc" || sExt == "cpp" || sExt == "h" || sExt == "hpp" || sExt == "cs" || sExt == "java" || sExt == "m") {
String sExe = Path.Combine(App.ProgramDir, @"Convert\astyle\astyle.exe");
sExe = Win32.GetShortPath(sExe);
string sSourceFile = Path.GetTempFileName();
sSourceFile = Path.ChangeExtension(sSourceFile, Path.GetExtension(App.Frame.Child.File));
File.WriteAllText(sSourceFile, App.Frame.Child.RTB.Text);
if (File.Exists(App.TempFile)) File.Delete(App.TempFile);
File.Copy(sSourceFile, App.TempFile);
string sIndent = App.ReadOption("IndentUnit", "\t");
sIndent = Util.Literalize(sIndent);
sCommand = sExe + " " + sSourceFile;
if (sIndent == "\t") sCommand = sExe + " --indent=tab " + sSourceFile;
Util.RunHideWait(sCommand);
sText = File.ReadAllText(sSourceFile);
File.Delete(sSourceFile);
}
else {
AddMessage("No formatter for ." + sExt + " files!");
return;
}

if (sText.Length == 0) {
AddMessage(" Error !");
return;
}

App.Frame.Child.RTB.Text = sText;
App.Frame.Child.RTB.Modified = true;
AddMessage(" Done !");
}

if (menuItem == menuMiscRepeatLine) {
sText = CR + rtb.RowText;
iIndex = rtb.RowStart + rtb.RowText.Length;
rtb.ReplaceRange(iIndex, iIndex, sText);
rtb.Index = iIndex + 1;
Util.Say(rtb.RowText);
}

if (menuItem == menuMiscSectionBreak) {
rtb.ReplaceRange(rtb.Index, rtb.Index, SectionBreak);
Util.Say(rtb.RowText);
}

if (menuItem == menuDeleteReplaceRegular) {
if (rtb.SelectionLength == 0) {
sTitle = "Replace All";
iStart = 0;
iEnd = rtb.TextLength;
}
else {
sTitle = "Replace Selected";
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

aLabels = new string[] {"&Match", "&Replace"};
sMatch = App.ReadData("Match", "");
sReplace = App.ReadData("Replace", "");
aValues = new string[] {sMatch, sReplace};
aResults = Dialog.MultiInput(sTitle, aLabels, aValues, new string[] {"ReplaceMatch", "ReplaceWith"});
if (aResults == null || aResults.Length == 0) return;

sMatch = aResults[0];
App.WriteData("Match", sMatch);
sMatch = Util.Literalize(sMatch, true);
sMatch = Regex.Escape(sMatch);
sReplace = aResults[1];
App.WriteData("Replace", sReplace);
sReplace = Util.Literalize(sReplace, true);

sText = rtb.GetRange(iStart, iEnd);
iCount = Util.RegExpCountEquiv(sText, sMatch);
sText = Util.RegExpReplaceEquiv(sText, sMatch, sReplace);
if (iCount > 0) rtb.ReplaceRange(iStart, iEnd, sText);
AddMessage(Util.Pluralize(iCount, "match", "matches"));
}

if (menuItem == menuDeleteReplaceWithRegExp) {
if (rtb.SelectionLength == 0) {
sTitle = "Replace All with Regular Expression";
iStart = 0;
iEnd = rtb.TextLength;
}
else {
sTitle = "Replace Selected with Regular Expression";
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

aLabels = new string[] {"&Pattern", "&Substitute"};
sPattern = App.ReadData("Pattern", "");
sSubstitute = App.ReadData("Substitute", "");
aValues = new string[] {sPattern, sSubstitute};
aResults = Dialog.MultiInput(sTitle, aLabels, aValues, new string[] {"ReplaceRegExpPattern", "ReplaceRegExpWith"});
if (aResults == null || aResults.Length == 0) return;

sPattern = aResults[0];
App.WriteData("Pattern", sPattern);
//sPattern = Util.Literalize(sPattern);
sSubstitute = aResults[1];
App.WriteData("Substitute", sSubstitute);
sSubstitute = Util.Literalize(sSubstitute);
sText = rtb.GetRange(iStart, iEnd);
iCount = Util.RegExpCountCase(sText, sPattern);
sText = Util.RegExpReplaceCase(sText, sPattern, sSubstitute);
if (iCount > 0) rtb.ReplaceRange(iStart, iEnd, sText);
AddMessage(Util.Pluralize(iCount, "match", "matches"));
}

if (menuItem == menuMiscYieldWithRegExp) {
if (rtb.SelectionLength == 0) {
sTitle = "Yield All with Regular Expression";
iStart = 0;
iEnd = rtb.TextLength;
}
else {
sTitle = "Yield Selected with Regular Expression";
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sLabel = "Pattern";
sValue = App.ReadData("Pattern", "");
sResult = Dialog.Input(sTitle, sLabel, sValue);
if (sResult.Length == 0) return;

App.WriteData("Pattern", sResult);
//sResult = Util.Literalize(sResult);
sText = rtb.GetRange(iStart, iEnd);
iCount = Util.RegExpCountCase(sText, sResult);
AddMessage(Util.Pluralize(iCount, "match", "matches"));
}

if (menuItem == menuMiscExtractWithRegExp) {
if (rtb.SelectionLength == 0) {
sTitle = "Extract All with Regular Expression";
iStart = 0;
iEnd = rtb.TextLength;
}
else {
sTitle = "Extract Selected with Regular Expression";
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sLabel = "Pattern";
sValue = App.ReadData("Pattern", "");
sResult = Dialog.Input(sTitle, sLabel, sValue);
if (sResult.Length == 0) return;

App.WriteData("Pattern", sResult);
//sResult = Util.Literalize(sResult);
sText = rtb.GetRange(iStart, iEnd);
aResults = Util.RegExpExtractCase(sText, sResult);
iCount = aResults.Length;
AddMessage(Util.Pluralize(iCount, "match", "matches"));
if (iCount == 0) return;

new MdiChild(this);
rtb = App.Frame.Child.RTB;
sText = String.Join(SectionBreak, aResults);
rtb.ReplaceRange(0, 0, sText);
rtb.Index = 0;
}

if (menuItem == menuMiscRunAtCursor) {
if (rtb.SelectionLength == 0) {
sTitle = "Run Chunk at Cursor";
object[] a = GetChunk();
sText = (string) a[1];
}
else {
sTitle = "Run Selected at Cursor";
sText = rtb.SelectedText;
}

sLabel = "Path";
sReplace = "";
sMatch = "(\r|\n)";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
sMatch = "^(\\<| )+";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
sMatch = "(\\>| |\\.)+$";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);

if (sText.Contains("://")) sText = sText.Trim(); //do nothing
else if (sText.ToLower().StartsWith("www.")) sText = "http://" + sText;
else if (sText.Contains("@") && !sText.ToLower().StartsWith("mailto")) sText = "MailTo:" + sText;

sResult = Dialog.Input(sTitle, sLabel, sText).Trim();
if (sResult.Length == 0) return;
Process.Start(sResult);
}

if (menuItem == menuMiscSpecialCharacter) {
string sCode = App.ReadData("Code", "");
sResult = Dialog.Input("Special Character", "Code:", sCode, "SpecialCharacter").Trim().ToLower();
if (sResult.Length == 0) return;

App.WriteData("Code", sResult);
if (sResult.StartsWith(@"\")) sResult = sResult.Remove(0, 1);
if (sResult.StartsWith("d")) {
sResult = sResult.Remove(0, 1);
try {
int iCode = Int32.Parse(sResult);
sText = Util.Code2String(iCode);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}
}
else {
if (sResult.StartsWith("u")) sResult = sResult.Remove(0, 1);
string s = sResult;
sText = Util.Literalize(@"\u" + sResult.PadLeft(4, '0'));
if (sText.Length == 0 || sText == "\u0000") {
Dialog.Show("Error", "Invalid Unicode number");
return;
}
}

iIndex = rtb.Index;
rtb.ReplaceRange(iIndex, iIndex, sText);
rtb.Index = iIndex + 1;
AddMessage(sText);
}

if (menuItem == menuDeleteHardLine) {
iStart = rtb.RowStart;
iEnd = rtb.Text.IndexOf("\n", iStart);
if (iEnd >= 0) iEnd++;
else iEnd = rtb.TextLength;
rtb.ReplaceRange(iStart, iEnd, "");
rtb.Index = iStart;
Util.Say(rtb.RowText);
}

if (menuItem == menuDeleteParagraph) {
iStart = rtb.RowStart;
sMatch = @"\n\s*\n";
object[] a = Util.RegExpContainsCase(rtb.Text, sMatch, iStart);
iEnd = (int) a[0];
//Dialog.Show(iEnd, ((string) a[1]).Length);
if (iEnd >= 0) iEnd += ((string) a[1]).Length;
else iEnd = rtb.TextLength;
rtb.ReplaceRange(iStart, iEnd, "");
rtb.Index = iStart;
Util.Say(rtb.RowText);
}

if (menuItem == menuDeleteLine) {
iStart = rtb.RowStart;
iEnd = iStart + rtb.RowLength;
rtb.ReplaceRange(iStart, iEnd, "");
rtb.Index = iStart;
Util.Say(rtb.RowText);
}

if (menuItem == menuDeleteRight) {
iStart = rtb.Index;
iEnd = rtb.RowStart + rtb.RowText.Length;
//if (iEnd != rtb.TextLength) iEnd--;
rtb.ReplaceRange(iStart, iEnd, "");
rtb.Index = iStart;
Util.Say(rtb.RowText);
}

if (menuItem == menuDeleteLeft) {
iStart = rtb.RowStart;
iEnd = rtb.Index;
rtb.ReplaceRange(iStart, iEnd, "");
rtb.Index = iStart;
Util.Say(rtb.RowText);
}

if (menuItem == menuDeleteDown) {
iStart = rtb.Index;
iEnd = rtb.TextLength;
rtb.ReplaceRange(iStart, iEnd, "");
rtb.Index = iStart;
Util.Say(rtb.RowText);
}

if (menuItem == menuDeleteUp) {
iStart = 0;
iEnd = rtb.Index;
rtb.ReplaceRange(iStart, iEnd, "");
rtb.Index = iStart;
Util.Say(rtb.RowText);
}

if (menuItem == menuDeleteFile) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

if (Dialog.Confirm("Confirm", "Delete " + child.Text + "?", "N") != "Y") return;
File.Delete(sFile);
child.Close();
}

if (menuItem == menuDeleteTrimBlanks) {
if (rtb.SelectionLength == 0) {
AddMessage("Line");
iStart = rtb.RowStart;
iEnd = iStart + rtb.RowText.Length;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sText = rtb.GetRange(iStart, iEnd);
sText = Util.RegExpReplaceCase(sText, "^( |\t)+", "");
sText = Util.RegExpReplaceCase(sText, "( |\t)+$", "");
sText = Util.RegExpReplaceCase(sText, "\n\n\n\n+", "\n\n\n");
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigateForwardFind || menuItem == menuNavigateForwardFindAtCursor || (!this.FindWithRegExp && menuItem == menuNavigateForwardFindAgain)) {
iIndex = rtb.Index;
iStart = 0;
sText = App.ReadData("Find", "");
if (menuItem == menuNavigateForwardFind) sText = Dialog.Input("Forward Find", "Text", sText, "Find");
else if (menuItem == menuNavigateForwardFindAtCursor) {
if (rtb.SelectionLength == 0) {
object[] a = GetChunk();
iStart = (int) a[0];
sText = (string) a[1];
}
else {
iStart = rtb.SelectionStart;
sText = rtb.SelectedText;
}
}
if (sText.Length == 0) return;

App.WriteData("Find", sText);
this.FindWithRegExp = false;
sText = Util.Literalize(sText, true);
sText = Util.Convert2MacLineBreak(sText);

if (menuItem == menuNavigateForwardFindAtCursor) {
iStart += sText.Length;
iEnd = -1;
}
else if (rtb.SelectionLength == 0) {
iStart = iIndex;
iEnd = -1;
}
else {
iStart = rtb.SelectionStart;
iEnd = rtb.SelectionStart + rtb.SelectionLength;
}

iIndex = rtb.Find(sText, iStart, iEnd, RichTextBoxFinds.NoHighlight);
if (iIndex >= 0) {
rtb.Index = iIndex + sText.Length;
Util.Say(rtb.RowText);
}
else AddMessage("Not found!");
}

if (menuItem == menuNavigateReverseFind || menuItem == menuNavigateReverseFindAtCursor || (!this.FindWithRegExp && menuItem == menuNavigateReverseFindAgain)) {
iIndex = rtb.Index;
iEnd = 0;
sText = App.ReadData("Find", "");
if (menuItem == menuNavigateReverseFind) sText = Dialog.Input("Reverse Find", "Text", sText, "Find");
else if (menuItem == menuNavigateReverseFindAtCursor) {
if (rtb.SelectionLength == 0) {
object[] a = GetChunk();
iEnd = (int) a[0];
sText = (string) a[1];
}
else {
iEnd = rtb.SelectionStart;
sText = rtb.SelectedText;
}
}
if (sText.Length == 0) return;

App.WriteData("Find", sText);
this.FindWithRegExp = false;
sText = Util.Literalize(sText, true);
sText = Util.Convert2MacLineBreak(sText);

if (menuItem == menuNavigateReverseFindAtCursor) {
iStart = 0;
}
else if (rtb.SelectionLength == 0) {
iStart = 0;
iEnd = iIndex;
}
else {
iStart = rtb.SelectionStart;
iEnd = rtb.SelectionStart + rtb.SelectionLength;
}

iIndex = rtb.Find(sText, iStart, iEnd, RichTextBoxFinds.Reverse | RichTextBoxFinds.NoHighlight);
//if (iIndex >= 0) {
if (iIndex >= 0 && iIndex < iEnd) {
rtb.Index = iIndex;
Util.Say(rtb.RowText);
}
else AddMessage("Not found!");
}

if (menuItem == menuNavigateForwardFindWithRegExp || (this.FindWithRegExp && menuItem == menuNavigateForwardFindAgain)) {
sMatch = App.ReadData("Pattern", "");
if (menuItem == menuNavigateForwardFindWithRegExp) sMatch = Dialog.Input("Forward Find with Regular Expression", "Pattern", sMatch, "FindRegExp");
if (sMatch.Length == 0) return;

App.WriteData("Pattern", sMatch);
this.FindWithRegExp = true;

if (rtb.SelectionLength == 0) {
iStart = rtb.Index;
iEnd = rtb.TextLength;
}
else {
iStart = rtb.SelectionStart;
iEnd = rtb.SelectionStart + rtb.SelectionLength;
}
sText = rtb.GetRange(iStart, iEnd);

object[] a = Util.RegExpContainsCase(sText, sMatch);
iIndex = (int) a[0];
if (iIndex >= 0) {
sValue = (string) a[1];
rtb.Index = iStart + iIndex + sValue.Length;
Util.Say(rtb.RowText);
}
else AddMessage("Not found!");
}

if (menuItem == menuNavigateReverseFindWithRegExp || (this.FindWithRegExp && menuItem == menuNavigateReverseFindAgain)) {
sMatch = App.ReadData("Pattern", "");
if (menuItem == menuNavigateReverseFindWithRegExp) sMatch = Dialog.Input("Reverse Find with Regular Expression", "Pattern", sMatch, "FindRegExp");
if (sMatch.Length == 0) return;

App.WriteData("Pattern", sMatch);
this.FindWithRegExp = true;

if (rtb.SelectionLength == 0) {
iStart = 0;
iEnd = rtb.Index;
}
else {
iStart = rtb.SelectionStart;
iEnd = rtb.SelectionStart + rtb.SelectionLength;
}
sText = rtb.GetRange(iStart, iEnd);

object[] a = Util.RegExpContainsLastCase(sText, sMatch);
iIndex = (int) a[0];
if (iIndex >= 0) {
sValue = (string) a[1];
rtb.Index = iStart + iIndex;
Util.Say(rtb.RowText);
}
else AddMessage("Not found!");
}

if (menuItem == menuNavigateJumpToLine || menuItem == menuNavigateJumpToLineAgain) {
sText = App.ReadData("Jump", "");
if (menuItem == menuNavigateJumpToLine) sText = Dialog.Input("Jump to", "Line", sText);
if (sText.Length == 0) return;

string[] a = sText.Split(',');
sLine = a[0].Trim();
if (sLine.Length == 0) sLine = rtb.Line.ToString();
string sColumn = a.Length > 1 ? a[1].Trim() : "1";

try {
iLine = Int32.Parse(sLine);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}

int iColumn = 1;
try {
iColumn = Int32.Parse(sColumn);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}

App.WriteData("Jump", sText);
try {
rtb.Line = iLine;
rtb.Column = iColumn;
Util.Say(rtb.RowText);
}
catch {
Dialog.Show("Error", "Invalid position!");
return;
}
}
if (menuItem == menuNavigateSearchForTopic || menuItem == menuNavigateSearchForTopicAgain) {
sText = App.ReadData("Topic", "");
if (menuItem == menuNavigateSearchForTopic) {
sText = Dialog.Input("Search For", "Topic", sText, "SearchTopic");
iStart = 0;
}
else iStart = rtb.Index;
if (sText.Length == 0) return;

App.WriteData("Topic", sText);
sMatch = SB + ".*?" + sText + ".*?" + LB;
sText = rtb.Text;
iIndex = (int) Util.RegExpContainsEquiv(sText, sMatch, iStart)[0];
if (iIndex == -1) {
AddMessage("Not found!");
return;
}
rtb.Index = iIndex + SB.Length;
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigateGoToPercent || menuItem == menuNavigateGoToPercentAgain) {
sText = App.ReadData("Percent", "");
if (menuItem == menuNavigateGoToPercent) sText = Dialog.Input("Go to", "Percent", sText);
if (sText.Length == 0) return;

try {
iPercent = Int32.Parse(sText);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}

App.WriteData("Percent", sText);
rtb.Percent = iPercent;
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigateGoToPart) {
sText = rtb.Text;
sMatch = @"^\s*((Chapter)|(Section)|(Part))\s+\d";
sMatch = App.ReadOption("NavigatePart", sMatch);
object[][] aMatches = Util.RegExpExtractWithIndex(sText, sMatch, false);
if (aMatches.Length == 0) {
AddMessage("No matches for NavigatePart expression!");
return;
}

iIndex = rtb.Index;
int iPosition = 0;
HomerList lNames = new HomerList();
HomerList lValues = new HomerList();
for (int i = 0; i < aMatches.Length; i++) {
object[] a = (object[]) aMatches[i];
string sIndex = ((int) a[0]).ToString();
string sPart = (string) a[1];
if (iIndex >= Int32.Parse(sIndex)) iPosition = i;
lNames.Add(sPart);
lValues.Add(sIndex);
}

string[] aNames = lNames.ToArray();
aValues = lValues.ToArray();
string s = Dialog.Pick("Go to Part", aValues, aNames, false, iPosition);
if (s.Length == 0) return;

iIndex = Int32.Parse(s);
rtb.Index = iIndex;
}

if (menuItem == menuNavigateSetBookmark) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

sText = App.ReadValue("Favorites", sFile, "");
HomerList hl = new HomerList(sText);
hl.KeepLike(@"\d+");
hl.Remove("-1");
sText = rtb.Index + "|" + (rtb.ReadOnly ? "G" : "M") + "|" + (string) Util.If(rtb.WordWrap, "W", "U");
hl.AddUniqueRange(sText);
sText = hl.Segments;
App.WriteValue("Favorites", sFile, sText);
}

if (menuItem == menuNavigateClearBookmark) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

sText = App.ReadValue("Favorites", sFile, "");
if (sText.Length == 0) return;
HomerList hl = new HomerList(sText);
hl.Remove(rtb.Index.ToString());
sText = hl.Segments;
App.WriteValue("Favorites", sFile, sText);
}

if (menuItem == menuNavigateGoToBookmark) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

sText = App.ReadValue("Favorites", sFile, "");
HomerList hl = new HomerList(sText);
hl.KeepLike(@"\d+");
hl.Remove("-1");
if (hl.Count == 0) {
AddMessage("No bookmark!");
return;
}

if (hl.Count == 1) sResult = hl[0];
else {
hl.PadLeft(hl.MaxLength(), ' ');
hl.Sort();
HomerList hlLines = new HomerList();
foreach (string sIndex in hl) {
iIndex = Int32.Parse(sIndex);
int iRow = rtb.GetLineFromCharIndex(iIndex);
iStart = rtb.GetFirstCharIndexFromLine(iRow);
iEnd = rtb.GetFirstCharIndexFromLine(iRow + 1);
if (iEnd == -1) iEnd = rtb.TextLength;
sLine = rtb.GetRange(iStart, iEnd).Trim();
hlLines.Add(sLine);
}
string[] aDisplay = hlLines.ToArray();
aValues = hl.ToArray();
iIndex = rtb.Index;
int iDefault = -1;
for (int i = 0; i < hl.Count; i++) {
//Dialog.Show(iIndex, Int32.Parse(hl[i]));
if (iIndex < Int32.Parse(hl[i])) {
iDefault = i;
break;
}
}
//Dialog.Show(iDefault);

sResult = Dialog.Pick("Bookmarks", aValues, aDisplay, false, iDefault);
if (sResult.Length == 0) return;
}

rtb.Index = Int32.Parse(sResult);
Util.Say(rtb.RowText);
}

if (menuItem == menuFileSetFavorite) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

sText = App.ReadValue("Favorites", sFile, "");
HomerList hl = new HomerList(sText);
hl.KeepLike(@"\d+");
if (hl.Count == 0) hl.Add("-1");
sText = (rtb.ReadOnly ? "G" : "M") + "|" + (string) Util.If(rtb.WordWrap, "W", "U");
hl.AddUniqueRange(sText);
sText = hl.Segments;
App.WriteValue("Favorites", sFile, sText);
}

if (menuItem == menuFileClearFavorite) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

App.DeleteKey("Favorites", sFile);
}

if (menuItem == menuFileListFavorites) {
aResults = App.ReadSectionKeys("Favorites");
List<string> list = new List<string>(aResults);
for (int i = list.Count - 1; i >=0; i--) {
string s = list[i];
if (File.Exists(s)) continue;
App.DeleteKey("Favorites", s);
list.RemoveAt(i);
}

aResults = list.ToArray();
if (aResults.Length == 0) {
AddMessage("No items!");
return;
}

string[] aDisplay = new string[aResults.Length];
for (int i = 0; i < aDisplay.Length; i++) aDisplay[i] = Path.GetFileName(aResults[i]);
sFile = Dialog.Pick("List Favorites", aResults, aDisplay, true, 0);
if (sFile.Length == 0) return;

OpenOrActivateWindow(sFile, 0);
}

if (menuItem == menuNavigateGoToStartOfSelection) {
rtb.Index = rtb.StartSelection;
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigateHomeCharacter) {
sText = rtb.RowText;
iLength = sText.TrimStart().Length;
if (iLength == 0) return;
iIndex = rtb.RowStart + (sText.Length - iLength);
rtb.Index = iIndex;
sText = rtb.GetRange(iIndex, iIndex + 1);
AddMessage(sText);
}

if (menuItem == menuNavigateEndCharacter) {
sText = rtb.RowText;
iLength = sText.TrimEnd().Length;
if (iLength == 0) return;
iIndex = rtb.RowStart + iLength - 1;
rtb.Index = iIndex;
sText = rtb.GetRange(iIndex, iIndex + 1);
AddMessage(sText);
}

if (menuItem == menuNavigateStartTag) {
iIndex = rtb.Index;
iEnd = rtb.Text.IndexOf(">", rtb.Index);
if (iEnd == -1) {
//AddMessage("Not found!");
//return;
iEnd = iIndex - 1;
}

iStart = 0;
iEnd++;
sText = rtb.GetRange(iStart, iEnd);
sMatch = @"</?\w+( |>)";
object[] a = Util.RegExpContainsLastCase(sText, sMatch);
iStart = (int) a[0];
if (iStart == -1) {
AddMessage("Not found!");
return;
}

string sWord = (string) a[1];
if (sWord.IndexOf("/") == -1) {
if (iStart == iIndex) {
if (iIndex > 0) iIndex = sText.Substring(0, iIndex - 1).LastIndexOf("<");
if (iStart == 0 || iIndex < 0) {
AddMessage("Not found!");
return;
}
}
else iIndex = iStart;
}

else {
sWord = "<" + sWord.Substring(2, sWord.Length - 3);
sMatch = sWord + "( |>)";
iIndex = (int) Util.RegExpContainsEquiv(rtb.Text, sMatch)[0];
if (iIndex == -1) {
AddMessage("Not found!");
return;
}

}

rtb.Index = iIndex;
sText = (string) GetChunk()[1];
Util.Say(sText);
}

if (menuItem == menuNavigateEndTag) {
iIndex = rtb.Index;
iEnd = rtb.Text.IndexOf(">", rtb.Index);
if (iEnd == -1) {
AddMessage("Not found!");
return;
}

iStart = 0;
iEnd++;
sText = rtb.GetRange(iStart, iEnd);
sMatch = @"</?\w+( |>)";
object[] a = Util.RegExpContainsLastCase(sText, sMatch);
iStart = (int) a[0];
if (iStart == -1) {
AddMessage("Not found!");
return;
}

string sWord = (string) a[1];
iStart += sWord.Length - 1;
if (sWord.IndexOf("/") >= 0) {
if (iStart == iIndex) {
if (iIndex < rtb.TextLength - 1) iIndex = rtb.Text.IndexOf(">", iIndex + 1);
if (iStart == rtb.TextLength - 1 || iIndex < 0) {
AddMessage("Not found!");
return;
}
}
else iIndex = iStart;
}

else {
sMatch = "</" + sWord.Substring(1, sWord.Length -2) + ">";
a = Util.RegExpContainsEquiv(rtb.Text, sMatch, iEnd);
iIndex = (int) a[0];
if (iIndex == -1) {
AddMessage("Not found!");
return;
}

iIndex += ((string) a[1]).Length - 1;
}

rtb.Index = iIndex;
sText = (string) GetChunk()[1];
Util.Say(sText);
}

if (menuItem == menuNavigateNextJustify) {
HorizontalAlignment ha = rtb.SelectionAlignment;
bool bBullet = rtb.SelectionBullet;
iIndex = rtb.Index;
iEnd = rtb.TextLength;
while (iIndex < iEnd && rtb.SelectionAlignment == ha && rtb.SelectionBullet == bBullet) {
iIndex++;
rtb.Index = iIndex;
}
if (iIndex == iEnd) AddMessage("Bottom!");
if (rtb.SelectionAlignment != ha || rtb.SelectionBullet != bBullet) {
sText = GetJustifyText();
AddMessage(sText);
}
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigatePriorJustify) {
HorizontalAlignment ha = rtb.SelectionAlignment;
bool bBullet = rtb.SelectionBullet;
iIndex = rtb.Index;
iStart = 0;
while (iIndex > iStart && rtb.SelectionAlignment == ha && rtb.SelectionBullet == bBullet) {
iIndex--;
rtb.Index = iIndex;
}
if (iIndex == iStart) AddMessage("Top!");
if (rtb.SelectionAlignment != ha || rtb.SelectionBullet != bBullet) {
sText = GetJustifyText();
AddMessage(sText);
}
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigateNextStyle) {
bool bBold = rtb.SelectionFont.Bold;
bool bItalic = rtb.SelectionFont.Italic;
bool bUnderline = rtb.SelectionFont.Underline;
iIndex = rtb.Index;
iEnd = rtb.TextLength;
while (iIndex < iEnd && rtb.SelectionFont.Bold == bBold && rtb.SelectionFont.Italic == bItalic && rtb.SelectionFont.Underline == bUnderline) {
iIndex++;
rtb.Index = iIndex;
}

if (iIndex == iEnd) AddMessage("Bottom!");
if (!(rtb.SelectionFont.Bold == bBold && rtb.SelectionFont.Italic == bItalic && rtb.SelectionFont.Underline == bUnderline)) {
sText = GetStyleText();
AddMessage(sText);
}
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigatePriorStyle) {
bool bBold = rtb.SelectionFont.Bold;
bool bItalic = rtb.SelectionFont.Italic;
bool bUnderline = rtb.SelectionFont.Underline;
iIndex = rtb.Index;
iStart = 0;
while (iIndex > iStart && rtb.SelectionFont.Bold == bBold && rtb.SelectionFont.Italic == bItalic && rtb.SelectionFont.Underline == bUnderline) {
iIndex--;
rtb.Index = iIndex;
}

if (iIndex == iStart) AddMessage("Top!");
if (!(rtb.SelectionFont.Bold == bBold && rtb.SelectionFont.Italic == bItalic && rtb.SelectionFont.Underline == bUnderline)) {
sText = GetStyleText();
AddMessage(sText);
}
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigateNextBaseline) {
int iOffset = rtb.SelectionCharOffset;
iIndex = rtb.Index;
iEnd = rtb.TextLength;
while (iIndex < iEnd && rtb.SelectionCharOffset == iOffset) {
iIndex++;
rtb.Index = iIndex;
}
if (iIndex == iEnd) AddMessage("Bottom!");
if (rtb.SelectionCharOffset != iOffset) {
sText = GetBaselineText();
AddMessage(sText);
}
if (iIndex < iEnd) {
sText = rtb.GetRange(iIndex, iIndex + 1);
AddMessage(sText);
}
}

if (menuItem == menuNavigatePriorBaseline) {
int iOffset = rtb.SelectionCharOffset;
iIndex = rtb.Index;
iStart = 0;
while (iIndex > iStart && rtb.SelectionCharOffset == iOffset) {
iIndex--;
rtb.Index = iIndex;
}
if (iIndex == iStart) AddMessage("Top!");
if (rtb.SelectionCharOffset != iOffset) {
sText = GetBaselineText();
AddMessage(sText);
}
if (rtb.TextLength > 0) {
sText = rtb.GetRange(iIndex, iIndex + 1);
AddMessage(sText);
}
}

if (menuItem == menuNavigateNextFont) {
string sFont = Util.Font2String(rtb.SelectionFont);
string sColor = Util.Color2String(rtb.SelectionColor);
iIndex = rtb.Index;
iEnd = rtb.TextLength;
while (iIndex < iEnd && Util.Font2String(rtb.SelectionFont) == sFont && Util.Color2String(rtb.SelectionColor) == sColor) {
iIndex++;
rtb.Index = iIndex;
}
if (iIndex == iEnd) AddMessage("Bottom!");
if (!(Util.Font2String(rtb.SelectionFont) == sFont && Util.Color2String(rtb.SelectionColor) == sColor)) {
sText = Util.Font2String(rtb.SelectionFont);
//if (Util.Color2String(rtb.SelectionColor) != sColor) sText+= ", Color " + Util.Color2String(rtb.SelectionColor);
if (Util.Color2String(rtb.SelectionColor) != sColor) sText = GetFontText(rtb.SelectionFont, rtb.SelectionColor);
AddMessage(sText);
}
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigatePriorFont) {
string sFont = Util.Font2String(rtb.SelectionFont);
string sColor = Util.Color2String(rtb.SelectionColor);
iIndex = rtb.Index;
iStart = 0;
while (iIndex > iStart && Util.Font2String(rtb.SelectionFont) == sFont && Util.Color2String(rtb.SelectionColor) == sColor) {
iIndex--;
rtb.Index = iIndex;
}
if (iIndex == iStart) AddMessage("Top!");
if (!(Util.Font2String(rtb.SelectionFont) == sFont && Util.Color2String(rtb.SelectionColor) == sColor)) {
sText = Util.Font2String(rtb.SelectionFont);
if (Util.Color2String(rtb.SelectionColor) != sColor) sText = GetFontText(rtb.SelectionFont, rtb.SelectionColor);
AddMessage(sText);
}
Util.Say(rtb.RowText);
}

if (menuItem == menuQueryBlock) {
char[] a = {' ', '\t'};
sLine = "";
string sComment = App.ReadOption("QuotePrefix", "> ");
int iLevels = GetIndent();
int i = iLevels;
int iRow = rtb.Row;
int iTop = 0;
string sPre = "";
if (this.KeyRepeat % 2 != 0) {
while (iRow > iTop) {
// Util.Say(iRow);
iRow--;
sLine = rtb.GetRowText(iRow).Trim(a);
sPre = sLine + "\n" + sPre;
if (sLine.Length == 0 || sLine.StartsWith(sComment)) continue;
i = GetIndent(iRow);
if (iLevels > i) break;
}
}

i = iLevels;
iRow = rtb.Row;
string sRest = rtb.GetRowText(iRow);
sRest = "";
iCount = rtb.Lines.Length;
while (iRow < iCount) {
// Util.Say(iRow);
sLine = rtb.GetRowText(iRow).Trim();
i = GetIndent(iRow);
if (sLine.Length == 0 || sLine.StartsWith(sComment) || i >= iLevels) {
sRest += sLine + "\n";
}
else break;
if ((iRow == rtb.Row + 1) && (i > iLevels)) iLevels = i;
iRow++;
}
sText = sPre + sRest;
AddMessage(sText);
}

if (menuItem == menuNavigateRightBrace || menuItem == menuNavigateLeftBrace || menuItem == menuQueryBraces) {
if (rtb.Text.Length == 0) {
AddMessage("No text!");
return;
}

sText = App.ReadOption("BraceMatch", "{}");
string sLeft = sText.Substring(0, 1);
string sRight = sText.Substring(1, 1);
iIndex = rtb.Index;
string s = rtb.GetRange(iIndex, iIndex + 1);
switch (s) {
case "{" :
case "}" :
sLeft = "{";
sRight = "}";
break;
case "<" :
case ">" :
sLeft = "<";
sRight = ">" ;
break;
case "[" :
case "]" :
sLeft = "[";
sRight = "]";
break;
case "(" :
case ")" :
sLeft = "(";
sRight = ")";
break;
}

if (menuItem == menuNavigateRightBrace) {
iStart = iIndex;
iEnd = rtb.TextLength;
sText = rtb.GetRange(iStart, iEnd);
iCount = 0;
int i = 0;
// Dialog.Show(i, sText.Length);
bool bLoop = true;
while (bLoop) {
if (i == sText.Length) {
bLoop = false;
AddMessage("Not found!");
}
else if (sText.Substring(i, 1) == sLeft && i > 0) {
iCount++;
i++;
}
else if (sText.Substring(i, 1) == sRight && iCount > 0) {
iCount --;
i++;
}
else if (sText.Substring(i, 1) == sRight && iCount == 0 && i > 0) {
bLoop = false;
iIndex = iStart + i;
rtb.Index = iIndex;
}
else i++;
}
sText = (string) GetChunk()[1];
Util.Say(sText);
//Util.Say(rtb.RowText);
}
else if (menuItem == menuNavigateLeftBrace) {
iStart = 0;
iEnd = iIndex;
sText = rtb.GetRange(iStart, iEnd);
sText = Util.Reverse(sText);
iCount = 0;
int i = 0;
bool bLoop = true;
while (bLoop) {
if (i == sText.Length) {
bLoop = false;
AddMessage("Not found!");
}
else if (sText.Substring(i, 1) == sRight) {
iCount++;
i++;
}
else if (sText.Substring(i, 1) == sLeft && iCount > 0) {
iCount --;
i++;
}
else if (sText.Substring(i, 1) == sLeft && iCount == 0) {
bLoop = false;
iIndex = iEnd - i - 1;
rtb.Index = iIndex;
}
else i++;
}
//Util.Say(rtb.RowText);
sText = (string) GetChunk()[1];
Util.Say(sText);
}
else if (menuItem == menuQueryBraces) {
iStart = iIndex;
iEnd = rtb.TextLength;
sText = rtb.GetRange(iStart, iEnd);
iCount = 0;
int i = 0;
bool bLoop = true;
while (bLoop) {
if (i == sText.Length) {
bLoop = false;
}
else if (sText.Substring(i, 1) == sLeft && i > 0) {
iCount++;
i++;
}
else if (sText.Substring(i, 1) == sRight) {
iCount--;
i++;
}
else i++;
}
int iOutLevels = iCount;
iStart = 0;
iEnd = iIndex;
sText = rtb.GetRange(iStart, iEnd);
sText = Util.Reverse(sText);
iCount = 0;
i = 0;
bLoop = true;
while (bLoop) {
if (i == sText.Length) {
bLoop = false;
}
else if (sText.Substring(i, 1) == sRight) {
iCount++;
i++;
}
else if (sText.Substring(i, 1) == sLeft) {
iCount--;
i++;
}
else i++;
}
int iInLevels = iCount;
AddMessage(Util.Absolute(iInLevels) + " left");
AddMessage(Util.Absolute(iOutLevels) + " right");
}
}

if (menuItem == menuNavigateNextBlock) {
char[] a = {' ', '\t'};
sLine = "";
string sComment = App.ReadOption("QuotePrefix", "> ");
int iLevels = GetIndent();
int i = iLevels;
int iRow = rtb.Row;
string sPre = "";
i = iLevels;
iRow = rtb.Row;
string sRest = rtb.GetRowText(iRow);
sRest = "";
iCount = rtb.Lines.Length;
// Dialog.Show(iRow, iCount);
if (iRow >= iCount - 1) {
AddMessage("Bottom!");
return;
}

bool bNested = false;
while (iRow < iCount) {
// Util.Say(iRow);
sLine = rtb.GetRowText(iRow).Trim();
i = GetIndent(iRow);
if (!bNested && i > iLevels) {
iLevels ++;
bNested = true;
}
if (sLine.Length == 0 || sLine.StartsWith(sComment) || i >= iLevels) {
sRest += sLine + "\n";
}
else break;
// if ((iRow == rtb.Row + 1) && (i > iLevels)) iLevels = i;
iRow++;
}
sText = sPre + sRest;
if (iRow == iCount) {
AddMessage("Bottom!");
rtb.Row = iRow - 1;
return;
}

rtb.Row = iRow;
AddMessage(rtb.RowText);
}

if (menuItem == menuNavigatePriorBlock) {
char[] a = {' ', '\t'};
sLine = "";
string sComment = App.ReadOption("QuotePrefix", "> ");
int iLevels = GetIndent();
int i = iLevels;
int iRow = rtb.Row;
int iTop = 0;
if (iRow == iTop) {
AddMessage("Top!");
return;
}

string sPre = "";
bool bNested = false;
while (iRow > iTop) {
// Util.Say(iRow);
iRow--;
sLine = rtb.GetRowText(iRow).Trim(a);
sPre = sLine + "\n" + sPre;
if (sLine.Length == 0 || sLine.StartsWith(sComment)) continue;
i = GetIndent(iRow);
if (!bNested && i > iLevels) {
iLevels ++;
bNested = true;
}
if (iLevels > i) break;
}

rtb.Row = iRow;
AddMessage(rtb.RowText);
}

if (menuItem == menuNavigateNextIndent) {
string sComment = App.ReadOption("QuotePrefix", "> ");
//sComment = Util.Literalize(sComment);
int iLevels = GetIndent();
int i = iLevels;
int iRow = rtb.Row;
int iBottom = rtb.BottomRow;
//rtb.BeginUpdate();
while (iRow < iBottom) {
iRow++;
rtb.Row = iRow;
sLine = rtb.RowText.Trim();
if (sLine.Length == 0 || sLine.StartsWith(sComment)) continue;
i = GetIndent();
if (iLevels != i) break;
}
//rtb.EndUpdate();

if (iLevels == i) {
AddMessage("Bottom!");
rtb.Index = rtb.TextLength;
}
else AddMessage(GetDelta(iLevels, i));
//Util.Say(rtb.RowText);
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigatePriorIndent) {
string sComment = App.ReadOption("QuotePrefix", "> ");
//sComment = Util.Literalize(sComment);
int iLevels = GetIndent();
int i = iLevels;
int iRow = rtb.Row;
int iTop = 0;
//rtb.BeginUpdate();
while (iRow > iTop) {
iRow--;
rtb.Row = iRow;
sLine = rtb.RowText.Trim();
if (sLine.Length == 0 || sLine.StartsWith(sComment)) continue;
i = GetIndent();
if (iLevels != i) break;
}
//rtb.EndUpdate();

if (iLevels == i) {
AddMessage("Top!");
rtb.Index = 0;
}
else AddMessage(GetDelta(iLevels, i));
//Util.Say(rtb.RowText);
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigateNextChunk) {
NavigateNextMatch(App.MatchChunk);
}

if (menuItem == menuNavigatePriorChunk) {
NavigatePriorMatch(App.MatchChunk);
}

if (menuItem == menuNavigateNextSentence) {
NavigateNextMatch(App.MatchSentence);
}

if (menuItem == menuNavigatePriorSentence) {
NavigatePriorMatch(App.MatchSentence);
}

if (menuItem == menuNavigateNextParagraph) {
NavigateNextMatch(App.MatchParagraph);
}

if (menuItem == menuNavigatePriorParagraph) {
NavigatePriorMatch(App.MatchParagraph);
}

if (menuItem == menuNavigateNextPart) {
sMatch = @"^\s*((Chapter)|(Section)|(Part))\s+\d";
sMatch = App.ReadOption("NavigatePart", sMatch);
NavigateNextMatch(sMatch, true);
}

if (menuItem == menuNavigatePriorPart) {
sMatch = @"^\s*((Chapter)|(Section)|(Part))\s+\d";
sMatch = App.ReadOption("NavigatePart", sMatch);
NavigatePriorMatch(sMatch, true);
}

if (menuItem == menuNavigateNextSection) {
iStart = rtb.Index;
sText = rtb.Text;
iEnd = rtb.TextLength;
iIndex = sText.IndexOf(SB, iStart);
if (iIndex == -1) {
AddMessage("Bottom!");
rtb.Index = iEnd;
sLine = rtb.Lines[rtb.Lines.Length - 1];
return;
}
else {
rtb.Index = iIndex + 2;
sText = sText.Substring(0, iIndex);
string[] aText = sText.Split('\n');
iLine = aText.Length;
sLine = rtb.Lines[iLine];
}
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigatePriorSection) {
iEnd = rtb.Index;
if (iEnd > 0) iEnd--;
sText = rtb.Text.Substring(0, iEnd);
iIndex = sText.LastIndexOf(SB);
if (iIndex == -1) {
AddMessage("Top!");
rtb.Index = 0;
}
else {
rtb.Index = iIndex + 2;
Util.Say(rtb.RowText);
}
}

if (menuItem == menuNavigateGoToSection) {
//sLine = SB + rtb.RowText;
sLine = SB + rtb.RowText + LF;
sText = rtb.Text;
iIndex = sText.IndexOf(sLine);
if (iIndex == -1) {
AddMessage("Not found!");
}
else {
rtb.Index = iIndex + SB.Length;
Util.Say(rtb.RowText);
}
}

if (menuItem == menuNavigateGoToContents) {
sText = rtb.Text;
iIndex = sText.IndexOf(LB, rtb.Index);
if (iIndex == -1) {
AddMessage("Not found!");
return;
}
iEnd = iIndex + LB.Length;
sText = sText.Substring(0, iEnd);
iIndex = sText.LastIndexOf(SB);
if (iIndex == -1) {
AddMessage("Not found!");
return;
}
iStart = iIndex + SB.Length;
iIndex = sText.IndexOf(LB, iStart);
if (iIndex == -1) {
AddMessage("Not found!");
return;
}
iEnd = iIndex + LB.Length;
sLine = LB + sText.Substring(iStart, iEnd - iStart);
iIndex = sText.IndexOf(sLine);
if (iIndex == -1) {
AddMessage("Not found!");
return;
}
rtb.Index = iIndex + LB.Length;
Util.Say(rtb.RowText);
}

if (menuItem == menuNavigateSearchForTopic) {
}

if (menuItem == menuQueryAddress) {
if (this.KeyRepeat % 2 == 0) SetStatusAddress(null, null);
else if (App.ReadOption("HardPageAddress", "N").ToLower().Substring(0, 1) != "y") AddMessage(GetPageAddress(rtb));
else AddMessage(GetPercentAddress(rtb));
}

if (menuItem == menuQueryIndent) {
// First press: the line's indentation level, measured with the
// document's own detected unit. Second press: the whole chain of
// enclosing blocks, outermost first -- for Python, the class, the
// function, and each nested statement the cursor sits inside; for a
// brace language, each enclosing opener. This answers the question
// "where am I?" that indentation shows a sighted reader at a glance.
if (this.KeyRepeat % 2 == 0) {
AddMessage("Level " + this.GetIndent());
return;
}

string sQuoteComment = App.ReadOption("QuotePrefix", "> ");
int iCurrentLevels = GetIndent();
int iLowest = iCurrentLevels;
List<string> lsChain = new List<string>();
int iChainRow = rtb.Row;
while (iChainRow > 0) {
iChainRow--;
string sChainLine = rtb.GetRowText(iChainRow);
string sChainTrim = sChainLine.Trim();
if (sChainTrim.Length == 0 || sChainTrim.StartsWith(sQuoteComment) || sChainTrim.StartsWith("#")) continue;
int iRowLevels = GetIndent(iChainRow);
if (iRowLevels < iLowest) {
lsChain.Insert(0, sChainTrim);
iLowest = iRowLevels;
if (iRowLevels == 0) break;
}
}
if (lsChain.Count == 0) AddMessage("Top level");
else AddMessage(String.Join(", ", lsChain.ToArray()));
}

if (menuItem == menuQueryPath) {
sText = child.File;
if (this.KeyRepeat % 2 == 0) AddMessage(sText);
else {
Util.Spell(sText);
}
}

if (menuItem == menuQueryTopic) {
sText = rtb.Text;
iIndex = sText.IndexOf(LB, rtb.Index);
if (iIndex == -1) {
AddMessage("Not found!");
return;
}
iEnd = iIndex + LB.Length;
sText = sText.Substring(0, iEnd);
iIndex = sText.LastIndexOf(SB);
if (iIndex == -1) {
AddMessage("Not found!");
return;
}
iStart = iIndex + SB.Length;
iIndex = sText.IndexOf(LB, iStart);
if (iIndex == -1) {
AddMessage("Not found!");
return;
}
iEnd = iIndex + LB.Length;
sLine = sText.Substring(iStart, iEnd - iStart);
sText = sLine;
if (this.KeyRepeat % 2 == 0) AddMessage(sText);
else {
Util.Spell(sText);
}
}

if (menuItem == menuQueryYield) {
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
if (iStart == iEnd) {
AddMessage("All");
iStart = 0;
iEnd = rtb.TextLength;
}
else  AddMessage("Selected");

sText = rtb.GetRange(iStart, iEnd);
iResult = sText.Length;
AddMessage(Util.Pluralize(iResult, "character"));

if (iResult > 0) {
iResult = Util.RegExpCountCase(sText, "\\S+");
AddMessage(Util.Pluralize(iResult, "Word"));
iResult = Util.RegExpCountCase(sText, LB) + 1;
AddMessage("\t" + Util.Pluralize(iResult, "Line"));
}

}

if (menuItem == menuQueryStatus) {
if (this.KeyRepeat % 2 == 0) {
sText = rtb.Modified ? "Modified" : "Unmodified" + "\t";
sText += rtb.WordWrap ? "Wrap" : "Unwrap";
sText += rtb.ReadOnly ? "Guard" : "";
//sText += App.ReadData("Compiler", "Default");
}
else {
if (child == null || child.YieldEncoding == null) sText = "No disk file with encoding information is open!";
else if (child.IsUnixLineBreak) sText = "Encoding Unicode (UTF-8N)";
else sText = "Encoding " + child.YieldEncoding.EncodingName + " = " + child.YieldEncoding.CodePage;
}
AddMessage(sText);
}

if (menuItem == menuQueryCompiler) {
AddMessage("Compiler " + App.ReadData("Compiler", "Default"));
AddMessage("Folder " + Path.GetFileName(Directory.GetCurrentDirectory()));
}

if (menuItem == menuQuerySelected) {
sText = rtb.SelectedText;
if (sText.Length == 0) sText = "No text!";
if (this.KeyRepeat % 2 == 0) AddMessage(sText);
else {
Util.Spell(sText);
}
}

if (menuItem == menuQueryChunk) {
object[] a = GetChunk();
sText = (string) a[1];
if (sText.Length == 0) sText = "No text!";
if (this.KeyRepeat % 2 == 0) AddMessage(sText);
else {
Util.Spell(sText);
}
}

if (menuItem == menuQueryReadAll) {
sText = rtb.Text;
if (sText.Length == 0) sText = "No text!";
sText = Util.ConvertQuotes(sText);
if (this.KeyRepeat % 2 == 0) AddMessage(sText);
else {
Util.Spell(sText);
}
}

if (menuItem == menuQueryWindowsOpen) {
WindowsOpen();
}

if (menuItem == menuQueryClipboard) {
sText = Util.GetClipboardText();
if (sText.Length == 0) sText = "No text!";
if (this.KeyRepeat % 2 == 0) AddMessage(sText);
else {
Util.Spell(sText);
// SetStatus(sText);
}
}

if (menuItem == menuQueryTime || menuItem == menuMiscInsertTime) {
string sDate, sTime;
GetDateAndTime(out sDate, out sTime);
//sText = dt.ToShortTimeString() + " on " + dt.ToLongDateString();
// sText = sTime + " on " + sDate;
sText = sTime;
if (sTime.Length > 0 && sDate.Length > 0) sText += " ";
sText += sDate;
if (menuItem == menuQueryTime) {
if (sTime.Length > 0) AddMessage(sTime);
if (sDate.Length > 0) AddMessage(sDate);
}
else {
rtb.ReplaceRange(rtb.Index, rtb.Index, sText);
Util.Say(rtb.RowText);
}
}

if (menuItem == menuQueryStyles) {
sText = GetStyleText() + " ";
sText += GetJustifyText() + " ";
sText += GetBaselineText() + " ";
sText = sText.Replace("Regular ", "");
sText = sText.Replace("Left ", "");
sText = sText.Replace("Flat ", "");
if (sText.Trim().Length == 0) sText = "Default";
AddMessage(sText);
}

if (menuItem == menuQueryFont) {
sText = GetFontText(rtb.SelectionFont, rtb.SelectionColor);
AddMessage(sText);
}

if (menuItem == menuMiscCalculateDate) {
CalculateDate();
}

if (menuItem == menuMiscHTMLFormat) {
// Treat the current document as Markdown source and convert it to a complete
// HTML document with the embedded Markdig library, then open the result in a
// new window.
string sBaseName = (child.File != null && child.File.Length > 0) ? Path.GetFileNameWithoutExtension(child.File) : "Untitled";
string sHtml = Util.Markdown2Html(rtb.Text, sBaseName);
if (sHtml.Length > 0) {
if (child.File != null && child.File.IndexOf(@":\") > 0) Directory.SetCurrentDirectory(Path.GetDirectoryName(child.File));
sFile = sBaseName + ".htm";
new MdiChild(this, sFile);
this.Child.File = sFile;
this.Child.RTB.Text = sHtml;
this.Child.RTB.Modified = false;
}
}

if (menuItem == menuMiscMarkdownToText) {
// Treat the current document as Markdown source and render it as plain
// text with the Markdig library: markup stripped, content and reading
// order preserved. Useful as a speech-friendly view of a Markdown
// document. The result opens in a new window.
string sBaseName = (child.File != null && child.File.Length > 0) ? Path.GetFileNameWithoutExtension(child.File) : "Untitled";
sText = Util.Markdown2Text(rtb.Text);
if (sText.Length > 0) {
if (child.File != null && child.File.IndexOf(@":\") > 0) Directory.SetCurrentDirectory(Path.GetDirectoryName(child.File));
sFile = sBaseName + ".txt";
new MdiChild(this, sFile);
this.Child.File = sFile;
this.Child.RTB.Text = sText;
this.Child.RTB.Modified = false;
}
}

if (menuItem == menuMiscHtmlToMarkdown) {
// Treat the current document as HTML source and convert it to Markdown
// with the ReverseMarkdown library (built on HtmlAgilityPack). The
// result opens in a new window.
string sBaseName = (child.File != null && child.File.Length > 0) ? Path.GetFileNameWithoutExtension(child.File) : "Untitled";
sText = Util.Html2Markdown(rtb.Text);
if (sText.Length > 0) {
if (child.File != null && child.File.IndexOf(@":\") > 0) Directory.SetCurrentDirectory(Path.GetDirectoryName(child.File));
sFile = sBaseName + ".md";
new MdiChild(this, sFile);
this.Child.File = sFile;
this.Child.RTB.Text = sText;
this.Child.RTB.Modified = false;
}
}

if (menuItem == menuMiscHtmlToText) {
// Treat the current document as HTML source and render it as plain
// text by chaining HTML to Markdown to plain text, which preserves
// paragraph breaks and list structure that a raw text extraction
// loses. The result opens in a new window.
string sBaseName = (child.File != null && child.File.Length > 0) ? Path.GetFileNameWithoutExtension(child.File) : "Untitled";
sText = Util.Html2Text(rtb.Text);
if (sText.Length > 0) {
if (child.File != null && child.File.IndexOf(@":\") > 0) Directory.SetCurrentDirectory(Path.GetDirectoryName(child.File));
sFile = sBaseName + ".txt";
new MdiChild(this, sFile);
this.Child.File = sFile;
this.Child.RTB.Text = sText;
this.Child.RTB.Modified = false;
}
}

if (menuItem == menuMiscPreviewMarkdown) {
// Render the current document's Markdown as HTML in an embedded web
// view, where the screen reader's own virtual buffer applies: JAWS and
// NVDA element-navigation keys (H for headings, K for links, T for
// tables, arrows for reading) work as on any web page. One keystroke
// replaces the convert, save, and render-in-browser sequence when the
// goal is reading rather than a saved file. Escape or Alt+F4 returns
// to the editor.
string sBaseName = (child.File != null && child.File.Length > 0) ? Path.GetFileNameWithoutExtension(child.File) : "Untitled";
string sHtml = Util.Markdown2Html(rtb.Text, sBaseName);
if (sHtml.Length > 0) PreviewForm.ShowPreview(this, sBaseName, sHtml);
}

if (menuItem == menuMiscPreviewMarkdownBrowser) {
// Render the current document's Markdown in the default web browser,
// via a temporary HTML file that EdSharp deletes at exit. One
// keystroke replaces the convert, save, and F5 sequence.
string sBaseName = (child.File != null && child.File.Length > 0) ? Path.GetFileNameWithoutExtension(child.File) : "Untitled";
string sHtml = Util.Markdown2Html(rtb.Text, sBaseName);
if (sHtml.Length > 0) {
string sHtmlFile = Path.Combine(Path.GetTempPath(), "edsharp_preview_" + Guid.NewGuid().ToString("N") + ".htm");
Util.String2FileU(sHtml, sHtmlFile);
App.TempFiles.Add(sHtmlFile);
Process.Start(sHtmlFile);
}
}

if (menuItem == menuMiscCheckMarkdown) {
// Check the current document as Markdown and report findings, one per
// line, in a new window. Pandoc has no lint mode and no solid .NET
// linting package exists, so these are EdSharp's own line-based rules,
// chosen for accessibility and correct conversion rather than style
// pedantry: heading levels that jump, images without alt text, bare
// URLs, duplicate headings, unclosed code fences, pipe-table rows whose
// cell counts disagree with their header, and reference links that are
// used but never defined or defined but never used. Fenced code blocks
// are skipped, so code examples do not raise false alarms.
string sReport = checkMarkdown(rtb.Text);
// The report opens in a plain new window that EdSharp titles itself,
// NoName style, like List Different Items -- temporary output earns a
// name only when the person decides to save it.
child = new MdiChild(this);
this.Child.RTB.Text = sReport;
this.Child.RTB.Modified = false;
this.Child.RTB.Index = 0;
}

if (menuItem == menuMiscRunCodeBlocks) {
// Run the document's executable code blocks and insert each block's
// results right below it. Two block languages are supported. A fenced
// block whose info string is sql runs through sqlean.exe, the SQLite
// command line shipped in the program folder: the fence line may name
// the database after the language (three backticks, sql, then a path),
// and with no path a .db file with the document's own base name in the
// document's folder is assumed. Query results arrive as a real
// Markdown table, so the .mdx pipeline carries them into docx and HTML
// as real tables. A fenced block whose info string is jscript runs
// through the built-in JScript .NET evaluator, and whatever text it
// returns is inserted as it is. Results live between output-begins and
// output-ends comment markers, and running the command again REPLACES
// the marked region, so a document can be refreshed repeatedly.
// Execution happens only by this explicit command -- never during
// conversion or export -- so opening or converting a document can
// never run code by surprise.
int iBlocksRun = runCodeBlocks();
AddMessage(Util.Pluralize(iBlocksRun, "block") + " run");
}

if (menuItem == menuMiscChatWithAI || menuItem == menuMiscChatWithDocument) {
// A simple, flexible chat with a local model through Ollama. The
// selection, or with no selection the whole document, travels with
// your instruction; the model's answer opens in a NEW window, so the
// source window is never touched and the reply is immediately a
// document -- readable with ordinary keys, savable, or convertible.
// This serves both conversation (empty document, just ask) and
// transformation (select text, say what to do to it). Everything runs
// locally: the model, the data, and the answer never leave the
// machine. The Ollama service installs from a checkbox at the end of
// EdSharp's setup and is shared with every other application that
// uses Ollama.
// What travels with the question. A selection is an explicit choice and
// always goes. With nothing selected, the instruction decides: asking
// the model to summarize, translate, or fix "this" clearly means the
// open document, while a general question -- Scott's "When is
// Thanksgiving this year?" of 25 August 2026 -- does not, and sending a
// whole document with it both slows the answer to minutes and drags the
// model's attention onto the wrong material. Shift+F12, Chat about
// Document, sends the open text whatever the wording, for when the
// instruction gives no hint -- the selection when there is one, the
// whole document otherwise. The status line names the choice, so it is
// never a mystery.
string sContext = "";
// A prompt is often several lines -- an instruction, then an example, or
// a list of things to change -- so the box is a proper multiline one:
// Enter starts a new line, Tab reaches OK and then Cancel, and
// Control+Enter submits from anywhere in the dialog.
string sInstruction = Dialog.Prompt("AI Chat", "Prompt", App.ReadData("ChatInstruction", "")).Trim();
if (sInstruction.Length == 0) return;
App.WriteData("ChatInstruction", sInstruction);
if (rtb.SelectionLength > 0) { sContext = rtb.SelectedText; AddMessage("With selection"); }
else if (menuItem == menuMiscChatWithDocument) { sContext = rtb.Text; AddMessage("With document"); }
else if (rtb.Text.Trim().Length > 0 && instructionWantsDocument(sInstruction)) { sContext = rtb.Text; AddMessage("With document"); }
else AddMessage("Question only");
string sModel = App.ReadOption("OllamaModel", "llama3.2");
AddMessage("Asking " + sModel);
string sPrompt = sInstruction;
if (sContext.Trim().Length > 0) sPrompt += "\n\n" + sContext;
// A local model can think for minutes on a long document, and a silent
// wait is indistinguishable from a hang. The request runs on a worker
// thread while this loop keeps the interface alive and speaks a
// succinct count every fifteen seconds, directly to the running screen
// reader, until the answer lands.
string sAnswer = "";
bool bAnswerDone = false;
System.Threading.Thread threadAsk = new System.Threading.Thread(delegate() {
try { sAnswer = askOllama(sPrompt, sModel); }
catch (Exception) { sAnswer = ""; }
bAnswerDone = true;
});
threadAsk.IsBackground = true;
threadAsk.Start();
int iWaited = 0;
int iSpoken = 0;
while (!bAnswerDone) {
System.Threading.Thread.Sleep(100);
Application.DoEvents();
iWaited += 100;
if (iWaited - iSpoken >= 15000) {
iSpoken = iWaited;
Util.Say((iWaited / 1000) + " seconds");
}
}
if (sAnswer.Length == 0) return;
Util.Say("Answer ready");
// The answer opens in a plain new window that EdSharp titles itself,
// NoName style -- the same pattern as List Different Items and Query
// Common Items. Temporary output earns a name only when the person
// decides to save it.
child = new MdiChild(this);
this.Child.RTB.Text = sAnswer.Replace("\r\n", "\n").Replace("\n", "\r\n");
this.Child.RTB.Modified = false;
this.Child.RTB.Index = 0;
AddMessage("Done");
}

if (menuItem == menuMiscTextConvert || menuItem == menuMiscTextCombine) {
List<string> list = new List<string>();
aResults = rtb.Lines;
string sDir = Directory.GetCurrentDirectory();
string sTempDir = "";
for (int i = 0; i < aResults.Length; i++) {
string s = aResults[i].Trim();
if (s.Length == 0) continue;
sTempDir = Path.GetDirectoryName(s);
if (sTempDir.Length == 0) s = Path.Combine(sDir, s);
else if (Directory.Exists(sTempDir)) sDir = sTempDir;
if (File.Exists(s)) list.Add(s);
}

aResults = list.ToArray();
if (aResults.Length == 0) {
AddMessage("No files found!");
return;
}

sText = Util.GetExtensions(aResults);
sResult = Dialog.Input("Filter", "Extensions", sText, "FilterExtensions").Trim();
if (sResult.Length == 0) return;

aResults = Util.GetPathsWithExtensions(aResults, sResult);
if (aResults.Length == 0) {
AddMessage("No files!");
return;
}

StringBuilder sb = new StringBuilder();
iCount = 0;
AddMessage("Converting");
for (int i = 0; i < aResults.Length; i++) {
string sSource = aResults[i];
string sTarget = Path.ChangeExtension(sSource, ".txt");
string sName = Path.GetFileName(sSource);
AddMessage(sName);
//sText = COM.WordFile2String(sSource);
//sText = COM.ConvertFile2String(sSource);
int iConvert = 2;
string sTargetExt = "txt";
bool bTextOnly = true;
sText = COM.ConvertFile2String(sSource, ref iConvert, ref sTargetExt, bTextOnly);
if (sText.Length == 0) {
AddMessage("Error!");
continue;
}

iCount++;
if (menuItem == menuMiscTextConvert) Util.String2File(sText, sTarget);
else if (iCount == 1) sb.Append(sName + LB + LB + sText);
else sb.Append(SectionBreak + sName + LB + LB + sText);
}

AddMessage("Converted " + Util.Pluralize(iCount, "file"), true);
if (menuItem == menuMiscTextConvert || iCount == 0) return;

if (!IsEmptyWindow()) new MdiChild(this);
sText = sb.ToString();
sText += EOD;
rtb = this.Child.RTB;
rtb.Text = sText;
rtb.Modified = false;
}

if (menuItem == menuMiscTextContents) {
//sMatch = "(^|(" + SB + "))" + "[^\n]*";
sMatch = "(\\A|(" + SB + "))" + "[^\n]*";
aResults = Util.RegExpExtractCase(rtb.Text, sMatch);
iCount = aResults.Length;
AddMessage(Util.Pluralize(iCount, "topic"));
if (iCount == 0) return;

sText = String.Join(LB, aResults);
sText = sText.Replace(SB, "");
sText = "Contents" + LB + LB + sText + SectionBreak;
rtb.ReplaceRange(0, 0, sText);
rtb.Index = 0;
}

if (menuItem == menuMiscSetDefaultFont) {
object[] a = Dialog.GetFont(rtb.Font, rtb.ForeColor);
if (a.Length == 0) return;

rtb.Font = (Font) a[0];
rtb.ForeColor = (Color) a[1];
string sFont = GetFontText(rtb.Font, rtb.ForeColor);
App.WriteOption("FontDefault", sFont);
}

if (menuItem == menuMiscConfigurationOptions) {
aResults = App.ReadDefaultOptions();
//Array.Sort(aResults);
aLabels = new string[aResults.Length];
string[] aDefaults = new string[aResults.Length];
aValues = new string[aResults.Length];
for (int i = 0; i < aResults.Length; i++) {
aLabels[i] = (aResults[i].IndexOf("&") >= 0 ? "" : "&") + aResults[i];
aDefaults[i] = App.ReadDefaultOption(aResults[i], "");
aValues[i] = App.ReadOption(aResults[i], aDefaults[i]);
}

string[] a = Dialog.MultiInput("Configuration Options", aLabels, aValues);
if (a.Length == 0) return;
for (int i = 0; i < a.Length; i++) App.WriteOption(aResults[i], a[i]);
}

if (menuItem == menuMiscManualOptions) {
//OpenOrActivateWindow(App.IniFile, 0);
string sCompiler = App.ReadData("Compiler", "Default");
//sText = sCompiler + " Compiler";
//sResult = Dialog.Choose("Manual Options", "", new string[] {"&Main", "&" + sText}, 0);
sResult = Dialog.Choose("Manual Options", "", new string[] {"&Main", "&" + sCompiler}, 0);
if (sResult.Length == 0) return;

if (sResult == "&Main") sFile = App.IniFile;
else sFile = Path.Combine(App.DataDir, sCompiler + ".ini");
OpenOrActivateWindow(sFile, 0);
}

if (menuItem == menuMiscResetConfiguration) {
/*
if (Dialog.Confirm("Confirm", "Reset Configuration to Default?", "Y") == "Y") {
System.IO.File.Delete(App.IniFile);
App.SetConfigurationValues();
*/

string sCompiler = App.ReadData("Compiler", "Default");
//sText = sCompiler + " Compiler";
//sResult = Dialog.Choose("Manual Options", "", new string[] {"&Main", "&" + sText}, 0);
sResult = Dialog.Choose("Reset Configuration", "", new string[] {"&Main", "&" + sCompiler, "&Both", "&New"}, 0);
if (sResult.Length == 0) return;

if (sResult == "&Main" || sResult == "&Both") {
if (System.IO.File.Exists(App.IniFile)) System.IO.File.Delete(App.IniFile);
System.IO.File.Copy(App.DefaultIniFile, App.IniFile);
}

if (sResult == sCompiler || sResult == "&Both") {
sFile = Path.Combine(App.DataDir, App.ReadData("Compiler", "Default") + ".ini");
if (System.IO.File.Exists(sFile)) System.IO.File.Delete(sFile);
}

if (sResult == "&New") {
aLabels = new string[] {"&Name", "&CompileCommand", "&JumpPosition", "&AbbreviateOutput", "&NavigatePart", "&QuotePrefix", "&ExtensionDefault", "&GoToEnvironment"};
aValues = new string[] {"", "", "", "", "", "", "", ""};
aResults = Dialog.MultiInput("Create Compiler setting", aLabels, aValues);
if (aResults.Length == 0) return;

sCompiler = aResults[0];
HomerList hl = new HomerList(aResults);
hl.RemoveAt(0);
string sSetting = hl.GetSegments('~');
App.WriteValue("Compilers", sCompiler, sSetting);
}
AddMessage("Done");
return;
}

if (menuItem == menuMiscGoToFolder) {
string sDir;
HomerList hl = new HomerList();
aResults = App.ReadSectionKeys("Recent");
foreach (string s in aResults) {
sDir = Path.GetDirectoryName(s);
if (!hl.Contains(sDir)) hl.Add(sDir);
}

aResults = App.ReadSectionKeys("Favorites");
foreach (string s in aResults) {
sDir = Path.GetDirectoryName(s);
if (!hl.Contains(sDir)) hl.Add(sDir);
}

if (hl.Count == 0) {
AddMessage("No items!");
return;
}

aResults = hl.ToArray();
string[] aDisplay = new string[aResults.Length];
for (int i = 0; i < aDisplay.Length; i++) aDisplay[i] = Path.GetFileName(aResults[i]);
sDir = Dialog.Pick("Go to Folder", aResults, aDisplay, true, 0);
if (sDir.Length == 0) return;

Directory.SetCurrentDirectory(sDir);
}

if (menuItem == menuMiscGoToSpecialFolder) {
string sDir = PickSpecialFolder();
if (sDir.Length == 0) return;

Directory.SetCurrentDirectory(sDir);
}

if (menuItem == menuMiscWordWrap) {
rtb.SetWrap(true);
SetRecent(child.File);
}

if (menuItem == menuMiscUnwrap) {
rtb.SetWrap(false);
SetRecent(child.File);
}

if (menuItem == menuMiscPathToClipboard) {
sText = child.File;
Util.SetClipboardText(sText);
AddMessage(sText);
}

if (menuItem == menuMiscPathList) {
sTitle = "Open Folder";
string sDir = Dialog.OpenFolder(sTitle, "Name", Directory.GetCurrentDirectory());
if (sDir.Length == 0) return;

Directory.SetCurrentDirectory(sDir);

sText = Util.GetExtensions(sDir);
if (sText.Length == 0) {
AddMessage("No files!");
return;
}

sResult = Dialog.Input("Filter", "Extensions", sText, "FilterExtensions");
if (sResult.Length == 0) return;

aResults = Util.GetPathsWithExtensions(Directory.GetFiles(sDir), sResult);
iLength = aResults.Length;
sText = Util.Pluralize(iLength, "file");
AddMessage(sText);

if (!IsEmptyWindow()) new MdiChild(this);
child = this.Child;
rtb = child.RTB;
for (int i = 0; i < iLength; i++) {
if (i == 0) sText = aResults[i] + "\n";
else sText += Path.GetFileName(aResults[i]) + "\n";
}
rtb.Text = sText;
rtb.Modified = true;
}

if (menuItem == menuMiscExplorerFolder) {
string sDir = GetDirChoice();
if (sDir.Length == 0) return;
ExplorerFolder(sDir);
}

if (menuItem == menuMiscEvaluateExpression) {
if (rtb.SelectionLength == 0) {
AddMessage("Line");
sText = rtb.RowText;
iIndex = rtb.RowStart + sText.Length;
}
else {
AddMessage("Selected");
sText = rtb.SelectedText;
iIndex = rtb.SelectionStart + sText.Length;
}

sText = Script.run(sText);
if (sText.Length == 0) return;

sText = LB + sText;
rtb.ReplaceRange(iIndex, iIndex, sText);
rtb.Index = iIndex + 1;
Util.Say(rtb.RowText);
}

if (menuItem == menuMiscReplaceTokens) {
if (rtb.SelectionLength == 0) {
//AddMessage("Chunk");
object[] a = GetChunk();
iStart = (int) a[0];
sText = (string) a[1];
iEnd = iStart + sText.Length;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}
sText = rtb.GetRange(iStart, iEnd);

if (rtb.SelectionLength == 0 && !sText.StartsWith("%")) {
AddMessage("Replace Snippet");
aResults = GetSnippetFiles(out aValues);
HomerList hlResults = new HomerList(aResults);
HomerList hlValues = new HomerList(aValues);
//sMatch = @"^" + sText + @".*";
sMatch = sText;
Regex rx = new Regex(sMatch, RegexOptions.IgnoreCase);
iCount = hlResults.Count;
for (int i = iCount - 1; i >= 0; i--) {
if (rx.IsMatch(aValues[i])) continue;
hlResults.RemoveAt(i);
hlValues.RemoveAt(i);
}

if (hlResults.Count == 0) {
AddMessage("No match!");
return;
}

aResults = hlResults.ToArray();
aValues = hlValues.ToArray();

string sSnippet;
if (aResults.Length == 1) sSnippet = aResults[0];
else {
sSnippet = Dialog.Pick("Pick", aResults, aValues, false, 0);
if (sSnippet.Length == 0) return;
}

InvokeSnippet(sSnippet, sText, iStart, iEnd);
}
else {
if (rtb.SelectionLength == 0) AddMessage("Replace Token");
else AddMessage("Replace Selected Tokens");
string sTemp = sText;
sText = ReplaceTokens(sText);
if (sText == sTemp) {
AddMessage("No match!");
return;
}

rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart + sText.Length;
AddMessage(rtb.RowText);
}
}

if (menuItem == menuMiscTransformFiles) {
TransFormFiles();
}

if (menuItem == menuMiscPickCompiler) {
aResults = Ini.ReadSectionKeys(App.IniFile, "Compilers");
//Array.Sort(aResults);
sResult = App.ReadData("Compiler", "Default");
//Dialog.Show(sResult, String.Join("\n", aResults));
int i = Array.IndexOf(aResults, sResult);
//Dialog.Show(i);
if (i == -1) i = 0;
sResult = Dialog.Pick("Pick Compiler", aResults, false, i);
if (sResult.Length == 0) return;

sFile = Path.Combine(App.DataDir, App.ReadData("Compiler", "Default") + ".ini");
string sDir = Directory.GetCurrentDirectory();
Ini.WriteValue(sFile, "Data", "Directory", sDir);
App.WriteData("Compiler", sResult);
sFile = Path.Combine(App.DataDir, sResult + ".ini");
string s = Ini.ReadValue(sFile, "Data", "Directory", "");
if (Directory.Exists(s) && !Util.Equiv(sDir, s)) {
AddMessage("Folder " + Path.GetFileName(s));
Directory.SetCurrentDirectory(s);
}

// A compiler is now defined by a named section, "Compiler <name>",
// with one clearly named key per setting -- CompileCommand,
// JumpPosition, AbbreviateOutput, NavigatePart, QuotePrefix,
// ExtensionDefault, GoToEnvironment -- kept verbatim (and multiline
// when needed) in EdSharp.inix. When such a section exists, it is the
// definition, and the value in the [Compilers] list is free to be a
// human-readable description. The old tilde-packed value is unpacked
// only for entries that have no section, so private legacy compiler
// lines keep working unchanged.
string sSection = "Compiler " + sResult;
string[] aKeys = new string[] {"CompileCommand", "JumpPosition", "AbbreviateOutput", "NavigatePart", "QuotePrefix", "ExtensionDefault", "GoToEnvironment"};
bool bSectionDefined = false;
foreach (string sKey in aKeys) {
if (Ini.ReadValue(App.IniFile, sSection, sKey, "\0") != "\0") { bSectionDefined = true; break; }
}
if (!bSectionDefined) {
sValue = Ini.ReadValue(App.IniFile, "Compilers", sResult, "");
string[] a = sValue.Split('~');
Ini.WriteQuote(App.IniFile, "Options", "CompileCommand", a[0]);
if (a.Length > 1) Ini.WriteQuote(App.IniFile, "Options", "JumpPosition", a[1]);
if (a.Length > 2) Ini.WriteQuote(App.IniFile, "Options", "AbbreviateOutput", a[2]);
if (a.Length > 3) Ini.WriteQuote(App.IniFile, "Options", "NavigatePart", a[3]);
if (a.Length > 4) Ini.WriteQuote(App.IniFile, "Options", "QuotePrefix", a[4]);
if (a.Length > 5) Ini.WriteQuote(App.IniFile, "Options", "ExtensionDefault", a[5]);
if (a.Length > 6) Ini.WriteQuote(App.IniFile, "Options", "GoToEnvironment", a[6]);
}
foreach (string sKey in aKeys) {
string sVal = Ini.ReadValue(App.IniFile, sSection, sKey, "\0");
if (sVal != "\0") Ini.WriteQuote(App.IniFile, "Options", sKey, sVal);
}
}

if (menuItem == menuMiscGoToEnvironment) {
string sCommand = @"%ProgDir%\ijs.exe";
sCommand = App.ReadOption("GoToEnvironment", sCommand);
if (this.Child == null) sFile = "temp.txt";
else sFile = child.File;
if (!sFile.Contains(@"\")) sFile = Path.Combine(Directory.GetCurrentDirectory(), sFile);
sCommand = Util.ExpandCommandLine(sCommand, sFile, sFile);
string sProcess = Path.GetFileNameWithoutExtension(sCommand);
if (!Util.ActivateProcess(sProcess)) Util.Run(sCommand);
}

if (menuItem == menuMiscCompile|| menuItem == menuMiscPromptCommand) {
string sCommand;
string sDefaultJump = "";
string sDefaultAbbreviate = "";
if (menuItem == menuMiscCompile) {
sCommand = App.ReadOption("CompileCommand", "");
// Built-in default: with no compiler configured, compile a C# (.cs) file with
// the latest available .NET Framework C# compiler (Roslyn csc if present, else
// the framework csc that always ships with .NET). This makes Control+F5 work on
// a .cs file out of the box, jumping to the first csc error position.
if (sCommand.Trim().Length == 0 && child.File.ToLower().EndsWith(".cs")) {
string sCsc = Util.FindCscPath();
if (sCsc.Length > 0) {
// 64-bit output by default, matching the compiler chosen above and the
// standing preference for 64-bit everywhere. The compiler names the
// source file in every message, so that prefix is abbreviated away.
sCommand = "\"" + sCsc + "\" /nologo /platform:x64 \"%SourceLong%\" 2>&1";
sDefaultJump = @"\(\d+,\d+\)";
sDefaultAbbreviate = @"^.*?\.cs";
}
}
// The same courtesy for Python: with no compiler configured, a .py or
// .pyw file runs with the python on PATH (the installer's optional
// Python task installs the latest official build there), jumping to
// the traceback's line number. Picking the Python compiler with
// Control+Shift+F5 still overrides this default.
if (sCommand.Trim().Length == 0 && (child.File.ToLower().EndsWith(".py") || child.File.ToLower().EndsWith(".pyw"))) {
// The official python.org build, by its own path when it can be found.
// Windows puts a stub named python on the path that only advertises the
// Microsoft Store; running a program through it produces an
// advertisement rather than an error, which is a baffling thing to hear
// when you expected your program to run.
string sPythonExe = Util.FindPythonPath();
sCommand = ((sPythonExe.Length > 0) ? "\"" + sPythonExe + "\"" : "python") + " \"%SourceLong%\" 2>&1";
sDefaultJump = @"line \d+";
// Python names the file in every traceback frame -- File "C:\long\path
// \script.py", line 4 -- and hearing your own path read out before the
// error wastes the moment that matters. Drop the file prefix and the
// traceback banner, so speech starts at "line 4" and reaches the
// message itself immediately.
sDefaultAbbreviate = @"(^[ \t]*File "".*?"", )|(^Traceback \(most recent call last\):[ \t]*\r?\n)";
}
// And for JavaScript: with no compiler configured, a .js, .mjs, or .cjs
// file runs with the node on PATH (the installer's optional Node task
// installs the 64-bit LTS build there). Node reports a syntax error as
// path:line above the source echo, and a runtime error in stack frames
// as path:line:column, so the jump pattern reads both; frames inside
// Node's own internals are abbreviated away so they can never win the
// earliest-error comparison.
if (sCommand.Trim().Length == 0 && (child.File.ToLower().EndsWith(".js") || child.File.ToLower().EndsWith(".mjs") || child.File.ToLower().EndsWith(".cjs"))) {
sCommand = "node \"%SourceLong%\" 2>&1";
sDefaultJump = @":\d+(:\d+)?";
sDefaultAbbreviate = @"(^[ \t]*at .*node:(internal|diagnostics_channel).*$)|(file:///.*/)|(^[A-Za-z]:\\.*\\)";
}
if (sCommand.Trim().Length == 0) {
AddMessage("No compiler configured. Press Control+Shift+F5 to pick one.");
return;
}
}
else {
sCommand = App.ReadOption("PromptCommand", "");
sCommand = Dialog.Input("Prompt", "Command", sCommand, "Command").Trim();
if (sCommand.Length == 0) return;
App.WriteOption("PromptCommand", sCommand);
}

// Compiling means compiling what you wrote, so the document goes to
// disk first: no separate Control+S, and never a stale file compiled
// while the fix sits unsaved in the window. A document with no disk
// home yet is offered the Save As dialog rather than refused outright.
sFile = child.File;
if (!sFile.Contains(@"\")) sFile = "";
if (sCommand.IndexOf("%Source") >= 0 && sFile.Length == 0) {
AddMessage("Save first");
string sNewFile = Dialog.SaveFile(child.Text, "");
if (sNewFile.Length == 0) return;
child.File = sNewFile;
child.Text = Path.GetFileName(sNewFile);
child.SaveTextOrRtfFile(sNewFile);
sFile = sNewFile;
}
else if (sFile.Length > 0 && (child.RTB.Modified || !File.Exists(sFile))) child.SaveTextOrRtfFile(sFile);

string sDir = Directory.GetCurrentDirectory();
if (sCommand.IndexOf("%SourceDir%") >=0) Directory.SetCurrentDirectory(Path.GetDirectoryName(sFile));
sCommand = Util.ExpandCommandLine(sCommand, sFile, Path.ChangeExtension(sFile, ".exe"));
// Dialog.Show(sCommand);

// Try with COMSpec
// string sOutput = Util.GetProgramOutput(@"c:\windows\system32\cmd.exe", "/c " + sCommand);

// Debug JAWS script
// if (!sCommand.Trim().EndsWith(">1")) sOutput = File.ReadAllText(App.TempFile);
// Util.Run(sCommand);

string sOutput = "";
// if (sCommand.Trim().EndsWith(">1") || sCommand.Trim().EndsWith("&1")) Util.GetProgramOutput("cmd.exe", "/c " + sCommand);
if (sCommand.Trim().EndsWith(">1") || sCommand.Trim().EndsWith("&1")) sOutput = Util.GetProgramOutput("cmd.exe", "/c " + sCommand);
else {
Util.RunHideWait(sCommand);
sOutput = File.ReadAllText(App.TempFile);
}

// Dialog.Show("output", sOutput.Length);

/*
Dialog.Show(sCommand);
int i = sCommand.IndexOf(".exe");
string sParams = sCommand.Substring(i + 5);
sCommand = sCommand.Substring(0, i + 4);
Dialog.Show(sCommand, sParams);
string sOutput = Util.GetProgramOutput(sCommand, sParams);
Dialog.Show(sOutput);
*/

if (sDir != Directory.GetCurrentDirectory()) Directory.SetCurrentDirectory(sDir);

if (menuItem == menuMiscCompile) {
// Speech first, then the cursor. The tool's noise -- your own file path
// in every Python traceback frame, the compiler's echo of the source
// file name, Node's frames inside its own internals -- is removed
// before anything else, so what you hear begins with an error rather
// than with your own directory read aloud.
string sAbbreviateOutput = App.ReadOption("AbbreviateOutput", "\r");
if (sDefaultAbbreviate.Length > 0 && (sAbbreviateOutput == "\\r" || sAbbreviateOutput.Trim().Length == 0)) sAbbreviateOutput = sDefaultAbbreviate;
sOutput = Util.RegExpReplaceEquiv(sOutput, sAbbreviateOutput, "\n").Trim();

// Then the EARLIEST error in the file, not the first one the tool
// happened to print. Compilers report in their own order -- C# by
// severity and file, PowerShell innermost first, Node with its stack
// above the location -- and a person working down a file wants the
// topmost problem each time, so every position in the output is read
// and the smallest line and column wins. Fix it, compile again, and the
// next one is waiting: the file is worked through from the top.
string sJumpPosition = App.ReadOption("JumpPosition", "");
if (sJumpPosition.Trim().Length == 0) sJumpPosition = sDefaultJump;
int iEarliestLine = 0, iEarliestColumn = 0, iEarliestAt = -1;
if (sJumpPosition.Trim().Length > 0) {
try {
foreach (Match match in Regex.Matches(sOutput, sJumpPosition)) {
MatchCollection matchNumbers = Regex.Matches(match.Value, @"\d+");
if (matchNumbers.Count == 0) continue;
int iThisLine = 0, iThisColumn = 1;
if (!Int32.TryParse(matchNumbers[0].Value, out iThisLine)) continue;
if (matchNumbers.Count > 1) Int32.TryParse(matchNumbers[1].Value, out iThisColumn);
if (iThisLine < 1) continue;
if (iEarliestLine == 0 || iThisLine < iEarliestLine || (iThisLine == iEarliestLine && iThisColumn < iEarliestColumn)) {
iEarliestLine = iThisLine;
iEarliestColumn = iThisColumn;
iEarliestAt = match.Index;
}
}
}
catch (Exception) {}
}

if (iEarliestLine > 0) {
// A tool that gives no column marks the spot with carets under an echo
// of the source line instead; read that when the column is still 1.
// An indentation error is about the whitespace at the START of the line,
// whatever the tool's marker points at -- Python puts its caret at the
// end of the line, which is the least useful place to land when the fix
// is at the beginning. The cursor goes to column 1, ready for the edit.
if (sOutput.IndexOf("IndentationError", StringComparison.OrdinalIgnoreCase) >= 0 || sOutput.IndexOf("TabError", StringComparison.OrdinalIgnoreCase) >= 0) iEarliestColumn = 1;
else if (iEarliestColumn <= 1) {
string sDocLine = "";
try { if (iEarliestLine >= 1 && iEarliestLine <= rtb.Lines.Length) sDocLine = rtb.GetRowText(iEarliestLine - 1); }
catch (Exception) {}
int iCaretColumn = Util.CaretMarkerColumn(sOutput, sDocLine);
if (iCaretColumn > 1) iEarliestColumn = iCaretColumn;
}
App.WriteData("Line", iEarliestLine + ", " + iEarliestColumn);
try {
rtb.Line = iEarliestLine;
rtb.Column = iEarliestColumn;
}
catch {}
// Speak from the earliest error onward, so the first words are the
// problem to fix. The lines above it stay in the output file, which
// Review Output shows in full.
if (iEarliestAt > 0) {
int iLineStart = sOutput.LastIndexOf('\n', Math.Min(iEarliestAt, sOutput.Length - 1));
if (iLineStart >= 0 && iLineStart + 1 < sOutput.Length) sOutput = sOutput.Substring(iLineStart + 1);
}
}

if (sOutput.Trim().Length == 0) sOutput = "Done";
AddMessage(sOutput);
}
Util.String2File(sOutput, App.TempFile);
}

if (menuItem == menuMiscReviewOutput) {
sFile = App.TempFile;
if (!File.Exists(sFile)) {
AddMessage("No output file found!");
return;
}
OpenOrActivateWindow(sFile, 0);
}

if (menuItem == menuMiscSaveSnippet) {
if (rtb.SelectionLength == 0) {
AddMessage("All");
iStart = 0;
iEnd = rtb.TextLength;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}
sText = rtb.GetRange(iStart, iEnd);

string sDir = @"Snippets\" + App.ReadData("Compiler", "Default");
sDir = Path.Combine(App.DataDir, sDir);
if (!Directory.Exists(sDir)) Directory.CreateDirectory(sDir);
sFile = Path.Combine(sDir, Path.GetFileName(child.File));
//sFile = Path.ChangeExtension(sFile, ".txt");
if (Path.GetExtension(sFile).Length == 0) sFile += ".txt";
sFile = Dialog.SaveFile("", sFile);
if (sFile.Length == 0) return;

if (rtb.SelectionLength == 0) child.SaveTextOrRtfFile(sFile);
else Util.String2File(sText, sFile);
AddMessage("Done");
}

if (menuItem == menuMiscInvokeSnippet) {
if (rtb.SelectionLength == 0) {
AddMessage("Cursor");
//iStart = 0;
//iEnd = rtb.TextLength;
iStart = rtb.Index;
iEnd = iStart;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}
sText = rtb.GetRange(iStart, iEnd);

aResults = GetSnippetFiles(out aValues);

/*
foreach (string sSnippetFile in aResults) {
Util.Say(Path.GetFileNameWithoutExtension(sSnippetFile));
string sSnippetText = Util.File2String(sSnippetFile);
string[] aSnippetLines = sSnippetText.Split('\n');
StringBuilder sbSnippet = new StringBuilder();
foreach (string sSnippetLine in aSnippetLines) {
if (sSnippetLine.Trim().Length == 0) continue;
sbSnippet.Append(sSnippetLine.Trim() + "\r\n");
}
Util.String2File(sbSnippet.ToString(), sSnippetFile);
}
*/

if (aResults.Length == 0) {
AddMessage("No files!");
return;
}

string sSnippet = Dialog.Pick("Pick", aResults, aValues, false, 0);
if (sSnippet.Length == 0) return;

InvokeSnippet(sSnippet, sText, iStart, iEnd);
}

if (menuItem == menuMiscViewSnippet) {
aResults = GetSnippetFiles(out aValues);
if (aResults.Length == 0) {
AddMessage("No files!");
return;
}

sResult = Dialog.Pick("Pick", aResults, aValues, false, 0);
if (sResult.Length == 0) return;

OpenOrActivateWindow(sResult, 0);
}

if (menuItem == menuMiscKeepUniqueItems) {
string sLimitItem = Util.Literalize(App.ReadOption("LimitItem", "\n"));
if (rtb.SelectionLength == 0) {
AddMessage("All");
iStart = 0;
iEnd = rtb.TextLength;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sText = rtb.GetRange(iStart, iEnd);
aResults = Regex.Split(sText, sLimitItem);
List<string> listNormal = new List<string>();
List<string> listLower = new List<string>();
foreach (string s in aResults) {
string sLower = s.ToLower();
if (listLower.Contains(sLower)) continue;
listLower.Add(sLower);
listNormal.Add(s);
}

aResults = listNormal.ToArray();
sText = String.Join(sLimitItem, aResults);
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
}

if (menuItem == menuMiscNumberItems) {
string sLimitItem = Util.Literalize(App.ReadOption("LimitItem", "\n"));
if (rtb.SelectionLength == 0) {
sTitle = "Number Items All";
iStart = 0;
iEnd = rtb.TextLength;
}
else {
sTitle = "Number Items Selected";
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sResult = Dialog.Input(sTitle, "Start", "1").Trim();
if (sResult.Length == 0) return;

try {
iLine = Int32.Parse(sResult);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}

sText = rtb.GetRange(iStart, iEnd);
aResults = Regex.Split(sText, sLimitItem);
for (int i = 0; i < aResults.Length; i++) {
string s = aResults[i];
// if (s.Trim().Length > 0) s = iLine++ + ". " + s;
s = iLine++ + ". " + s;
aResults[i] = s;
}

sText = String.Join(sLimitItem, aResults);
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
}

if (menuItem == menuMiscOrderItems) {
string sLimitItem = Util.Literalize(App.ReadOption("LimitItem", "\n"));
if (rtb.SelectionLength == 0) {
AddMessage("All");
iStart = 0;
iEnd = rtb.TextLength;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sText = rtb.GetRange(iStart, iEnd);
// aResults = sText.Split('\n');
aResults = Regex.Split(sText, sLimitItem);
string[] a = new string[aResults.Length];
for (int i = 0; i < a.Length; i++) a[i] = aResults[i].ToLower();
Array.Sort(a, aResults);
// sText = String.Join("\n", aResults);
sText = String.Join(sLimitItem, aResults);
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
}

if (menuItem == menuMiscReverseItems) {
string sLimitItem = Util.Literalize(App.ReadOption("LimitItem", "\n"));
if (rtb.SelectionLength == 0) {
AddMessage("All");
iStart = 0;
iEnd = rtb.TextLength;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

sText = rtb.GetRange(iStart, iEnd);
aResults = Regex.Split(sText, sLimitItem);
Array.Reverse(aResults);
sText = String.Join(sLimitItem, aResults);
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
}

if (menuItem == menuMiscListDifferentItems) {
string sLimitItem = Util.Literalize(App.ReadOption("LimitItem", "\n"));
// aResults = rtb.GetRange(0, rtb.RowStart).Split('\n');
sText = rtb.GetRange(0, rtb.RowStart);
aResults = Regex.Split(sText, sLimitItem);
sText = rtb.GetRange(rtb.RowStart, rtb.TextLength);
string[] a = Regex.Split(sText, sLimitItem);
// string[] a = rtb.GetRange(rtb.RowStart, rtb.TextLength).Split('\n');
List<string> list = new List<string>();
foreach (string s in aResults) if (s.Trim().Length > 0 && Array.IndexOf(a, s) == -1) list.Add(s);
aResults = list.ToArray();
if (aResults.Length == 0) {
AddMessage("No output!");
return;
}

AddMessage(Util.Pluralize(aResults.Length, "line"));
// sText = String.Join("\n", aResults).TrimEnd('\n') + "\n";
sText = String.Join(sLimitItem, aResults);
child = new MdiChild(App.Frame);
Child.RTB.Text = sText;
rtb.Index = 0;
}

if (menuItem == menuMiscQueryCommonItems) {
// aResults = rtb.GetRange(0, rtb.RowStart).Split('\n');
// string[] a = rtb.GetRange(rtb.RowStart, rtb.TextLength).Split('\n');
string sLimitItem = Util.Literalize(App.ReadOption("LimitItem", "\n"));
sText = rtb.GetRange(0, rtb.RowStart);
aResults = Regex.Split(sText, sLimitItem);
sText = rtb.GetRange(rtb.RowStart, rtb.TextLength);
string[] a = Regex.Split(sText, sLimitItem);

List<string> list = new List<string>();
foreach (string s in aResults) if (s.Trim().Length > 0 && Array.IndexOf(a, s) >= 0) list.Add(s);
aResults = list.ToArray();
if (aResults.Length == 0) {
AddMessage("No output!");
return;
}

AddMessage(Util.Pluralize(aResults.Length, "line"));
// sText = String.Join("\n", aResults).TrimEnd('\n') + "\n";
sText = String.Join(sLimitItem, aResults);
child = new MdiChild(App.Frame);
Child.RTB.Text = sText;
rtb.Index = 0;
}

if (menuItem == menuMiscCommandPrompt) {
string sDir = GetDirChoice();
if (sDir.Length == 0) return;
CommandPrompt(sDir);
}

if (menuItem == menuMiscBurnToCD) {
BurnToCD();
}

if (menuItem == menuMiscWebDownload) {
string sButton = "Web Page";
if (App.Frame.Child != null) {
sButton = Dialog.Choose("Choose Source of URLs", "", new string[] {"&Web Page", "&Document"}, 0);
if (sButton.Length == 0) return;
} // if

List<string[]> listLinks;
if (sButton.Replace("&", "") == "Web Page") {
string sUrl = COM.GetUrl();
if (sUrl.Length == 0) sUrl = App.ReadData("Url", "");
sUrl = Dialog.Input("Web Download", "Address", sUrl, "WebAddress");
if (sUrl.Length == 0) return;

AddMessage("Please wait");
App.WriteData("Url", sUrl);
listLinks = Homer.Web.getLinks(sUrl);
}
else {
listLinks = new List<string[]>();
aResults = Util.RegExpExtractCase(App.Frame.Child.RTB.Text, @"\w+\:\/\/[^\s""\'\)]+");
if (aResults.Length == 0) {
AddMessage("No URLs found!");
return;
}

for (int i = 0; i < aResults.Length; i++) {
listLinks.Add(new string[] {aResults[i], ""});
} // for
}

List<string> listFiles = new List<string>();
string sRef;
// Name each link the way it will actually be saved (Homer.Web.suggestedName):
// the server-recommended name from Content-Disposition, with an extension
// inferred from the MIME type when the URL carries none.  That makes the
// extension filter below work for links that end in a query rather than a file
// name.  A URL that already ends in a name with an extension costs no request.
Dictionary<string, string> dNames = new Dictionary<string, string>();
foreach (string[] aLink in listLinks) {
sRef = aLink[0];
sFile = Homer.Web.suggestedName(sRef);
if (sFile.Length == 0) sFile = Util.GetFileFromUri(sRef);
dNames[sRef] = sFile;
listFiles.Add(sFile);
}

string[] aFiles = listFiles.ToArray();
sText = Util.GetExtensions(aFiles);
sResult = Dialog.Input("Filter", "Extensions", sText, "DownloadExtensions").Replace(".", "").Trim().ToLower();
if (sResult.Length == 0) return;

aResults = Util.GetPathsWithExtensions(aFiles, sResult);

listFiles.Clear();
List<string> listItems = new List<string>();
List<string> listRefs = new List<string>();
foreach (string[] aLink in listLinks) {
sRef = aLink[0];
sFile = dNames.ContainsKey(sRef) ? dNames[sRef] : Util.GetFileFromUri(sRef);
string sExt = Path.GetExtension(sFile).TrimStart('.').ToLower();
//if (Array.IndexOf(aResults, sExt) == -1) continue;
if (Array.IndexOf(aResults, sFile) == -1) continue;

sText = aLink[1];
if (String.IsNullOrEmpty(sText)) sText = sRef;

listItems.Add(sText + " = " + sFile);
listFiles.Add(sFile);
listRefs.Add(sRef);
}

if (listItems.Count == 0) {
AddMessage("No items!");
return;
}

aValues = listItems.ToArray();
//aResults = Dialog.MultiPick("Pick Files", aValues, new int[] {}, false);
aResults = Dialog.MultiCheck("Pick Files", aValues, new int[] {}, false, 0);
if (aResults.Length == 0) return;

sTitle = "Open Folder";
string sDir = App.ReadData("DownloadFolder", Directory.GetCurrentDirectory());
sDir = Dialog.OpenFolder(sTitle, "Name", sDir);

if (sDir.Length == 0) return;

App.WriteData("DownloadFolder", sDir);
Directory.SetCurrentDirectory(sDir);
AddMessage("Downloading");
foreach (string s in aResults) {
int i = listItems.IndexOf(s);
sFile = listFiles[i];
sRef = listRefs[i];
// Homer.Web.download follows redirects with a real User-Agent and modern TLS,
// takes the file name from the Content-Disposition header when the server
// supplies one (otherwise the link's name plus an extension guessed from the
// content type), and sanitizes and uniquifies the result within sDir.
string sSaved = Homer.Web.download(sRef, sDir, Path.GetFileName(sFile));
if (sSaved.Length > 0) AddMessage(Path.GetFileName(sSaved));
else AddMessage("Could not download " + Path.GetFileName(sFile));
}
AddMessage("Done", true);
}
if (menuItem == menuWindowNext) {
NextWindow();
}

if (menuItem == menuWindowPrior) {
PriorWindow();
}

if (menuItem == menuWindowArrangeIcons) {
this.LayoutMdi(MdiLayout.ArrangeIcons);
return;
}

if (menuItem == menuWindowCascade) {
this.LayoutMdi(MdiLayout.Cascade);
return;
}

if (menuItem == menuWindowTileHorizontal) {
this.LayoutMdi(MdiLayout.TileHorizontal);
return;
}

if (menuItem == menuWindowTileVertical) {
this.LayoutMdi(MdiLayout.TileVertical);
return;
}

if (menuItem == menuHelpAbout) {
sText = "EdSharp 5.0 beta\nJune 16, 2026\n\n";
sText += "Copyright 2007 - 2026 by Jamal Mazrui\nGNU Lesser General Public License (LGPL)\n\n";
sText += ".NET Framework " + RuntimeEnvironment.GetSystemVersion() + "\n\n";
sText += Util.GetPortableExecutableKind();
Dialog.Show("About", sText);
}

if (menuItem == menuHelpDocumentation) {
sFile = Path.Combine(App.ProgramDir, App.ProgramName) + ".htm";
Process.Start(sFile);
}

if (menuItem == menuHelpTutorial) {
// Open the tutorial in the default browser associated with the .htm extension.
sFile = Path.Combine(App.ProgramDir, "Tutorial.htm");
Process.Start(sFile);
}
if (menuItem == menuHelpHistoryOfChanges) {
sFile = Path.Combine(App.ProgramDir, "History.txt");
OpenOrActivateWindow(sFile, 1);
}

if (menuItem == menuHelpSamplePrograms) {
// One menu item, one list. The samples live in a folder rather than in
// menus, snippets or the compiler table, because a sample is something
// you READ once in a while, not something you invoke while working:
// putting them anywhere that grows -- the snippet list a person picks
// from every day, or the menus themselves -- would make the tools you
// use constantly noisier in order to shelve things you open rarely.
//
// The list is built from the folder, so adding a sample is a matter of
// dropping a file in, with no code, no menu entry and no setting. Each
// row reads as its file name followed by the first sentence of what it
// is, taken from the ReadMe beside it when there is one, so the list
// explains itself rather than demanding the ReadMe be opened first.
string sSamplesDir = Path.Combine(App.ProgramDir, "Samples");
if (!Directory.Exists(sSamplesDir)) {
Dialog.Show("Sample Programs", "The Samples folder was not found at:\n" + sSamplesDir);
return;
}
List<object> lPaths = new List<object>();
List<string> lDisplay = new List<string>();
collectSamples(sSamplesDir, "", lPaths, lDisplay);
if (lPaths.Count == 0) {
Dialog.Show("Sample Programs", "No samples were found in:\n" + sSamplesDir);
return;
}
object[] aPicked = Dialog.PickAndChoose("Sample Programs", lPaths.ToArray(), lDisplay.ToArray(), new string[] {"&Open", "&Folder"}, false, 0);
if (aPicked.Length == 0) return;
string sPicked = aPicked[0].ToString();
string sButton = ((string) aPicked[1]).Replace("&", "");
if (sButton == "Folder") {
try { Process.Start("explorer.exe", "/select,\"" + sPicked + "\""); }
catch (Exception ex) { Dialog.Show("Sample Programs", ex.Message); }
return;
}
// Open: a program or a document opens in EdSharp, where it can be read,
// compiled and run; a web page opens in the browser, where it runs.
string sExt = Path.GetExtension(sPicked).ToLower();
if (sExt == ".htm" || sExt == ".html") {
try { Process.Start(sPicked); AddMessage("Opened in the browser"); }
catch (Exception ex) { Dialog.Show("Sample Programs", ex.Message); }
return;
}
OpenOrActivateWindow(sPicked, 0);
}

if (menuItem == menuHelpCopyLog) {
// Put this session's run log path on the clipboard in BOTH formats at
// once, the way HomerView and FileDir's Copy command do: a file drop
// list, so Control+V in a new mail message attaches the log file
// itself, and plain text, so any program that just reads clipboard
// text gets the path. Both matter; a DataObject carries the two
// formats together.
if (App.LogFile == null || App.LogFile.Length == 0 || !File.Exists(App.LogFile)) {
AddMessage("No log for this session!");
return;
}
try {
DataObject dataLog = new DataObject();
System.Collections.Specialized.StringCollection colLogFiles = new System.Collections.Specialized.StringCollection();
colLogFiles.Add(App.LogFile);
dataLog.SetFileDropList(colLogFiles);
dataLog.SetText(App.LogFile);
Clipboard.SetDataObject(dataLog, true);
AddMessage("Log path copied");
}
catch (Exception ex) {
Dialog.Show("Copy Log", ex.Message);
}
}

if (menuItem == menuHelpKeyDescriber) {
if (this.KeyDescriber) {
SetMessage("No Key Describer");
this.KeyDescriber = false;
}
else {
SetMessage("Key Describer On");
this.KeyDescriber = true;
}
}

if (menuItem == menuHelpHotKeySummary) {
sFile = Path.Combine(App.ProgramDir, "HotKeys.txt");
OpenOrActivateWindow(sFile, 1);
}

if (menuItem == menuHelpAlternateMenu) {
AlternateMenu();
}

if (menuItem == menuHelpContextMenu) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

ContextMenu(sFile);
}

if (menuItem == menuHelpSendToMenu) {
sFile = child.File;
if (!sFile.Contains(@"\")) {
AddMessage("No disk file is open for this command!");
return;
}

SendToMenu(sFile);
}

if (menuItem == menuHelpElevateVersion) {
ElevateVersion();
}

} // menuItem_Click handler

// Ask the local Ollama service for one completion and return its text.
// Plain HttpWebRequest against the generate endpoint with streaming
// off; the tiny JSON involved is built and read here directly, so no
// serializer assembly is needed. Failures explain themselves: the
// most common is simply that Ollama is not installed or not running.
// Fetch a model in the background and report the outcome in a plain
// message box -- no console window to find or close. Ollama may have
// been installed moments ago, in which case this process's PATH
// predates it, so the per-user install location is tried when the
// bare name is not found.
// True when an instruction plainly refers to the open document, so the
// document should travel with it. Three signals: verbs that only make
// sense applied to text at hand (summarize, translate, proofread,
// refactor ...); a pointing word attached to a word for the material
// ("this paragraph", "the code", "these functions"); and short bare
// commands such as "summarize" or "explain it". A general-knowledge
// question matches none of them and is sent alone, which is both
// faster and more accurate.
public bool instructionWantsDocument(string sInstruction) {
string sText = sInstruction.ToLowerInvariant().Trim();
string[] aVerbs = new string[] {"summar", "translat", "rewrit", "rephras", "paraphras", "proofread", "outlin", "shorten", "condense", "simplify", "refactor", "critique", "polish", "tighten", "bullet", "keyword", "heading", "docstring", "transcribe"};
foreach (string sVerb in aVerbs) if (sText.Contains(sVerb)) return true;
string sNouns = "(document|text|file|code|selection|passage|article|chapter|section|paragraph|page|snippet|function|method|class|script|essay|draft|list|table|note|email|message|letter)";
if (Regex.IsMatch(sText, @"\b(this|that|these|those|the|my|above|below|following|preceding)\b(\s+\w+){0,2}\s+" + sNouns + @"\b")) return true;
if (Regex.IsMatch(sText, @"^(summarize|translate|proofread|explain|expand|edit|fix|format|review|improve|check)\s*(it|this|that)?\s*[.?!]?$")) return true;
if (Regex.IsMatch(sText, @"\b(explain|describe|analyze|check|improve|review|edit|fix|expand)\b.*\b(this|that|it|above|below)\b")) return true;
return false;
} // instructionWantsDocument method

public void pullOllamaModel(string sModel) {
string sExe = "ollama";
string sUserCopy = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Ollama\ollama.exe");
if (File.Exists(sUserCopy)) sExe = sUserCopy;
AddMessage("Pulling " + sModel);
Util.Log("ollama pull " + sModel + " via " + sExe);
System.Threading.Thread thread = new System.Threading.Thread(delegate() {
string sOutcome;
bool bWorked = false;
try {
ProcessStartInfo psi = new ProcessStartInfo(sExe, "pull " + sModel);
psi.UseShellExecute = false;
psi.CreateNoWindow = true;
psi.RedirectStandardOutput = true;
psi.RedirectStandardError = true;
using (Process process = Process.Start(psi)) {
string sOut = process.StandardOutput.ReadToEnd();
string sErr = process.StandardError.ReadToEnd();
process.WaitForExit();
bWorked = (process.ExitCode == 0);
Util.Log("ollama pull exit code " + process.ExitCode);
if (bWorked) sOutcome = "The " + sModel + " model is ready. Press F12 to chat.";
else {
string sTail = (sErr.Trim().Length > 0 ? sErr : sOut).Trim();
if (sTail.Length > 400) sTail = sTail.Substring(sTail.Length - 400);
sOutcome = "The " + sModel + " model download did not finish.\n" + sTail;
}
}
}
catch (Exception ex) {
sOutcome = "The model download could not start: " + ex.Message + "\nIf Ollama was installed a moment ago, sign out and back in so new programs are on the path, or restart EdSharp.";
}
try {
this.BeginInvoke((MethodInvoker) delegate() {
MessageBox.Show(sOutcome, "Chat with AI");
if (bWorked) AddMessage("Model ready");
});
}
catch (Exception) {}
});
thread.IsBackground = true;
thread.Start();
} // pullOllamaModel method

public string askOllama(string sPrompt, string sModel) {
string sUrl = App.ReadOption("OllamaUrl", "http://localhost:11434").TrimEnd('/') + "/api/generate";
int iTimeoutSeconds = 300;
try { iTimeoutSeconds = Int32.Parse(App.ReadOption("OllamaTimeout", "300")); } catch {}
string sBody = "{\"model\":\"" + jsonEscape(sModel) + "\",\"prompt\":\"" + jsonEscape(sPrompt) + "\",\"stream\":false}";
try {
System.Net.HttpWebRequest request = (System.Net.HttpWebRequest) System.Net.WebRequest.Create(sUrl);
request.Method = "POST";
request.ContentType = "application/json";
request.Timeout = iTimeoutSeconds * 1000;
request.ReadWriteTimeout = iTimeoutSeconds * 1000;
byte[] aBytes = Encoding.UTF8.GetBytes(sBody);
request.ContentLength = aBytes.Length;
using (Stream stream = request.GetRequestStream()) stream.Write(aBytes, 0, aBytes.Length);
string sJson;
using (System.Net.WebResponse response = request.GetResponse())
using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8)) sJson = reader.ReadToEnd();
string sAnswer = jsonExtractString(sJson, "response");
if (sAnswer.Length == 0) {
string sError = jsonExtractString(sJson, "error");
Dialog.Show("Chat with AI", sError.Length > 0 ? "Ollama reported: " + sError : "Ollama sent an empty answer.");
}
return sAnswer;
}
catch (System.Net.WebException ex) {
// A reachable Ollama still answers some requests with an HTTP error --
// asking for a model that has not been pulled returns 404 with a JSON
// body naming the problem -- and HttpWebRequest turns any error status
// into this exception. So read the body first: a real explanation like
// "model \"llama3.2\" not found" beats "(404) Not Found" with install
// hints for a service that is plainly running.
string sDetail = "";
try {
if (ex.Response != null) using (StreamReader reader = new StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8)) sDetail = reader.ReadToEnd();
}
catch {}
string sReported = jsonExtractString(sDetail, "error");
string sHint;
if (sReported.Length > 0) {
sHint = "Ollama reported: " + sReported;
if (sReported.IndexOf("not found", StringComparison.OrdinalIgnoreCase) >= 0) {
// No manual steps: offer to fetch the model right now. The pull runs
// in a visible command window so its progress is readable, and the
// window stays open at the end so the outcome can be reviewed.
if (MessageBox.Show("The " + sModel + " model is not on this machine yet. Fetch it now? It is about 2 gigabytes, shared by every app that uses Ollama.", "Chat with AI", MessageBoxButtons.YesNo) == DialogResult.Yes) {
pullOllamaModel(sModel);
return "";
}
sHint += "\n\nWhen you are ready, fetch it with: ollama pull " + sModel + "\nThe OllamaModel setting picks a different model.";
}
}
else sHint = "Could not reach Ollama at " + sUrl + ".\n" + ex.Message + "\n\nIf Ollama is not installed, rerun the EdSharp installer and check the Ollama box at the finish page, or run: winget install Ollama.Ollama\nIf the model is missing, run: ollama pull " + sModel;
Dialog.Show("Chat with AI", sHint);
return "";
}
catch (Exception ex) {
Dialog.Show("Chat with AI", ex.Message);
return "";
}
} // askOllama method

static string jsonEscape(string sText) {
StringBuilder sbJson = new StringBuilder();
foreach (char c in sText) {
if (c == '"') sbJson.Append("\\\"");
else if (c == '\\') sbJson.Append("\\\\");
else if (c == '\n') sbJson.Append("\\n");
else if (c == '\r') sbJson.Append("\\r");
else if (c == '\t') sbJson.Append("\\t");
else if (c < ' ') sbJson.Append("\\u").Append(((int) c).ToString("x4"));
else sbJson.Append(c);
}
return sbJson.ToString();
} // jsonEscape method

// Read one string field from a flat JSON object: find the key, then
// walk the value honoring escapes. Enough for Ollama's responses
// without a serializer dependency.
static string jsonExtractString(string sJson, string sKey) {
string sMarker = "\"" + sKey + "\":\"";
int iStart = sJson.IndexOf(sMarker);
if (iStart == -1) return "";
iStart += sMarker.Length;
StringBuilder sbValue = new StringBuilder();
int i = iStart;
while (i < sJson.Length) {
char c = sJson[i];
if (c == '\\' && i + 1 < sJson.Length) {
char cNext = sJson[i + 1];
if (cNext == 'n') sbValue.Append('\n');
else if (cNext == 'r') sbValue.Append('\r');
else if (cNext == 't') sbValue.Append('\t');
else if (cNext == '"') sbValue.Append('"');
else if (cNext == '\\') sbValue.Append('\\');
else if (cNext == '/') sbValue.Append('/');
else if (cNext == 'u' && i + 5 < sJson.Length) {
try { sbValue.Append((char) Convert.ToInt32(sJson.Substring(i + 2, 4), 16)); } catch {}
i += 4;
}
i += 2;
continue;
}
if (c == '"') break;
sbValue.Append(c);
i++;
}
return sbValue.ToString();
} // jsonExtractString method

public int runCodeBlocks() {
const string c_sOutputBegins = "<!-- output begins -->";
const string c_sOutputEnds = "<!-- output ends -->";
HomerRichTextBox rtb = this.Child.RTB;
string sDocFile = this.Child.File;
string sDocDir = (sDocFile != null && sDocFile.Contains("\\")) ? Path.GetDirectoryName(sDocFile) : Directory.GetCurrentDirectory();
List<string> lsLines = new List<string>(rtb.Text.Replace("\r\n", "\n").Split('\n'));
int iBlocksRun = 0;
int i = 0;
while (i < lsLines.Count) {
string sTrim = lsLines[i].TrimStart();
bool bFence = sTrim.StartsWith("```") || sTrim.StartsWith("~~~");
if (!bFence) { i++; continue; }
string sFenceMark = sTrim.Substring(0, 3);
string sInfo = sTrim.Substring(3).Trim();
string sLanguage = sInfo.Split(' ')[0].ToLowerInvariant();
if (sLanguage != "sql" && sLanguage != "jscript") { 
// Skip to this block's closing fence so its body cannot start a run.
i++;
while (i < lsLines.Count && !lsLines[i].TrimStart().StartsWith(sFenceMark)) i++;
i++;
continue;
}
string sArgument = (sInfo.Length > sLanguage.Length) ? sInfo.Substring(sLanguage.Length).Trim() : "";
StringBuilder sbBody = new StringBuilder();
int iBodyStart = i + 1;
int iClose = -1;
for (int j = iBodyStart; j < lsLines.Count; j++) {
if (lsLines[j].TrimStart().StartsWith(sFenceMark)) { iClose = j; break; }
sbBody.Append(lsLines[j]).Append("\n");
}
if (iClose == -1) { i = lsLines.Count; continue; }

string sOutput;
if (sLanguage == "sql") sOutput = runSqlBlock(sbBody.ToString(), sArgument, sDocFile, sDocDir);
else sOutput = runJscriptBlock(sbBody.ToString());
iBlocksRun++;

// Replace an existing marked region right after the fence, or insert one.
int iRegionStart = iClose + 1;
int iRegionEnd = -1;
if (iRegionStart < lsLines.Count && lsLines[iRegionStart].Trim() == c_sOutputBegins) {
for (int j = iRegionStart + 1; j < lsLines.Count; j++) {
if (lsLines[j].Trim() == c_sOutputEnds) { iRegionEnd = j; break; }
}
}
if (iRegionEnd >= 0) lsLines.RemoveRange(iRegionStart, iRegionEnd - iRegionStart + 1);
List<string> lsRegion = new List<string>();
lsRegion.Add(c_sOutputBegins);
foreach (string sOutputLine in sOutput.Replace("\r\n", "\n").TrimEnd('\n').Split('\n')) lsRegion.Add(sOutputLine);
lsRegion.Add(c_sOutputEnds);
lsLines.InsertRange(iRegionStart, lsRegion);
i = iRegionStart + lsRegion.Count;
}
if (iBlocksRun > 0) {
int iIndex = rtb.Index;
rtb.Text = String.Join("\n", lsLines.ToArray());
rtb.Modified = true;
if (iIndex <= rtb.TextLength) rtb.Index = iIndex;
}
return iBlocksRun;
} // runCodeBlocks method

string runSqlBlock(string sBody, string sArgument, string sDocFile, string sDocDir) {
string sSqleanFile = Path.Combine(App.ProgramDir, "sqlean.exe");
if (!File.Exists(sSqleanFile)) return "No sqlean.exe was found in the program folder, so the sql block did not run.";
string sDbFile = sArgument.Trim('"');
if (sDbFile.Length == 0) {
if (sDocFile != null && sDocFile.Contains("\\")) sDbFile = Path.ChangeExtension(sDocFile, ".db");
else return "Name the database on the fence line (three backticks, sql, then the path), or save the document so a database with its base name can be assumed.";
}
if (!Path.IsPathRooted(sDbFile)) sDbFile = Path.Combine(sDocDir, sDbFile);
if (!File.Exists(sDbFile)) return "The database was not found: " + sDbFile;
try {
ProcessStartInfo psi = new ProcessStartInfo(sSqleanFile, "-csv -header \"" + sDbFile + "\"");
psi.UseShellExecute = false;
psi.CreateNoWindow = true;
psi.RedirectStandardInput = true;
psi.RedirectStandardOutput = true;
psi.RedirectStandardError = true;
using (Process process = Process.Start(psi)) {
process.StandardInput.Write(sBody);
process.StandardInput.Close();
string sOut = process.StandardOutput.ReadToEnd();
string sErr = process.StandardError.ReadToEnd();
if (!process.WaitForExit(60000)) {
try { process.Kill(); } catch {}
return "The sql block was stopped after 60 seconds.";
}
if (sErr.Trim().Length > 0 && sOut.Trim().Length == 0) return "SQL error: " + sErr.Trim();
if (sOut.Trim().Length == 0) return "0 rows";
Homer.InixTable.TableData table = Homer.InixTable.tableFromDelimitedText(sOut, ',');
string sTable = Homer.InixTable.tableToMarkdown(table);
string sNote = Util.Pluralize(table.Rows.Count, "row");
if (sErr.Trim().Length > 0) sNote += "; SQL messages: " + sErr.Trim();
return sTable + "\n" + sNote;
}
}
catch (Exception ex) {
return "The sql block failed: " + ex.Message;
}
} // runSqlBlock method

string runJscriptBlock(string sBody) {
string sResult = Script.run(sBody);
if (sResult == null || sResult.Length == 0) return "(no result)";
return sResult;
} // runJscriptBlock method

public string checkMarkdown(string sText) {
List<string> lsFindings = new List<string>();
string[] aLines = sText.Replace("\r\n", "\n").Split('\n');
int iPriorHeadingLevel = 0;
Dictionary<string, int> dHeadings = new Dictionary<string, int>();
Dictionary<string, int> dRefsDefined = new Dictionary<string, int>();
Dictionary<string, int> dRefsUsed = new Dictionary<string, int>();
bool bInFence = false;
string sFenceMark = "";
int iFenceLine = 0;
int iTableHeaderCells = 0;
int iTableHeaderLine = 0;
Regex rexHeading = new Regex(@"^(#{1,6})\s+(.*)$");
Regex rexEmptyAlt = new Regex(@"!\[\s*\]\(");
Regex rexBareUrl = new Regex(@"(^|[\s])(https?://[^\s)>\]]+)");
Regex rexRefDefinition = new Regex(@"^\s*\[([^\]^]+)\]:\s*\S");
Regex rexRefUse = new Regex(@"\[[^\]]*\]\[([^\]]+)\]");

for (int i = 0; i < aLines.Length; i++) {
string sLine = aLines[i];
string sTrim = sLine.TrimStart();
int iLineNumber = i + 1;

// Fences first: findings inside a code block would be false alarms.
if (sTrim.StartsWith("```") || sTrim.StartsWith("~~~")) {
string sMark = sTrim.Substring(0, 3);
if (!bInFence) { bInFence = true; sFenceMark = sMark; iFenceLine = iLineNumber; }
else if (sMark == sFenceMark) bInFence = false;
continue;
}
if (bInFence) continue;

// Pipe tables: a row whose cell count differs from its header will
// not convert as intended.
if (sTrim.StartsWith("|")) {
bool bSeparator = true;
foreach (char c in sTrim) if (c != '|' && c != '-' && c != ':' && c != ' ' && c != '\t') { bSeparator = false; break; }
int iCells = sTrim.Trim('|').Split('|').Length;
if (iTableHeaderCells == 0) { iTableHeaderCells = iCells; iTableHeaderLine = iLineNumber; }
else if (!bSeparator && iCells != iTableHeaderCells) {
lsFindings.Add("Line " + iLineNumber + ": table row has " + Util.Pluralize(iCells, "cell") + " but the header on line " + iTableHeaderLine + " has " + iTableHeaderCells + ".");
}
}
else iTableHeaderCells = 0;

Match matchHeading = rexHeading.Match(sLine);
if (matchHeading.Success) {
int iLevel = matchHeading.Groups[1].Value.Length;
if (iPriorHeadingLevel > 0 && iLevel > iPriorHeadingLevel + 1) {
lsFindings.Add("Line " + iLineNumber + ": heading level jumps from " + iPriorHeadingLevel + " to " + iLevel + "; screen reader heading navigation works best when levels step by one.");
}
iPriorHeadingLevel = iLevel;
string sHeadingText = matchHeading.Groups[2].Value.Trim().ToLowerInvariant();
if (dHeadings.ContainsKey(sHeadingText)) {
lsFindings.Add("Line " + iLineNumber + ": duplicate heading; the same text is a heading on line " + dHeadings[sHeadingText] + ", which makes links to it ambiguous.");
}
else dHeadings[sHeadingText] = iLineNumber;
}

if (rexEmptyAlt.IsMatch(sLine)) {
lsFindings.Add("Line " + iLineNumber + ": image with no alt text; a screen reader has nothing to announce for it.");
}

foreach (Match matchUrl in rexBareUrl.Matches(sLine)) {
int iAt = matchUrl.Groups[2].Index;
char cBefore = (iAt > 0) ? sLine[iAt - 1] : ' ';
if (cBefore == '(' || cBefore == '<' || cBefore == '"') continue;
lsFindings.Add("Line " + iLineNumber + ": bare web address; reader-friendly link text with the address behind it reads far better aloud.");
}

Match matchDefinition = rexRefDefinition.Match(sLine);
if (matchDefinition.Success) {
string sRef = matchDefinition.Groups[1].Value.Trim().ToLowerInvariant();
if (!dRefsDefined.ContainsKey(sRef)) dRefsDefined[sRef] = iLineNumber;
}
foreach (Match matchUse in rexRefUse.Matches(sLine)) {
string sRef = matchUse.Groups[1].Value.Trim().ToLowerInvariant();
if (!dRefsUsed.ContainsKey(sRef)) dRefsUsed[sRef] = iLineNumber;
}
}

if (bInFence) lsFindings.Add("Line " + iFenceLine + ": code fence is never closed, so everything after it becomes part of the code block.");
foreach (KeyValuePair<string, int> pairUsed in dRefsUsed) {
if (!dRefsDefined.ContainsKey(pairUsed.Key)) lsFindings.Add("Line " + pairUsed.Value + ": reference link [" + pairUsed.Key + "] is used but never defined, so it will not become a link.");
}
foreach (KeyValuePair<string, int> pairDefined in dRefsDefined) {
if (!dRefsUsed.ContainsKey(pairDefined.Key)) lsFindings.Add("Line " + pairDefined.Value + ": reference link [" + pairDefined.Key + "] is defined but never used.");
}

StringBuilder sbReport = new StringBuilder();
sbReport.Append("Check Markdown: " + Util.Pluralize(lsFindings.Count, "finding") + "\r\n\r\n");
foreach (string sFinding in lsFindings) sbReport.Append(sFinding + "\r\n");
if (lsFindings.Count == 0) sbReport.Append("No problems found by the rules: heading level jumps, missing image alt text, bare web addresses, duplicate headings, unclosed code fences, uneven table rows, and undefined or unused reference links.\r\n");
return sbReport.ToString();
} // checkMarkdown method

object[] GetChunk() {
bool bLoop;
int iIndex, iStart, iEnd;
string c, sText;
HomerRichTextBox rtb = this.Child.RTB;
sText = rtb.Text;
iIndex = rtb.Index;
c = "";

bLoop = true;
while (bLoop) {
if (iIndex == sText.Length) c = "";
else c = sText.Substring(iIndex, 1);
bLoop = (c.Trim().Length == 0);
bLoop = (bLoop && iIndex > 0);
if (bLoop) iIndex--;
}

bLoop = iIndex < sText.Length;
while (bLoop) {
c = sText.Substring(iIndex, 1);
bLoop = (c.Trim().Length > 0);
bLoop = (bLoop && iIndex > 0);
if (bLoop) iIndex--;
}
if (c.Trim().Length == 0) iIndex++;
iStart = iIndex;

bLoop = iIndex < sText.Length;
while (bLoop) {
c = sText.Substring(iIndex, 1);
bLoop = (c.Trim().Length > 0);
iIndex++;
bLoop = (bLoop && iIndex < sText.Length);
}
iEnd = iIndex;

if (iStart == iEnd) sText = "";
else sText = rtb.GetRange(iStart, iEnd);
sText = sText.TrimEnd();
return new object[] {iStart, sText};
} // GetChunk method

public void InvokeSnippet(string sSnippet, string sText, int iStart, int iEnd) {
string[] aLabels, aValues, aResults;
int iIndex;
HomerRichTextBox rtb = this.Child.RTB;
string sLabel, sValue, sMatch;
string sExt = Path.GetExtension(sSnippet).ToLower().TrimStart('.');
string sBody = Util.File2String(sSnippet);
sBody = Util.Convert2UnixLineBreak(sBody);
if (sExt == "js") {
Script.run(sBody);
return;
}
else if (sExt == "boo") {
if (App.Boo == null) App.Boo = COM.CreateObject("Iron.COM");
sSnippet = (string) COM.CallMethod(App.Boo, "Eval", new string[] {sBody, "", "", "", ""});
//Dialog.Show(sSnippet);
return;
}

aResults = sBody.Split('\n');

string sPre = "";
string sPost = "";

string sKeywords = aResults[0];
string[] aKeywords = sKeywords.Split(' ');
string sType = aKeywords[0];
if (sType == "text") {
sText = sBody.Substring(sKeywords.Length + 1);
iIndex = rtb.Index;
}
else if (sType == "html") {
List<string> listResults = new List<string>(aResults);
for (int i = listResults.Count - 1; i > 0; i--) if (listResults[i].Trim().Length == 0 || listResults[i].StartsWith(";")) listResults.RemoveAt(i);
aResults = listResults.ToArray();
sPre = "<" + aResults[1];
if (Array.IndexOf(aResults, "empty") == -1) sPost = "</" + sPre.Substring(1) + ">";

if (aResults.Length > 2) {
List<string> listLabels = new List<string>();
List<string> listValues = new List<string>();
for (int i = 2; i < aResults.Length; i++) {
string sLine = aResults[i];
//if (sLine.Trim().Length == 0 || sLine.StartsWith(";")) continue;
string[] a = sLine.Split('=');
sLabel = a[0];
sValue = "";
if (a.Length > 1) sValue = a[1];
listLabels.Add("&" + sLabel);
listValues.Add(sValue);
}

aLabels = listLabels.ToArray();
aValues = listValues.ToArray();
aResults = Dialog.MultiInput("Attributes", aLabels, aValues);
if (aResults.Length == 0) return;

for (int i = 0; i < aResults.Length; i++) {
sLabel = aLabels[i].Substring(1);
sValue = aResults[i];
sValue = Util.Literalize(sValue);
if (sValue.Length == 0) continue;
sPre += " " + sLabel + "=\"" + sValue + "\"";
}
}

sPre += ">";
if (Array.IndexOf(aKeywords, "phrase") == -1) sPost += "\n";
sText = sPre + sText + sPost;
}
else {
sText = sBody;
iIndex = iStart + sText.Length;
}

if (Array.IndexOf(aKeywords, "form") >= 0) {
string sDate, sTime;
GetDateAndTime(out sDate, out sTime);
sText = sText.Replace("%Date%", sDate);
sText = sText.Replace("%Time%", sTime);
string sUserName = Environment.UserName;
sUserName = sUserName.Replace(".", " ");
sText = sText.Replace("%UserName%", sUserName);
string[] aNames = (sUserName + " ").Split(' ');
sText = sText.Replace("%UserFirstName%", aNames[0]);
sText = sText.Replace("%UserLastName%", aNames[1]);

sMatch = @"\%\w+\=.*?\%";
string[] aVars = Util.RegExpExtractCase(sText, sMatch);
if (aVars.Length > 0) {
List<string> listLabels = new List<string>();
List<string> listValues = new List<string>();
List<string> listVars = new List<string>(aVars);
foreach (string sVar in aVars) {
string[] aParts = sVar.Split('=');
sLabel = aParts[0];
sLabel = "&" + sLabel.Substring(1, sLabel.Length - 1);
sValue = aParts[1];
sValue = sValue.Substring(0, sValue.Length - 1);
if (listLabels.Contains(sLabel)) {
// Stop reverse bug
listVars.Reverse();
listVars.Remove(sVar);
listVars.Reverse();
continue;
}

listLabels.Add(sLabel);
listValues.Add(sValue);
}
aLabels = listLabels.ToArray();
aValues = listValues.ToArray();
aResults = Dialog.MultiInput("Variables", aLabels, aValues);
if (aResults.Length == 0) return;

aVars = listVars.ToArray();
for (int i = 0; i < aVars.Length; i++) {
string sVar = aVars[i];
// Dialog.Show("sVar=" + sVar, "result=" + aResults[i]);
sText = sText.Replace(sVar, aResults[i]);
sVar = sVar.Split('=')[0] + "=%";
sText = sText.Replace(sVar, aResults[i]);
}
}

sText = ReplaceTokens(sText);
}

if (Array.IndexOf(aKeywords, "caret") >= 0) {
int i = sText.IndexOf("^^");
if (i >= 0) {
iIndex = iStart + i;
sText = sText.Remove(i, 2);
}
else iIndex = iStart;
}
else {
iIndex = iStart + sText.Length;
if (rtb.SelectionLength == 0) iIndex -= sPost.Length;
}

rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iIndex;
Util.Say(rtb.RowText);
} // InvokeSnippet method

// ===== Python structure engine (rewritten 24 August 2026) ==================
// PyBrace and PyDent convert between Python's indentation and a flat,
// brace-marked form that is far friendlier under a screen reader: block
// structure is spoken as visible text ("... {" and "} end def") instead
// of counted spaces. The previous engine corrupted three common things:
// dictionary and set literals (any line ending in a brace was taken for
// a block), docstrings and other triple-quoted strings (their interior
// indentation was flattened), and braces inside strings such as
// f-strings. This engine tracks string and bracket state character by
// character, so only real compound statements (def, class, if, elif,
// else, for, while, try, except, finally, with, match, case, and their
// async forms) open blocks, only "} end word" lines close them, string
// interiors pass through verbatim, and continuation lines keep an extra
// indent unit. Round trips are exact: PyDent(PyBrace(code)) reproduces
// the original, apart from normalized indentation and the spoken
// "# end" markers PyDent adds at block ends.

static readonly string[] c_aPythonBlockWords = new string[] {"def", "class", "if", "elif", "else", "for", "while", "try", "except", "finally", "with", "match", "case"};

static bool isPythonBlockWord(string sWord) {
foreach (string s in c_aPythonBlockWords) if (s == sWord) return true;
return false;
} // isPythonBlockWord method

// Walk one raw line, carrying triple-quote state across lines. Returns
// the line with string interiors blanked and any comment removed, and
// reports the bracket-depth change from this line's real code.
static string scanPythonLine(string sLine, ref bool bInTriple, ref char cTriple, out int iDepthDelta) {
StringBuilder sbCode = new StringBuilder();
int iDepth = 0;
bool bInSingle = false;
char cSingle = ' ';
int i = 0;
while (i < sLine.Length) {
char c = sLine[i];
if (bInTriple) {
if (c == cTriple && i + 2 < sLine.Length + 1 && sLine.Length - i >= 3 && sLine[i + 1] == cTriple && sLine[i + 2] == cTriple) {
bInTriple = false; i += 3; sbCode.Append("   "); continue;
}
i++; sbCode.Append(' '); continue;
}
if (bInSingle) {
if (c == '\\') { i += 2; sbCode.Append("  "); continue; }
if (c == cSingle) bInSingle = false;
i++; sbCode.Append(' '); continue;
}
if (c == '"' || c == '\'') {
if (sLine.Length - i >= 3 && sLine[i + 1] == c && sLine[i + 2] == c) {
bInTriple = true; cTriple = c; i += 3; sbCode.Append("   "); continue;
}
bInSingle = true; cSingle = c; i++; sbCode.Append(' '); continue;
}
if (c == '#') break;
if (c == '(' || c == '[' || c == '{') iDepth++;
else if (c == ')' || c == ']' || c == '}') iDepth--;
sbCode.Append(c); i++;
}
iDepthDelta = iDepth;
return sbCode.ToString();
} // scanPythonLine method

static string pythonFirstWord(string sCode) {
string sTrim = sCode.Trim();
if (sTrim.StartsWith("async ")) sTrim = sTrim.Substring(6).TrimStart();
StringBuilder sbWord = new StringBuilder();
foreach (char c in sTrim) {
if (Char.IsLetterOrDigit(c) || c == '_') sbWord.Append(c);
else break;
}
return sbWord.ToString();
} // pythonFirstWord method

// The indent unit of a Python document: the leading whitespace of its
// first indented code line (uniform characters only), else the given
// default. This makes PyBrace respect the file's own style.
static string detectPythonUnit(string sText, string sDefault) {
bool bInTriple = false;
char cTriple = ' ';
int iDelta;
foreach (string sRaw in sText.Replace("\r\n", "\n").Split('\n')) {
bool bWasTriple = bInTriple;
scanPythonLine(sRaw, ref bInTriple, ref cTriple, out iDelta);
if (bWasTriple) continue;
string sTrim = sRaw.TrimStart();
if (sTrim.Length == 0 || sTrim.StartsWith("#")) continue;
string sWs = sRaw.Substring(0, sRaw.Length - sTrim.Length);
if (sWs.Length == 0) continue;
char cFirst = sWs[0];
foreach (char c in sWs) if (c != cFirst) return sDefault;
return sWs;
}
return sDefault;
} // detectPythonUnit method

static int pythonIndentLevels(string sWs, string sUnit) {
int iLevels = 0;
while (sUnit.Length > 0 && sWs.StartsWith(sUnit)) { iLevels++; sWs = sWs.Substring(sUnit.Length); }
return iLevels;
} // pythonIndentLevels method

// True for a marker line PyDent wrote: "# end" followed by one word.
static bool isEndMarkerLine(string sTrim) {
if (!sTrim.StartsWith("# end ")) return false;
string sWord = sTrim.Substring(6).Trim();
if (sWord.Length == 0) return false;
foreach (char c in sWord) if (!Char.IsLetterOrDigit(c) && c != '_') return false;
return true;
} // isEndMarkerLine method

public string PyDent2Brace(string sText) {
string sUnitDefault = Util.Literalize(App.ReadOption("IndentUnit", "\t"));
if (sUnitDefault.Trim('\t', ' ').Length > 0 || sUnitDefault.Length == 0) sUnitDefault = "\t";
string sUnit = detectPythonUnit(sText, sUnitDefault);
List<string> lsOut = new List<string>();
List<string> lsBlanks = new List<string>();
List<int> lsStackLevel = new List<int>();
List<string> lsStackWord = new List<string>();
bool bInTriple = false;
char cTriple = ' ';
int iContDepth = 0;
bool bBackslash = false;
foreach (string sRaw in sText.Replace("\r\n", "\n").Split('\n')) {
string sTrim = sRaw.Trim();
if (isEndMarkerLine(sTrim)) continue;
bool bWasTriple = bInTriple;
bool bWasCont = iContDepth > 0 || bBackslash;
int iDelta;
string sScanned = scanPythonLine(sRaw, ref bInTriple, ref cTriple, out iDelta);
if (bWasTriple) {
// Triple-quoted interiors keep their exact text: indentation there is content.
lsOut.AddRange(lsBlanks); lsBlanks.Clear();
lsOut.Add(sRaw);
iContDepth = Math.Max(0, iContDepth + iDelta);
bBackslash = sScanned.TrimEnd().EndsWith("\\");
continue;
}
if (bWasCont) {
lsOut.AddRange(lsBlanks); lsBlanks.Clear();
lsOut.Add(sTrim);
iContDepth = Math.Max(0, iContDepth + iDelta);
bBackslash = sScanned.TrimEnd().EndsWith("\\");
continue;
}
if (sTrim.Length == 0) { lsBlanks.Add(""); continue; }
string sWs = sRaw.Substring(0, sRaw.Length - sRaw.TrimStart().Length);
int iLevel = pythonIndentLevels(sWs, sUnit);
// A line at a shallower indent closes the blocks it leaves; comments
// close blocks by their own indent but never open one. Blank lines
// are held back so closers land right after the block's last line.
while (lsStackLevel.Count > 0 && lsStackLevel[lsStackLevel.Count - 1] >= iLevel) {
lsOut.Add("} end " + lsStackWord[lsStackWord.Count - 1]);
lsStackLevel.RemoveAt(lsStackLevel.Count - 1);
lsStackWord.RemoveAt(lsStackWord.Count - 1);
}
lsOut.AddRange(lsBlanks); lsBlanks.Clear();
string sCodeTrimmed = sScanned.TrimEnd();
if (!sTrim.StartsWith("#") && sCodeTrimmed.EndsWith(":") && iContDepth + iDelta == 0 && !bInTriple) {
int iColon = sCodeTrimmed.Length - 1 - sWs.Length;
string sBody = (iColon >= 0 && iColon <= sTrim.Length) ? sTrim.Substring(0, iColon).TrimEnd() : sTrim.TrimEnd(':', ' ');
lsOut.Add(sBody + " {");
string sWord = pythonFirstWord(sScanned);
lsStackLevel.Add(iLevel);
lsStackWord.Add(sWord.Length > 0 ? sWord : "block");
}
else lsOut.Add(sTrim);
iContDepth = Math.Max(0, iContDepth + iDelta);
bBackslash = sScanned.TrimEnd().EndsWith("\\") && !bInTriple;
}
while (lsStackLevel.Count > 0) {
lsOut.Add("} end " + lsStackWord[lsStackWord.Count - 1]);
lsStackLevel.RemoveAt(lsStackLevel.Count - 1);
lsStackWord.RemoveAt(lsStackWord.Count - 1);
}
lsOut.AddRange(lsBlanks);
return String.Join("\n", lsOut.ToArray()).Trim('\n') + "\n";
} // PyDent2Brace method

public string PyBrace2Dent(string sText) {
string sUnit = Util.Literalize(App.ReadOption("IndentUnit", "\t"));
if (sUnit.Trim('\t', ' ').Length > 0 || sUnit.Length == 0) sUnit = "\t";
List<string> lsOut = new List<string>();
List<string> lsStack = new List<string>();
bool bInTriple = false;
char cTriple = ' ';
int iContDepth = 0;
bool bBackslash = false;
foreach (string sRaw in sText.Replace("\r\n", "\n").Split('\n')) {
string sTrim = sRaw.Trim();
bool bWasTriple = bInTriple;
bool bWasCont = iContDepth > 0 || bBackslash;
if (!bWasTriple && sTrim.StartsWith("} end ")) {
string sWord = sTrim.Substring(6).Trim();
if (lsStack.Count > 0) lsStack.RemoveAt(lsStack.Count - 1);
lsOut.Add(Util.Replicate(sUnit, lsStack.Count) + "# end " + sWord);
continue;
}
// Only a compound-statement keyword line ending in a lone brace opens
// a block; a dictionary literal like "table = {" is ordinary code.
bool bOpener = false;
string sScanSource = sRaw;
if (!bWasTriple && !bWasCont && sTrim.EndsWith("{") && !sTrim.StartsWith("#")) {
if (isPythonBlockWord(pythonFirstWord(sTrim))) {
bOpener = true;
sScanSource = sTrim.Substring(0, sTrim.Length - 1).TrimEnd();
}
}
int iDelta;
string sScanned = scanPythonLine(sScanSource, ref bInTriple, ref cTriple, out iDelta);
if (bWasTriple) lsOut.Add(sRaw);
else if (sTrim.Length == 0) lsOut.Add("");
else if (bOpener) {
lsOut.Add(Util.Replicate(sUnit, lsStack.Count) + sTrim.Substring(0, sTrim.Length - 1).TrimEnd() + ":");
lsStack.Add(pythonFirstWord(sTrim));
}
else {
int iExtra = (bWasCont && sTrim.Length > 0 && sTrim[0] != ')' && sTrim[0] != ']' && sTrim[0] != '}') ? 1 : 0;
lsOut.Add(Util.Replicate(sUnit, lsStack.Count + iExtra) + sTrim);
}
iContDepth = Math.Max(0, iContDepth + iDelta);
bBackslash = sScanned.TrimEnd().EndsWith("\\") && !bInTriple;
}
return String.Join("\n", lsOut.ToArray()).Trim('\n') + "\n";
} // PyBrace2Dent method

public void HardLineBreak() {
bool bLoop;
string sResult, sTitle, sBody, sText, sLine;
int iWidth, iLength, iIndex, iStart, iEnd, i;
HomerRichTextBox rtb = this.Child.RTB;
if (rtb.SelectionLength == 0) {
sTitle = "Hard Line Break All";
iStart = 0;
iEnd = rtb.TextLength;
}
else {
sTitle = "Hard Line Break Selected";
iStart = rtb.SelectionStart;
iEnd = iStart + rtb.SelectionLength;
}

iWidth = 0;
foreach (string s in rtb.Lines) {
iLength = s.Length;
if (iLength > iWidth) iWidth = iLength;
}

sText = iWidth.ToString();
sResult = Dialog.Input(sTitle, "Width", sText).Trim();
if (sResult.Length == 0) return;

try {
iWidth = Int32.Parse(sResult);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}

sText = rtb.GetRange(iStart, iEnd);
sBody = "";
iIndex = 0;
iLength = sText.Length;
bLoop = true;

while (bLoop) {
//Dialog.Show("index", iIndex);
if (iLength - iIndex <= iWidth) {
sLine = sText.Substring(iIndex);
bLoop = false;
}
else {
sLine = sText.Substring(iIndex, iWidth);
i = sLine.LastIndexOf("\n");
//Dialog.Show("break", i);
if (i >=0) {
sLine = sLine.Substring(0, i + 1);
}
else {
i = sLine.LastIndexOf(" ");
//Dialog.Show("space", i);
if (i >=0) sLine = sLine.Substring(0, i + 1);
}
}
sBody += sLine.TrimEnd('\n') + "\n";
iIndex += sLine.Length;
}

rtb.ReplaceRange(iStart, iEnd, sBody);
rtb.Index = iStart + sText.Length;
Util.Say(rtb.RowText);
} // HardLineBreak method

public void CalculateDate() {
DateTime dt = new DateTime();
int iIndex, iResult, iYear, iMonth, iWeek, iDay;
string sText, sTitle, sYear, sMonth, sWeek, sDay;
string[] aResults, aLabels, aValues;
HomerRichTextBox rtb = this.Child.RTB;

sYear = App.ReadData("Year", "");
sMonth = App.ReadData("Month", "");
sWeek = App.ReadData("Week", "");
sDay = App.ReadData("Day", "");
aLabels = new string[] {"&Year", "&Month", "&Week", "&Day"};
aValues = new string[] {sYear, sMonth, sWeek, sDay};
sTitle = "Calculate Date";
aResults = Dialog.MultiInput(sTitle, aLabels, aValues);
if (aResults.Length == 0) return;

sYear = aResults[0].Trim();
sMonth = aResults[1].Trim();
sWeek = aResults[2].Trim();
sDay = aResults[3].Trim();
App.WriteData("Year", sYear);
App.WriteData("Month", sMonth);
App.WriteData("Week", sWeek);
App.WriteData("Day", sDay);

iResult = Util.Month2Num(sMonth);
if (iResult != -1) sMonth = iResult.ToString();
iResult = Util.Day2Num(sDay);
if (iResult != -1) sDay = iResult.ToString();

try {
iYear = (sYear.Length == 0) ? 0 : Int32.Parse(sYear);
iMonth = (sMonth.Length == 0) ? 0 : Int32.Parse(sMonth);
iWeek = (sWeek.Length == 0) ? 0 : Int32.Parse(sWeek);
iDay = (sDay.Length == 0) ? 0 : Int32.Parse(sDay);
if (iWeek == 0) {
dt = new DateTime(iYear, iMonth, iDay);
}
else {
dt = new DateTime(iYear, iMonth, 1);
dt = dt.AddDays(7 * (iWeek - 1));
while ((int) dt.DayOfWeek != iDay) dt = dt.AddDays(1);
}
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
}

iIndex = rtb.Index;
sText = dt.ToLongDateString();
rtb.ReplaceRange(iIndex, iIndex, sText);
iIndex += sText.Length;
rtb.Index = iIndex;
Util.Say(rtb.RowText);
} // CalculateDate method

// Mail a document as the body of a message, or as an attachment, using
// the mail program Windows already knows about -- no Microsoft Word.
//
// For the body, a mailto link carries the subject and text: every mail
// program on Windows registers for mailto, including Outlook, Thunderbird
// and the Mail app, so the message opens in whatever the person already
// uses. Mail links have a practical length limit, so a long document is
// offered as an attachment instead rather than arriving truncated.
//
// For an attachment, the document is written to a temporary file and the
// Windows "mailto with attachment" path is used where the mail program
// supports it; when it does not, EdSharp says so plainly and puts the
// file on the clipboard in both formats, so one paste into a new message
// attaches it -- the same dual-format trick Copy Log uses.
public void Mail(bool bAttach) {
if (App.ReadOption("Mailer", "Windows").Trim().ToLower() == "word") { MailWord(bAttach); return; }
HomerRichTextBox rtb = this.Child.RTB;
string sBody = rtb.Text;
if (sBody.Trim().Length == 0) { AddMessage("No text!"); return; }
string sSubject = Path.GetFileNameWithoutExtension(this.Child.Text);
if (sSubject.Trim().Length == 0) sSubject = "Document";

if (!bAttach) {
// mailto carries about 2000 characters safely across mail programs.
const int c_iBodyLimit = 1800;
if (sBody.Length > c_iBodyLimit) {
if (Dialog.Confirm("Mail", "This document is longer than a mail link can carry (" + Util.Pluralize(sBody.Length, "character") + ").\n\nSend it as an attachment instead?", "Y") != "Y") return;
bAttach = true;
}
else {
string sUrl = "mailto:?subject=" + Uri.EscapeDataString(sSubject) + "&body=" + Uri.EscapeDataString(Util.Convert2WinLineBreak(sBody));
try {
Process.Start(sUrl);
AddMessage("Mail message opened");
Util.Log("mailto opened, " + sBody.Length + " characters");
}
catch (Exception ex) {
Dialog.Show("Mail", "No mail program answered.\n" + ex.Message);
}
return;
}
}

// Attachment: write the document beside the temporary files EdSharp
// already cleans up, then hand the file to the mail program.
string sName = this.Child.Text;
if (Path.GetExtension(sName).Length == 0) sName += ".txt";
string sFile = Path.Combine(Path.GetTempPath(), sName);
try {
if (File.Exists(sFile)) File.Delete(sFile);
Util.String2File(Util.Convert2WinLineBreak(sBody), sFile);
App.TempFiles.Add(sFile);
}
catch (Exception ex) {
Dialog.Show("Mail", "The attachment could not be written.\n" + ex.Message);
return;
}

// Most mail programs accept an attachment on a mailto link; Outlook
// does not, so the clipboard path below is the reliable answer there.
bool bOpened = false;
try {
Process.Start("mailto:?subject=" + Uri.EscapeDataString(sSubject) + "&attach=" + Uri.EscapeDataString(sFile));
bOpened = true;
}
catch (Exception) { bOpened = false; }
try {
DataObject dataFile = new DataObject();
System.Collections.Specialized.StringCollection colFiles = new System.Collections.Specialized.StringCollection();
colFiles.Add(sFile);
dataFile.SetFileDropList(colFiles);
dataFile.SetText(sFile);
Clipboard.SetDataObject(dataFile, true);
}
catch (Exception) {}
Util.Log("mail attachment prepared: " + sFile);
if (bOpened) AddMessage("Mail message opened; the file is also on the clipboard");
else Dialog.Show("Mail", "The document is saved as:\n" + sFile + "\n\nIt is on the clipboard as a file, so pressing Control+V in a new mail message attaches it.");
} // Mail method

public void MailWord(bool bAttach) {
bool bCreate, bVisible, bSendMailAttach;
int iDisplayAlerts;
string sText, sFile, sDir;
object oApp, oOptions, oDocs, oDoc;

HomerRichTextBox rtb = this.Child.RTB;
sText = rtb.Text;

if (sText.Length == 0) {
AddMessage("No text!");
return;
}
sText = Util.Convert2WinLineBreak(sText);
sFile = this.Child.Text;
sDir = Path.GetTempPath();
sFile = Path.Combine(sDir, sFile);
if (Path.GetExtension(sFile).Length == 0) sFile += ".txt";
//Util.String2File(sText, sFile);
Util.String2File(sText, App.TempFile);
App.TempFiles.Add(sFile);

bool bAppVisible = false;
//oApp = COM.GetOrCreateObject("Word.Application", out bCreate);
oApp = COM.WordAccess(out bCreate);
bVisible = (bool) COM.GetProperty(oApp, "Visible");
iDisplayAlerts = (int) COM.GetProperty(oApp, "DisplayAlerts");
COM.SetProperty(oApp, "Visible", bAppVisible);
COM.SetProperty(oApp, "DisplayAlerts", 0);
oOptions = COM.GetProperty(oApp, "Options");
bSendMailAttach = (bool) COM.GetProperty(oOptions, "SendMailAttach");
COM.SetProperty(oOptions, "SendMailAttach", bAttach);
oDocs = COM.GetProperty(oApp, "Documents");
oDoc = VB.WordOpen(oDocs, App.TempFile, bAppVisible);
if (File.Exists(sFile)) File.Delete(sFile);
VB.WordSaveAs(oDoc, sFile, 2);
COM.CallMethod(oDoc, "SendMail");
VB.WordClose(oDoc);
COM.Release(ref oDoc);
COM.Release(ref oDocs);
if (bCreate) {
//VB.WordQuit(oApp);
}
else {
COM.SetProperty(oApp, "Visible", bVisible);
COM.SetProperty(oApp, "DisplayAlerts", iDisplayAlerts);
COM.SetProperty(oOptions, "SendMailAttach", bSendMailAttach);
}
COM.Release(ref oOptions);
COM.Release(ref oApp);
File.Delete(sFile);

App.Frame.Activate();
App.Frame.Child.RTB.Select();
} // MailBody method

// ===== Spell check without Word ============================================
// Windows has carried a spell checking service since Windows 8: the same
// engine that underlines misspellings in built-in edit controls, with
// the user's own added words. It is reached through COM, so no library
// ships with EdSharp, nothing is downloaded, and a machine without
// Microsoft Office spell checks exactly as well as one with it. The
// interfaces are declared here rather than referenced from an assembly,
// keeping EdSharp a single binary.

// The interfaces exactly as Windows declares them in spellcheck.h.
//
// EVERY METHOD IS DECLARED, in order, including the ones EdSharp never
// calls. The previous version skipped them with _VtblGap markers, which
// exist for exactly that purpose -- but the object came back refusing to
// be an ISpellChecker, which is the signature of a call landing on the
// wrong slot: the gaps were not counted as the header counts them.
// Naming every slot removes the guesswork. An unused method needs no
// real signature, because it is never called, only counted.

[ComImport, Guid("8E018A9D-2415-4677-BF08-794EA61F94BB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface ISpellCheckerFactory {
void get_SupportedLanguages();
void IsSupported([MarshalAs(UnmanagedType.LPWStr)] string languageTag, [MarshalAs(UnmanagedType.Bool)] out bool value);
// Returned as a plain object rather than as ISpellChecker: the runtime
// would otherwise demand that interface at the moment of return and
// throw if the answer were no, leaving nothing to examine. Asking
// afterwards lets the code say what it actually received.
[return: MarshalAs(UnmanagedType.IUnknown)] object CreateSpellChecker([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
} // ISpellCheckerFactory interface

[ComImport, Guid("B6FD0B71-E2BC-4653-8D05-F197E412770A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface ISpellChecker {
void get_LanguageTag();
[return: MarshalAs(UnmanagedType.Interface)] IEnumSpellingError Check([MarshalAs(UnmanagedType.LPWStr)] string text);
[return: MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IEnumString Suggest([MarshalAs(UnmanagedType.LPWStr)] string word);
void Add([MarshalAs(UnmanagedType.LPWStr)] string word);
void Ignore([MarshalAs(UnmanagedType.LPWStr)] string word);
void AutoCorrect([MarshalAs(UnmanagedType.LPWStr)] string from, [MarshalAs(UnmanagedType.LPWStr)] string to);
void GetOptionValue();
void get_OptionIds();
void get_Id();
void get_LocalizedName();
void add_SpellCheckerChanged();
void remove_SpellCheckerChanged();
void GetOptionDescription();
[return: MarshalAs(UnmanagedType.Interface)] IEnumSpellingError ComprehensiveCheck([MarshalAs(UnmanagedType.LPWStr)] string text);
} // ISpellChecker interface

// ISpellChecker2 is ISpellChecker with one more method. An object that
// refuses the first identifier may accept this one, so it is worth
// asking before giving up.
[ComImport, Guid("E7ED1C71-87F7-4378-A840-C9200DACEE47"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface ISpellChecker2 {
void get_LanguageTag();
[return: MarshalAs(UnmanagedType.Interface)] IEnumSpellingError Check([MarshalAs(UnmanagedType.LPWStr)] string text);
[return: MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IEnumString Suggest([MarshalAs(UnmanagedType.LPWStr)] string word);
void Add([MarshalAs(UnmanagedType.LPWStr)] string word);
void Ignore([MarshalAs(UnmanagedType.LPWStr)] string word);
void AutoCorrect([MarshalAs(UnmanagedType.LPWStr)] string from, [MarshalAs(UnmanagedType.LPWStr)] string to);
void GetOptionValue();
void get_OptionIds();
void get_Id();
void get_LocalizedName();
void add_SpellCheckerChanged();
void remove_SpellCheckerChanged();
void GetOptionDescription();
[return: MarshalAs(UnmanagedType.Interface)] IEnumSpellingError ComprehensiveCheck([MarshalAs(UnmanagedType.LPWStr)] string text);
void Remove([MarshalAs(UnmanagedType.LPWStr)] string word);
} // ISpellChecker2 interface

[ComImport, Guid("803E3BD4-2828-4410-8290-418D1D73C762"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IEnumSpellingError {
[return: MarshalAs(UnmanagedType.Interface)] ISpellingError Next();
} // IEnumSpellingError interface

[ComImport, Guid("B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface ISpellingError {
uint StartIndex { get; }
uint Length { get; }
uint CorrectiveAction { get; }
// The marshaling attribute goes on the getter, since this is a property;
// on the property itself the compiler ignores it with a warning.
string Replacement { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
} // ISpellingError interface

[ComImport, Guid("7AB36653-1796-484B-BDFA-E74F1DB7C1DC")]
class SpellCheckerFactoryClass {
} // SpellCheckerFactoryClass

// One misspelling found in the document: where it starts, how long it
// is, the word itself, and what the system suggests instead.
public class SpellingProblem {
public int Start;
public int Length;
public string Word;
public List<string> Suggestions;
} // SpellingProblem class

// Ask Windows to check a stretch of text. Returns the problems in
// document order, each with its suggestions. Returns null when the
// service is unavailable, so the caller can say so plainly.
// What a COM object will admit to being. Each identifier is offered in
// turn and the ones it accepts are named, which turns "no such
// interface" from a dead end into a description of what came back.
public string describeInterfaces(object oObject) {
string[,] aKnown = new string[,] {
{"B6FD0B71-E2BC-4653-8D05-F197E412770A", "ISpellChecker"},
{"E7ED1C71-87F7-4378-A840-C9200DACEE47", "ISpellChecker2"},
{"8E018A9D-2415-4677-BF08-794EA61F94BB", "ISpellCheckerFactory"},
{"AA176B85-0E12-4844-8E1A-EEF1DA77F586", "IUserDictionariesRegistrar"},
{"00000101-0000-0000-C000-000000000046", "IEnumString"},
{"803E3BD4-2828-4410-8290-418D1D73C762", "IEnumSpellingError"},
{"00020400-0000-0000-C000-000000000046", "IDispatch"}
};
List<string> lsSupported = new List<string>();
IntPtr ptrUnknown = IntPtr.Zero;
try {
ptrUnknown = Marshal.GetIUnknownForObject(oObject);
for (int i = 0; i < aKnown.GetLength(0); i++) {
Guid guid = new Guid(aKnown[i, 0]);
IntPtr ptrInterface;
if (Marshal.QueryInterface(ptrUnknown, ref guid, out ptrInterface) == 0) {
lsSupported.Add(aKnown[i, 1]);
Marshal.Release(ptrInterface);
}
}
}
catch (Exception ex) { return "could not be examined: " + ex.Message; }
finally { if (ptrUnknown != IntPtr.Zero) Marshal.Release(ptrUnknown); }
if (lsSupported.Count == 0) return "none of the spell checking interfaces";
return String.Join(", ", lsSupported.ToArray());
} // describeInterfaces method

// The object as a spell checker, if it is one. Returns null rather than
// throwing, so the caller can report what it got instead of dying on the
// cast.
// Not public: the interface it returns is private to this class, and C#
// will not let a public method hand back a type the caller cannot name.
// His machine answered the diagnostic with "ISpellChecker2" and nothing
// else: the object supports the newer interface but refuses the older
// one, whatever the header implies about their relationship. So the
// newer one is used when it is offered. The two share their first
// fourteen methods, so the code below works through either.
ISpellChecker asSpellChecker(object oObject) {
try { return (ISpellChecker) oObject; }
catch (Exception) {}
return null;
} // asSpellChecker method

ISpellChecker2 asSpellChecker2(object oObject) {
try { return (ISpellChecker2) oObject; }
catch (Exception) {}
return null;
} // asSpellChecker2 method

// Held as a plain object: naming the library's own type would make the
// compiler read every member of it, and some are declared in terms of a
// span, which .NET Framework 4.8 does not have. Only the two members
// EdSharp uses are reached, through reflection, just below.
static object wordListHunspell = null;
static bool bHunspellTried = false;

// The dictionary, loaded once. Two plain files in the Dictionaries
// folder beside the program: an affix file describing how words change,
// and a word list. Nothing is registered with Windows, nothing is
// downloaded at run time, and a missing dictionary simply means this
// engine is unavailable rather than an error.
static object hunspellDictionary(string sLanguage) {
if (bHunspellTried) return wordListHunspell;
bHunspellTried = true;
try {
string sDir = Path.Combine(App.ProgramDir, "Dictionaries");
string sBase = (sLanguage.Length > 0) ? sLanguage.Replace("-", "_") : "en_US";
string sAff = Path.Combine(sDir, sBase + ".aff");
string sDic = Path.Combine(sDir, sBase + ".dic");
if (!File.Exists(sAff) || !File.Exists(sDic)) {
sAff = Path.Combine(sDir, "en_US.aff");
sDic = Path.Combine(sDir, "en_US.dic");
}
if (!File.Exists(sAff) || !File.Exists(sDic)) {
Util.Log("spell check: no dictionary in " + sDir);
return null;
}
// Even loading the dictionary goes through reflection, so the compiler
// never reads the library's members and cannot object to the ones
// declared with types this framework lacks. The assembly is still
// referenced and still does all the work.
Type typeWordList = Type.GetType("WeCantSpell.Hunspell.WordList, WeCantSpell.Hunspell");
if (typeWordList == null) {
Util.Log("spell check: the Hunspell library was not found beside EdSharp");
return null;
}
System.Reflection.MethodInfo methodCreate = typeWordList.GetMethod("CreateFromFiles", new Type[] {typeof(string), typeof(string)});
if (methodCreate == null) {
Util.Log("spell check: the Hunspell library has no CreateFromFiles(string, string)");
return null;
}
wordListHunspell = methodCreate.Invoke(null, new object[] {sDic, sAff});
Util.Log("spell check: dictionary loaded from " + sDic);
}
catch (Exception ex) {
Util.Log("spell check: the dictionary could not be loaded: " + ex.Message);
wordListHunspell = null;
}
return wordListHunspell;
} // hunspellDictionary method

// Hunspell's own Check and Suggest are declared in terms of a span, a
// type .NET Framework 4.8 does not have, so calling them directly asks
// the compiler for an assembly that does not exist here. Reflection
// picks the plain-string form of each at run time, which the library
// still provides; the work is Hunspell's either way.
static System.Reflection.MethodInfo methodHunspellCheck = null;
static System.Reflection.MethodInfo methodHunspellSuggest = null;

static System.Reflection.MethodInfo hunspellMethod(object oWordList, string sName) {
foreach (System.Reflection.MethodInfo method in oWordList.GetType().GetMethods()) {
if (method.Name != sName) continue;
System.Reflection.ParameterInfo[] aParameters = method.GetParameters();
if (aParameters.Length == 1 && aParameters[0].ParameterType == typeof(string)) return method;
}
return null;
} // hunspellMethod method

static bool hunspellCheck(object oWordList, string sWord) {
if (methodHunspellCheck == null) methodHunspellCheck = hunspellMethod(oWordList, "Check");
if (methodHunspellCheck == null) return true;   // cannot judge, so do not accuse
return (bool) methodHunspellCheck.Invoke(oWordList, new object[] {sWord});
} // hunspellCheck method

static List<string> hunspellSuggest(object oWordList, string sWord) {
List<string> lsSuggestions = new List<string>();
if (methodHunspellSuggest == null) methodHunspellSuggest = hunspellMethod(oWordList, "Suggest");
if (methodHunspellSuggest == null) return lsSuggestions;
object oResult = methodHunspellSuggest.Invoke(oWordList, new object[] {sWord});
System.Collections.IEnumerable enumResult = oResult as System.Collections.IEnumerable;
if (enumResult != null) foreach (object oSuggestion in enumResult) if (oSuggestion != null) lsSuggestions.Add(oSuggestion.ToString());
return lsSuggestions;
} // hunspellSuggest method

// Check with Hunspell: walk the text word by word, and report each word
// the dictionary does not know, with its suggestions. Words are taken as
// runs of letters and apostrophes, so punctuation and numbers are left
// alone, and a word the person has added to their own list is accepted.
public List<SpellingProblem> findSpellingProblemsHunspell(string sText, string sLanguage) {
object wordList = hunspellDictionary(sLanguage);
if (wordList == null) return null;
List<SpellingProblem> lsProblems = new List<SpellingProblem>();
List<string> lsAdded = userDictionaryWords();
int i = 0;
while (i < sText.Length) {
if (!Char.IsLetter(sText[i])) { i++; continue; }
int iStart = i;
while (i < sText.Length && (Char.IsLetter(sText[i]) || sText[i] == '\'' || sText[i] == '\u2019')) i++;
string sWord = sText.Substring(iStart, i - iStart).Trim('\'', '\u2019');
if (sWord.Length < 2) continue;
if (lsAdded.Contains(sWord.ToLowerInvariant())) continue;
if (hunspellCheck(wordList, sWord)) continue;
SpellingProblem problem = new SpellingProblem();
problem.Start = iStart;
problem.Length = sWord.Length;
problem.Word = sWord;
problem.Suggestions = new List<string>();
try {
foreach (string sSuggestion in hunspellSuggest(wordList, sWord)) {
if (!problem.Suggestions.Contains(sSuggestion)) problem.Suggestions.Add(sSuggestion);
if (problem.Suggestions.Count >= 20) break;
}
}
catch (Exception) {}
lsProblems.Add(problem);
}
return lsProblems;
} // findSpellingProblemsHunspell method

// Words the person has added, kept in a plain file beside the settings
// so they survive updates and can be edited by hand.
public List<string> userDictionaryWords() {
List<string> lsWords = new List<string>();
try {
string sFile = Path.Combine(App.DataDir, "Dictionary.txt");
if (File.Exists(sFile)) foreach (string sLine in File.ReadAllLines(sFile)) {
string sWord = sLine.Trim().ToLowerInvariant();
if (sWord.Length > 0) lsWords.Add(sWord);
}
}
catch (Exception) {}
return lsWords;
} // userDictionaryWords method

public bool addWordToUserDictionary(string sWord) {
try {
string sFile = Path.Combine(App.DataDir, "Dictionary.txt");
File.AppendAllText(sFile, sWord + "\r\n");
return true;
}
catch (Exception ex) {
Util.Log("could not add to the personal dictionary: " + ex.Message);
return false;
}
} // addWordToUserDictionary method

public List<SpellingProblem> findSpellingProblems(string sText, string sLanguage) {
List<SpellingProblem> lsProblems = new List<SpellingProblem>();
try {
ISpellCheckerFactory factory = (ISpellCheckerFactory) new SpellCheckerFactoryClass();
Util.Log("spell check: factory created");
bool bSupported = false;
try { factory.IsSupported(sLanguage, out bSupported); }
catch (Exception) { bSupported = true; }
if (!bSupported && sLanguage != "en-US") {
sLanguage = "en-US";
try { factory.IsSupported(sLanguage, out bSupported); } catch (Exception) {}
}
object oChecker = factory.CreateSpellChecker(sLanguage);
Util.Log("spell check: created an object for " + sLanguage + "; " + describeInterfaces(oChecker));
ISpellChecker checker = asSpellChecker(oChecker);
ISpellChecker2 checker2 = (checker == null) ? asSpellChecker2(oChecker) : null;
if (checker == null && checker2 == null) {
App.SpellCheckError = "Windows returned an object that is not a spell checker. What it does support: " + describeInterfaces(oChecker);
Util.Log("spell check: " + App.SpellCheckError);
return null;
}
Util.Log("spell check: checker ready for " + sLanguage + " through " + ((checker != null) ? "ISpellChecker" : "ISpellChecker2"));
IEnumSpellingError errors = (checker != null) ? checker.ComprehensiveCheck(sText) : checker2.ComprehensiveCheck(sText);
while (true) {
ISpellingError error = null;
try { error = errors.Next(); }
catch (Exception) { break; }
if (error == null) break;
SpellingProblem problem = new SpellingProblem();
problem.Start = (int) error.StartIndex;
problem.Length = (int) error.Length;
if (problem.Start < 0 || problem.Length <= 0 || problem.Start + problem.Length > sText.Length) continue;
problem.Word = sText.Substring(problem.Start, problem.Length);
problem.Suggestions = new List<string>();
// CorrectiveAction 3 is "replace with this word" -- the system is
// certain -- so that replacement leads the suggestion list.
try {
if (error.CorrectiveAction == 3) {
string sReplacement = error.Replacement;
if (!String.IsNullOrEmpty(sReplacement)) problem.Suggestions.Add(sReplacement);
}
}
catch (Exception) {}
try {
// Suggest returns the standard COM string enumerator, walked one at a
// time; the count it fills is zero when the list is exhausted.
System.Runtime.InteropServices.ComTypes.IEnumString enumSuggestions = (checker != null) ? checker.Suggest(problem.Word) : checker2.Suggest(problem.Word);
if (enumSuggestions != null) {
string[] aOne = new string[1];
IntPtr ptrFetched = Marshal.AllocCoTaskMem(sizeof(int));
try {
while (problem.Suggestions.Count < 20) {
aOne[0] = null;
enumSuggestions.Next(1, aOne, ptrFetched);
if (Marshal.ReadInt32(ptrFetched) == 0) break;
string sSuggestion = aOne[0];
if (!String.IsNullOrEmpty(sSuggestion) && !problem.Suggestions.Contains(sSuggestion)) problem.Suggestions.Add(sSuggestion);
}
}
finally { Marshal.FreeCoTaskMem(ptrFetched); }
}
}
catch (Exception) {}
lsProblems.Add(problem);
}
return lsProblems;
}
catch (Exception ex) {
App.SpellCheckError = ex.Message;
Util.Log("spell check unavailable: " + ex.ToString());
return null;
}
} // findSpellingProblems method

// Add a word to the user's Windows dictionary, so every program that
// uses the system spell checker knows it from now on.
public bool addWordToDictionary(string sWord, string sLanguage) {
try {
ISpellCheckerFactory factory = (ISpellCheckerFactory) new SpellCheckerFactoryClass();
object oChecker = factory.CreateSpellChecker(sLanguage);
ISpellChecker checker = asSpellChecker(oChecker);
if (checker != null) { checker.Add(sWord); return true; }
ISpellChecker2 checker2 = asSpellChecker2(oChecker);
if (checker2 != null) { checker2.Add(sWord); return true; }
return false;
}
catch (Exception ex) {
Util.Log("could not add to dictionary: " + ex.Message);
return false;
}
} // addWordToDictionary method

// Walk the document's misspellings one at a time, the way Compile walks
// errors: the earliest problem in the text first, the caret moved to it
// and the word selected so the surrounding line reads with ordinary
// reading keys, then a small dialog whose first control already holds
// the answer.
//
// The shape follows what the screen readers do well. JAWS, NVDA, and
// VoiceOver all handle a spell check best when the misspelled word is
// SPOKEN FIRST and spelled out letter by letter, when the suggestions
// are a real list the arrow keys walk rather than a graphic, and when
// the choice needs no hunting for controls. So the dialog announces the
// word, spells it, and gives its position in the pass; the first
// control is an editable combo box holding the best suggestion, where
// Down and Up walk the rest and typing corrects the word by hand; and
// the buttons are few, plainly named, and ordered by how often they are
// wanted -- Replace, Skip, Add to Dictionary, Cancel -- each with its
// own access key (R, S, A, C) and Replace the default, so Enter accepts
// the top suggestion and moves on. Escape
// stops the pass and leaves the rest of the document alone.
public void SpellCheck() {
HomerRichTextBox rtb = this.Child.RTB;
if (App.ReadOption("SpellChecker", "Windows").Trim().ToLower() == "word") { SpellCheckWord(); return; }

int iStart;
string sText;
bool bWholeDocument = (rtb.SelectionLength == 0);
if (bWholeDocument) { AddMessage("All"); iStart = 0; sText = rtb.Text; }
else { AddMessage("Selected"); iStart = rtb.SelectionStart; sText = rtb.SelectedText; }
if (sText.Trim().Length == 0) { AddMessage("No text!"); return; }

string sLanguage = App.ReadOption("SpellLanguage", "en-US").Trim();
if (sLanguage.Length == 0) sLanguage = "en-US";
AddMessage("Checking");
// Hunspell first: it needs nothing from Windows, so it cannot be refused
// by it. The Windows checker is tried when Hunspell has no dictionary,
// and Word remains available through the SpellChecker option.
string sEngine = App.ReadOption("SpellChecker", "Hunspell").Trim().ToLower();
List<SpellingProblem> lsProblems = null;
if (sEngine != "windows") lsProblems = findSpellingProblemsHunspell(sText, sLanguage);
if (lsProblems == null) lsProblems = findSpellingProblems(sText, sLanguage);
if (lsProblems == null) {
Dialog.Show("Spell Check", "No spell checker could be started.\n\nThe dictionary should be here:\n" + Path.Combine(App.ProgramDir, @"Dictionaries\en_US.dic") + "\n\nThe Windows checker also refused: \n\n" + App.SpellCheckError + "\n\nThe run log has the detail:\n" + App.LogFile + "\n\nIf Microsoft Word is installed, set the SpellChecker option to Word with Configuration Options, Alt+Shift+C.");
return;
}
if (lsProblems.Count == 0) { AddMessage("No spelling problems"); return; }

// Problems are visited in document order. Replacements are applied to
// a working copy of the text and the positions of later problems are
// shifted by however much each replacement changed the length, so a
// long correction early in the file cannot misplace a later one.
int iChanged = 0, iSkipped = 0, iAdded = 0;
StringBuilder sbText = new StringBuilder(sText);
int iDrift = 0;
for (int iProblem = 0; iProblem < lsProblems.Count; iProblem++) {
SpellingProblem problem = lsProblems[iProblem];
int iWordStart = iStart + problem.Start + iDrift;
try {
rtb.Select(iWordStart, problem.Length);
rtb.ScrollToCaret();
}
catch (Exception) {}

string sContext = spellingContext(sbText.ToString(), problem.Start + iDrift, problem.Length);
List<string> lsChoices = new List<string>(problem.Suggestions);
if (lsChoices.Count == 0) lsChoices.Add(problem.Word);
StringBuilder sbSpelled = new StringBuilder();
foreach (char cLetter in problem.Word) { if (sbSpelled.Length > 0) sbSpelled.Append(' '); sbSpelled.Append(cLetter); }
string sPrompt = "Not in dictionary: " + problem.Word + ", spelled " + sbSpelled.ToString() + ". Word " + (iProblem + 1) + " of " + lsProblems.Count;
object[] aResult = Dialog.SpellChoice("Spell Check", sPrompt, sContext, lsChoices);
string sButton = (string) aResult[0];
string sReplacement = ((string) aResult[1]).Trim();

if (sButton == "Cancel") { AddMessage("Stopped"); break; }
if (sButton == "Skip") { iSkipped++; continue; }
if (sButton == "Add to Dictionary") {
// Both dictionaries learn the word: the personal list Hunspell reads,
// and Windows' own, so every program benefits when that checker works.
bool bAdded = addWordToUserDictionary(problem.Word);
addWordToDictionary(problem.Word, sLanguage);
if (bAdded) iAdded++;
else AddMessage("Could not add the word");
continue;
}
// Replace, the default: anything else that comes back leaves the word
// alone rather than guessing.
if (sButton != "Replace" || sReplacement.Length == 0 || sReplacement == problem.Word) { iSkipped++; continue; }
sbText.Remove(problem.Start + iDrift, problem.Length);
sbText.Insert(problem.Start + iDrift, sReplacement);
iDrift += sReplacement.Length - problem.Length;
iChanged++;
}

if (iChanged > 0) {
string sNewText = sbText.ToString();
if (bWholeDocument) {
int iCaret = rtb.Index;
rtb.Text = sNewText;
if (iCaret <= rtb.TextLength) rtb.Index = iCaret;
}
else {
rtb.Select(iStart, sText.Length);
rtb.SelectedText = sNewText;
}
rtb.Modified = true;
}
else rtb.Select(iStart, 0);
AddMessage(Util.Pluralize(iChanged, "change") + ", " + Util.Pluralize(iSkipped, "skip") + ", " + Util.Pluralize(iAdded, "word") + " added");
} // SpellCheck method

// The words around a misspelling, for the dialog to show: enough to
// place the word in its sentence without reading the line again.
string spellingContext(string sText, int iStart, int iLength) {
if (iStart < 0 || iLength <= 0 || iStart + iLength > sText.Length) return "";
int iFrom = Math.Max(0, iStart - 60);
int iTo = Math.Min(sText.Length, iStart + iLength + 60);
string sPart = sText.Substring(iFrom, iTo - iFrom);
return sPart.Replace("\r", " ").Replace("\n", " ").Trim();
} // spellingContext method

// The original Microsoft Word spell check, kept as a fallback for
// anyone who prefers its dialog and has Word installed. Reached by
// setting the SpellChecker option to "Word".
public void SpellCheckWord() {
bool bCreate, bVisible;
int iDisplayAlerts, iStart, iEnd, iLength;
string sText, sOldText;
object oApp, oDocs, oDoc, oSelection;

HomerRichTextBox rtb = this.Child.RTB;
if (rtb.SelectionLength == 0) {
AddMessage("All");
iStart = 0;
sText = rtb.Text;
}
else {
AddMessage("Selected");
iStart = rtb.SelectionStart;
sText = rtb.SelectedText;
}

if (sText.Length == 0) {
AddMessage("No text!");
return;
}

iEnd = iStart + sText.Length;
sText = sText.TrimEnd();
sOldText = sText;
sText = Util.Convert2MacLineBreak(sText);

bool bAppVisible = true;
//oApp = COM.GetOrCreateObject("Word.Application", out bCreate);
oApp = COM.WordAccess(out bCreate);
bVisible = (bool) COM.GetProperty(oApp, "Visible");
iDisplayAlerts = (int) COM.GetProperty(oApp, "DisplayAlerts");
COM.SetProperty(oApp, "Visible", bAppVisible);
COM.SetProperty(oApp, "DisplayAlerts", 0);
oDocs = COM.GetProperty(oApp, "Documents");
oDoc = COM.CallMethod(oDocs, "Add");
COM.CallMethod(oDoc, "Activate");
oSelection = COM.GetProperty(oApp, "Selection");
COM.CallMethod(oSelection, "TypeText", sText);
Util.ActivateProcess("WinWord");
COM.CallMethod(oDoc, "CheckSpelling");
iLength = (int) COM.GetProperty(oSelection, "StoryLength");
COM.CallMethod(oSelection, "SetRange", new object[] {0, iLength});
sText = (string) COM.GetProperty(oSelection, "Text");
sText = sText.Trim();
COM.Release(ref oSelection);
VB.WordClose(oDoc);
COM.Release(ref oDoc);
COM.Release(ref oDocs);
if (bCreate) {
//VB.WordQuit(oApp);
}
else {
COM.SetProperty(oApp, "Visible", bVisible);
COM.SetProperty(oApp, "DisplayAlerts", iDisplayAlerts);
}
COM.Release(ref oApp);

App.Frame.Activate();
App.Frame.Child.RTB.Select();
sText = Util.Convert2UnixLineBreak(sText);
if (sText == sOldText) AddMessage("No changes!");
else {
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
//AddMessage("Done");
}
} // SpellCheck method

// Shift+F7: synonyms without Microsoft Word. WordNet, Princeton's freely
// licensed lexical database, gives synonyms grouped by MEANING -- which
// is the part a thesaurus is really for, since "light" as the opposite
// of heavy and "light" as illumination want different words -- along
// with the part of speech and a short definition of each sense. The
// database is read by a small Python helper that EdSharp installs on
// request, the same way the PDF reader is installed; when it is absent,
// the Word thesaurus still answers for anyone who has Word.
public void Thesaurus() {
if (App.ReadOption("Thesaurus", "WordNet").Trim().ToLower() == "word") { ThesaurusWord(); return; }
HomerRichTextBox rtb = this.Child.RTB;
string sWord;
int iStart, iLength;
if (rtb.SelectionLength > 0) { sWord = rtb.SelectedText.Trim(); iStart = rtb.SelectionStart; iLength = rtb.SelectionLength; }
else {
// The word under the caret, taken by letters so punctuation stops it.
string sLine = rtb.RowText;
int iColumn = rtb.Column - 1;
if (sLine.Length == 0) { AddMessage("No word!"); return; }
if (iColumn >= sLine.Length) iColumn = sLine.Length - 1;
if (iColumn < 0) iColumn = 0;
int iFrom = iColumn, iTo = iColumn;
while (iFrom > 0 && (Char.IsLetter(sLine[iFrom - 1]) || sLine[iFrom - 1] == '\'')) iFrom--;
while (iTo < sLine.Length && (Char.IsLetter(sLine[iTo]) || sLine[iTo] == '\'')) iTo++;
sWord = sLine.Substring(iFrom, iTo - iFrom).Trim();
iStart = rtb.RowStart + iFrom;
iLength = sWord.Length;
}
if (sWord.Length == 0 || !Char.IsLetter(sWord[0])) { AddMessage("No word!"); return; }

AddMessage("Looking up " + sWord);
string sHelper = Path.Combine(App.ProgramDir, @"Convert\wordNet.py");
if (!File.Exists(sHelper)) { Dialog.Show("Thesaurus", "The thesaurus helper was not found at:\n" + sHelper); return; }
string sOutput = Util.GetProgramOutput("python", "\"" + sHelper + "\" \"" + sWord + "\"");
if (sOutput == null) sOutput = "";
sOutput = sOutput.Trim();
if (sOutput.StartsWith("NOTINSTALLED")) {
Dialog.Show("Thesaurus", "The free thesaurus database is not installed yet.\n\nRun installPdfTools.cmd in the EdSharp program folder -- it installs the thesaurus as well -- or set the Thesaurus option to Word with Configuration Options, Alt+Shift+C, to use the Microsoft Word thesaurus instead.");
return;
}
if (sOutput.StartsWith("NOWORD") || sOutput.Length == 0) { AddMessage("No entry for " + sWord); return; }

// Each line the helper writes is one choice: the replacement word, then
// a tab, then how it reads in the list -- word, part of speech, and the
// sense it belongs to.
List<object> lValues = new List<object>();
List<string> lDisplay = new List<string>();
foreach (string sLine in sOutput.Replace("\r\n", "\n").Split('\n')) {
if (sLine.Trim().Length == 0) continue;
string[] aParts = sLine.Split('\t');
lValues.Add(aParts[0]);
lDisplay.Add((aParts.Length > 1) ? aParts[1] : aParts[0]);
}
if (lValues.Count == 0) { AddMessage("No entry for " + sWord); return; }

object[] aChoice = Dialog.PickAndChoose("Thesaurus: " + sWord, lValues.ToArray(), lDisplay.ToArray(), new string[] {"&Replace", "Cop&y"}, false, 0);
if (aChoice.Length == 0) return;
string sChosen = aChoice[0].ToString();
string sButton = ((string) aChoice[1]).Replace("&", "");
if (sButton == "Copy") {
try { Clipboard.SetText(sChosen); AddMessage("Copied " + sChosen); }
catch (Exception ex) { Dialog.Show("Thesaurus", ex.Message); }
return;
}
// Replace, keeping the original capitalization of the word replaced.
if (Char.IsUpper(sWord[0]) && sChosen.Length > 0) sChosen = Char.ToUpper(sChosen[0]) + sChosen.Substring(1);
rtb.Select(iStart, iLength);
rtb.SelectedText = sChosen;
rtb.Modified = true;
AddMessage("Replaced with " + sChosen);
} // Thesaurus method

public void ThesaurusWord() {
bool bCreate, bVisible;
int iDisplayAlerts, iStart, iEnd, iLength;
string sText, sOldText;
object[] aResults;
object oApp, oDocs, oDoc, oSelection, oRange;

HomerRichTextBox rtb = this.Child.RTB;
if (rtb.SelectionLength == 0) {
//AddMessage("Chunk");
aResults = GetChunk();
iStart = (int) aResults[0];
sText = (string) aResults[1];
}
else {
//AddMessage("Selected");
iStart = rtb.SelectionStart;
sText = rtb.SelectedText;
}

sText = sText.TrimEnd();
if (sText.Length == 0) {
AddMessage("No text!");
return;
}

iEnd = iStart + sText.Length;
sOldText = sText;
sText = Util.Convert2MacLineBreak(sText);

bool bAppVisible = true;
//oApp = COM.GetOrCreateObject("Word.Application", out bCreate);
oApp = COM.WordAccess(out bCreate);
bVisible = (bool) COM.GetProperty(oApp, "Visible");
iDisplayAlerts = (int) COM.GetProperty(oApp, "DisplayAlerts");
COM.SetProperty(oApp, "Visible", bAppVisible);
COM.SetProperty(oApp, "DisplayAlerts", 0);
oDocs = COM.GetProperty(oApp, "Documents");
oDoc = COM.CallMethod(oDocs, "Add");
oSelection = COM.GetProperty(oApp, "Selection");
COM.CallMethod(oSelection, "TypeText", sText);
oRange = COM.GetProperty(oSelection, "Range");
Util.ActivateProcess("WinWord");
COM.CallMethod(oRange, "CheckSynonyms");
iLength = (int) COM.GetProperty(oSelection, "StoryLength");
COM.CallMethod(oSelection, "SetRange", new object[] {0, iLength});
sText = (string) COM.GetProperty(oSelection, "Text");
sText = sText.Trim();
COM.Release(ref oRange);
COM.Release(ref oSelection);
VB.WordClose(oDoc);
COM.Release(ref oDoc);
COM.Release(ref oDocs);
if (bCreate) {
//VB.WordQuit(oApp);
}
else {
COM.SetProperty(oApp, "Visible", bVisible);
COM.SetProperty(oApp, "DisplayAlerts", iDisplayAlerts);
}
COM.Release(ref oApp);

App.Frame.Activate();
App.Frame.Child.RTB.Select();
sText = Util.Convert2UnixLineBreak(sText);
if (sText == sOldText) AddMessage("No changes!");
else {
rtb.ReplaceRange(iStart, iEnd, sText);
rtb.Index = iStart;
//AddMessage("Done");
}
} // Thesaurus method

public void ElevateVersion() {
// Check GitHub for the latest EdSharp release and, if the user agrees, download
// and run its installer.  This replaces the old AppStamp.ini / Win32.Url2File
// mechanism with the GitHub Releases approach used by the sibling DbDo project.
// All network work goes through Homer.Web, which sends a User-Agent and uses
// modern TLS.  EdSharp is not closed here: the Inno Setup installer detects the
// running EdSharp and offers to close it before proceeding.
string sOwnerRepo = "JamalMazrui/EdSharp";
string sReleasesUrl = "https://github.com/" + sOwnerRepo + "/releases/latest";
string sName = "EdSharp_Setup.exe";

Util.Say("Checking for updates");
string sTag = Util.FetchLatestReleaseTag(sOwnerRepo);
if (sTag.Length == 0) {
Dialog.Show("Elevate Version", "Could not check for updates right now.\nPlease check your internet connection and try again.\nYou can also download the latest installer from\n" + sReleasesUrl);
return;
}

string sLocal = App.VersionString;
string sLatest = sTag.TrimStart('v', 'V').Trim();
int iCompare = Util.CompareVersions(sLatest, sLocal);
string sDefault = "N";
string sMsg;
if (iCompare > 0) {
sMsg = "A newer EdSharp is available.\nInstalled: " + sLocal + "\nAvailable: " + sLatest + "\n\nDownload and run the new installer now?";
sDefault = "Y";
}
else if (iCompare == 0) sMsg = "EdSharp's version number (" + sLocal + ") matches the latest release (" + sLatest + "), so no newer version was detected.\nA newer build may still have been published under the same version number.\n\nDownload and install the latest release from the web now?";
else sMsg = "EdSharp's version number (" + sLocal + ") is higher than the latest public release (" + sLatest + ").\n\nDownload and install the latest public release from the web anyway?";
if (Dialog.Confirm("Elevate Version", sMsg, sDefault) != "Y") return;

Util.Say("Downloading installer");
string sUrl = sReleasesUrl + "/download/" + sName;
string sFile = Homer.Web.download(sUrl, Path.GetTempPath(), sName);
if (sFile.Length == 0) {
Dialog.Show("Elevate Version", "The download did not complete.\nYou can download the installer manually from\n" + sReleasesUrl);
return;
}

Util.Say("Starting installer");
try {
ProcessStartInfo processStartInfo = new ProcessStartInfo();
processStartInfo.FileName = sFile;
processStartInfo.UseShellExecute = true;
Process.Start(processStartInfo);
}
catch (Exception ex) {
Dialog.Show("Elevate Version", "The installer downloaded but could not be started.\n" + ex.Message + "\n\nThe file is here:\n" + sFile);
}
} // ElevateVersion method

public bool ExitApp() {
while (this.Child != null) {
if (!CloseWindow(this.Child, true)) return false;
}
Application.Exit();
return true;
} // ExitApp method

public bool CloseWindow(MdiChild child) {
bool bExiting = false;
return CloseWindow(child, bExiting);
} // CloseWindow method

public bool CloseWindow(MdiChild child, bool bExiting) {
HomerRichTextBox rtb = child.RTB;
if (rtb.Modified) {
switch (Dialog.Confirm("Confirm", "Save changes to " + child.Text + "?", "Y")) {
case "Y" :
menuFileSave.PerformClick();
if (rtb.Modified) return false;
else break;
case "" :
return false;
}
}

if (bExiting && !rtb.Modified && child.File.IndexOf(@"\") >=0 && !Util.Equiv(child.File, App.IniFile)) Ini.WriteValue(App.IniFile, "Previous", child.File, rtb.Index.ToString());
child.Close();
return true;
} // CloseWindow method

public void SetRecent(string sFile) {
if (!sFile.Contains(@"\")) return;
if (Util.Equiv(sFile, App.IniFile)) return;

DateTime dt = DateTime.Now;
string sTime = dt.ToString("u");
sTime = sTime.Substring(0, sTime.Length - 1);
int iIndex = this.Child.RTB.Index;
sTime += "|" + iIndex;
if (this.MdiChildren.Length == 0) sTime += "|N|W";
else sTime += "|" + (this.Child.RTB.ReadOnly ? "G" : "M") + "|" + Util.If(this.Child.RTB.WordWrap, "W", "U");
App.WriteValue("Recent", sFile, sTime);
string sDir = Path.GetDirectoryName(sFile);
if (Directory.Exists(sDir)) Directory.SetCurrentDirectory(sDir);
sFile = Path.Combine(App.DataDir, App.ReadData("Compiler", "Default") + ".ini");
Ini.WriteValue(sFile, "Data", "Directory", sDir);
} // SetRecent method

bool ApplyWrap(string sSection, string sFile) {
string sText = App.ReadValue(sSection, sFile, "");
if (sText.Length == 0) return false;

HomerRichTextBox rtb = this.Child.RTB;
sText = App.ReadOption("WordWrap", "Y");
sText = "-1|M|" + Util.If((sText == "N"), "U", "W");
sText = App.ReadValue(sSection, sFile, sText);
bool b = (bool) Util.If(sText.EndsWith("U"), false, true);
if (b && !rtb.WordWrap) {
AddMessage("Word wrap");
rtb.SetWrap(true);
}
else if (!b && rtb.WordWrap) {
AddMessage("Unwrap");
rtb.SetWrap(false);
}
return true;
} // ApplyWrap method

bool ApplyGuard(string sSection, string sFile) {
string sText = App.ReadValue(sSection, sFile, "");
if (sText.Length == 0) return false;

HomerRichTextBox rtb = this.Child.RTB;
sText = App.ReadOption("WordWrap", "Y");
sText = "-1|M|" + Util.If((sText == "N"), "U", "W");
sText = App.ReadValue(sSection, sFile, sText);
if (sText.IndexOf("G") >= 0) {
AddMessage("Guard");
rtb.SetGuard(true);
}
return true;
} // ApplyGuard method

public void ApplyFileOptions(string sFile) {
if (!ApplyGuard("Favorites", sFile)) ApplyGuard("Recent", sFile);
if (!ApplyWrap("Favorites", sFile)) {
ApplyWrap("Recent", sFile);
return;
}

HomerRichTextBox rtb = this.Child.RTB;
string sText = App.ReadValue("Favorites", sFile, "");
try {
string[] a = sText.Split('|');
sText = a[0];
rtb.Index = Int32.Parse(sText);
AddMessage("Bookmark at percent " + rtb.Percent);
}
catch {}
} // ApplyFavorite method

public string PickSpecialFolder() {
string sName = "";
string sPath = "";
StringBuilder sbNames = new StringBuilder();
StringBuilder sbPaths = new StringBuilder("\n");
object oShell = COM.CreateObject("Shell.Application");
for (int i = 0; i < 100; i++) {
try {
Object oDir = COM.CallMethod(oShell, "Namespace", new object[] {i});
Object oItem = COM.GetProperty(oDir, "Self");
sPath = (string) COM.GetProperty(oItem, "Path");
if (!Directory.Exists(sPath)) continue;
if (Util.IsNumeric(Path.GetFileName(sPath))) continue;
if (sbPaths.ToString().ToLower().Trim('\\').Contains("\n" + sPath.ToLower().Trim('\\') + "\n")) continue;
sbPaths.Append(sPath + "\n");
sName = (string) COM.GetProperty(oItem, "Name");
if (Util.Equiv(sName, "Temporary Internet Files")) sName = "Internet Cache";
else if (Util.Equiv(sName, "History")) sName = "Internet History";
else if (Util.Equiv(sName, "NetHood")) sName = "Network Neighborhood";
else if (Util.Equiv(sName, "PrintHood")) sName = "Printer Neighborhood";
else if ((@"\" + sPath.ToLower() + @"\").Contains(@"\all users\")) sName = "Common " + sName;
else if (!Util.Equiv(sName, "History") && (@"\" + sPath.ToLower() + @"\").Contains(@"\local settings\")) sName = "Local " + sName;
sbNames.Append(sName + "\n");
}
catch {
continue;
}
}

Environment.SpecialFolder folder;
for (int i = 0; i < 100; i++) {
sPath = "";
try {
folder = (Environment.SpecialFolder) i;
sPath = Environment.GetFolderPath(folder);
}
catch {
continue;
}
if (!Directory.Exists(sPath)) continue;
if (Util.IsNumeric(Path.GetFileName(sPath))) continue;
if (sbPaths.ToString().ToLower().Trim('\\').Contains("\n" + sPath.ToLower().Trim('\\') + "\n")) continue;
sbPaths.Append(sPath + "\n");
sName = folder.ToString();
sbNames.Append(sName + "\n");
}
sbNames.Append("Temp" + "\n");
sbPaths.Append(Util.GetTempFolder() + "\n");

string[] aNames = sbNames.ToString().Trim().Split('\n');
string[] aPaths = sbPaths.ToString().Trim().Split('\n');

string sDir = Dialog.Pick("Go to Special Folder", aPaths, aNames, true, 0);
return sDir;
} // PickSpecialFolder method

public void OpenOrActivateWindow(string sFile) {
int iConvert = 0;
OpenOrActivateWindow(sFile, iConvert);
} // OpenOrActivateWindow method

public void OpenOrActivateWindow(string sFile, int iConvert) {
string sLine = "";
string sColumn = "";
OpenOrActivateWindow(sFile, iConvert, sLine, sColumn);
} // OpenOrActivateWindow method

public void OpenOrActivateWindow(string sFile, int iConvert, string sLine, string sColumn) {
string sText;
sFile = Util.Unquote(sFile);
if (!File.Exists(sFile)) {
AddMessage("File not found!");
return;
}

sFile = Util.GetLfn(sFile);
// ApplyFileOptions(sFile);
// SetRecent(sFile);
object[] children = this.MdiChildren;
foreach (MdiChild child in children) {
if (Util.Equiv(child.File, sFile)) {
Util.Say("returning");
child.Activate();
SetCursorPosition(child.RTB, sLine, sColumn);
return;
}
}

string sTargetExt = "txt";
if (iConvert == 0) sText = "";
else {
// Dialog.Show("iConvert " + iConvert, "sTargetExt " + sTargetExt);
sText = COM.ConvertFile2String(sFile, ref iConvert, ref sTargetExt);
// Dialog.Show("iConvert " + iConvert, "sTargetExt " + sTargetExt);

if (iConvert >= 1 && sText.Trim().Length == 0) {
AddMessage("No text!");
return;
}
// Disable because also speaks after recent files
// else App.Frame.AddMessage("Done");
}

// Did so above
// SetRecent(sFile);
//if (!IsEmptyWindow()) new MdiChild(this);
//if (!IsEmptyWindow()) new MdiChild(this, "");
if (!IsEmptyWindow()) new MdiChild(this, sFile);
if (iConvert <= 0) {
this.Child.LoadTextOrRtfFile(sFile, (iConvert == 0 ? true : false));
//Dialog.Show(sFile);
ApplyFileOptions(sFile);

if (sFile == App.IniFile) return;
}
else {
string s = sText.Trim().ToLower();
if (s.StartsWith(@"{\rtf") && s.EndsWith("}")) this.Child.RTB.Rtf = sText;
else this.Child.RTB.Text = sText;
this.Child.Text = Path.GetFileNameWithoutExtension(sFile) + "." + sTargetExt;
this.Child.File = this.Child.Text;
// A converted document opens under a real base name in the title, so
// closing it with Control+F4 must ask about saving -- the content came
// from work the person asked for and has no disk home yet. Windows
// titled NoName style stay quiet on close, since that pattern marks
// temporary output whose saving is a deliberate act.
this.Child.RTB.Modified = true;

}
// Try disabling for auto bookmark
// SetRecent(sFile);

SetCursorPosition(this.Child.RTB, sLine, sColumn);
} // OpenOrActivateWindow method

public static bool SetCursorPosition(HomerRichTextBox rtb, string sLine, string sColumn) {
bool bReturn = false;
try {
if (sLine.Length > 0) rtb.Line = Int32.Parse(sLine);
if (sColumn.Length > 0) rtb.Column = Int32.Parse(sColumn);
bReturn = true;
}
catch {}
return bReturn;
} // SetCursorPosition method

public void NextWindow() {
object[] children = this.MdiChildren;
if (children.Length == 0) AddMessage("No windows!");
else if (children.Length == 1) AddMessage("Only this window!");
else {
MdiChild child = this.Child;
int iPosition = Array.IndexOf(children, child);
iPosition++;
if (iPosition == children.Length) iPosition = 0;
((MdiChild) children[iPosition]).Activate();
}
} // NextWindow method

public void PriorWindow() {
object[] children = this.MdiChildren;
if (children.Length == 0) AddMessage("No windows!");
else if (children.Length == 1) AddMessage("Only this window!");
else {
MdiChild child = this.Child;
int iPosition = Array.IndexOf(children, child);
iPosition--;
if (iPosition == -1) iPosition = children.Length - 1;
((MdiChild) children[iPosition]).Activate();
}
} // PriorWindow method

public void CloseAllButCurrentWindow() {
MdiChild child = this.Child;
if (child == null) return;

object[] children = this.MdiChildren;
int iCount = 0;
foreach (MdiChild o in children) {
if (o != child) {
o.Close();
iCount++;
}
}
} // CloseAllButCurrent method

public void WindowsOpen() {
object[] children = this.MdiChildren;
int iCount = children.Length;
if (iCount == 0) AddMessage("No windows!");
else {
//string s = Util.Pluralize(iCount, "window");
AddMessage(iCount);
foreach (MdiChild child in children) {
string sTitle = child.Text;
string sText = sTitle;
if (this.KeyRepeat % 2 == 0) AddMessage(sText);
else {
Util.Spell(sText);
}
}
}
} // WindowsOpen method

public void NavigateNextMatch(string sMatch) {
bool bLine = false;
NavigateNextMatch(sMatch, bLine);
} // NavigateNextMatch method

public void NavigateNextMatch(string sMatch, bool bLine) {
int iIndex, iStart, iEnd, iForward;
string sValue, sText;
object[] aResults;
HomerRichTextBox rtb = this.Child.RTB;
if (bLine) iIndex = rtb.RowEnd + 1;
else iIndex = rtb.Index;
iStart = iIndex;
iEnd = rtb.TextLength;
if (iStart >= iEnd) aResults = new object[] {-1, ""};
else {
sText = rtb.GetRange(iStart, iEnd);
aResults = Util.RegExpContainsEquiv(sText, sMatch);
}
if ((int) aResults[0] == -1) {
this.AddMessage("Bottom!");
iStart = iEnd;
iIndex = iEnd;
}
else if (bLine) {
iIndex += (int) aResults[0];
iIndex += ((string) aResults[1]).Length;
}
else {
iForward = (int) aResults[0];
sValue = (string) aResults[1];
iIndex += iForward + sValue.Length;
iStart = iIndex;
sText = rtb.GetRange(iStart, iEnd);
aResults = Util.RegExpContainsEquiv(sText, sMatch);
if ((int) aResults[0] == -1) {
}
else {
iForward = (int) aResults[0];
sValue = (string) aResults[1];
iEnd = iStart + iForward + sValue.Length;
}
}

if (bLine) {
rtb.Index = iIndex;
rtb.Col = 0;
sText = rtb.RowText;
}
else {
sText = rtb.GetRange(iStart, iEnd);
rtb.Index = iIndex;
}
this.AddMessage(sText);
} // NavigateNextMatch method

public void NavigatePriorMatch(string sMatch) {
bool bLine = false;
NavigatePriorMatch(sMatch, bLine);
} // NavigatePriorMatch method

public void NavigatePriorMatch(string sMatch, bool bLine) {
int iIndex, iStart, iEnd, iBackward;
string sValue, sText;
object[] aResults;
HomerRichTextBox rtb = this.Child.RTB;
if (bLine) iIndex = rtb.RowStart;
else iIndex = rtb.Index;
iStart = 0;
iEnd = iIndex;
sText = rtb.GetRange(iStart, iEnd);
aResults = Util.RegExpContainsLastEquiv(sText, sMatch);
if ((int) aResults[0] == -1) {
this.AddMessage("Top!");
iIndex = iStart;
iEnd = iStart;
}
else if (bLine) {
iIndex = (int) aResults[0];
iIndex += ((string) aResults[1]).Length;

if (iIndex == rtb.Index) {
iEnd = (int) aResults[0];
sText = rtb.GetRange(iStart, iEnd);
aResults = Util.RegExpContainsLastEquiv(sText, sMatch);
iIndex = (int) aResults[0];
if ((int) aResults[0] == -1) {
this.AddMessage("Top!");
iIndex = iStart;
iEnd = iStart;
}
else iIndex += ((string) aResults[1]).Length;
}
}
else {
iBackward = (int) aResults[0];
sValue = (string) aResults[1];
// Dialog.Show(sValue, iBackward);
iEnd = iBackward;
sText = rtb.GetRange(iStart, iEnd);
aResults = Util.RegExpContainsLastEquiv(sText, sMatch);
if ((int) aResults[0] == -1) {
iIndex = iStart;
}
else {
iBackward = (int) aResults[0];
sValue = (string) aResults[1];
// Dialog.Show(sValue, iBackward);
iStart = iBackward + sValue.Length;
iIndex = iStart;
}
}

if (bLine) {
rtb.Index = iIndex;
rtb.Col = 0;
sText = rtb.RowText;
}
else {
sText = rtb.GetRange(iStart, iEnd);
rtb.Index = iIndex;
}
this.AddMessage(sText);
} // NavigatePriorMatch method

public void FileFind() {
string sContains, sFilter, sDir, sFile;
string[] aLabels, aValues, aFilters, aResults, aFiles, aNames;

sContains = App.ReadData("Contains", "");
sFilter = App.ReadData("Filter", "*.*");
string sTitle = "Open Folder";
sDir = Dialog.OpenFolder(sTitle, "Name", Directory.GetCurrentDirectory());
if (sDir.Length == 0) return;

Directory.SetCurrentDirectory(sDir);
aLabels = new string[] {"&Contains", "&Filter"};
aValues = new string[] {sContains, sFilter};
aResults = Dialog.MultiInput("Criteria", aLabels, aValues);
if (aResults.Length == 0) return;

sContains = aResults[0];
sFilter = aResults[1].Trim();
if (sFilter.Length == 0) sFilter = "*.*";
App.WriteData("Contains", sContains);
App.WriteData("Filter", sFilter);
aFilters = sFilter.Split('|');
sDir = Directory.GetCurrentDirectory();
aFiles = Util.FindInFiles(sContains, sDir, aFilters, false);
if (aFiles.Length == 0) {
Dialog.Show("Alert", "No matches!");
return;
}

aNames = new string[aFiles.Length];
for (int i = 0; i < aNames.Length; i++) aNames[i] = Path.GetFileName(aFiles[i]);
//Array.Sort(aNames, aFiles);
sFile = Dialog.Pick("Pick", aFiles, aNames, true, 0);
if (sFile.Length == 0) return;

OpenOrActivateWindow(sFile, 1);
/*
string[] aNames = null;
string[] aPaths = null;
int iIndex = -1;
string sPath = "";
string sName = "";
string sPaths = "";
string sNames = "";

string sDir = Directory.GetCurrentDirectory();
string sMatch = App.ReadData("FileFindMatch", "");
string sFilter = App.ReadData("FileFindFilter", "");
string sFields = "&Text\t&Filter";
string sValues = sMatch + "\t" + sFilter;
string[] aFields = sFields.Split('\t');
string[] aValues = sValues.Split('\t');
string[] aResults = Dialog.MultiInput("File Find", aFields, aValues);
if (aResults.Length == 0) return;

sMatch = aResults[0];
sFilter = aResults[1];
if (true) {
//if (sDir == App.sFileFindDir && sMatch == App.sFileFindMatch && sFilter == App.sFileFindFilter) {
AddMessage("Repeat search");
//aNames = App.aFileFind;
//iIndex = App.iFileFind + 1;
if (iIndex == -1) iIndex = 0;
}
else {
AddMessage("Please wait");
//ReadOnlyCollection<string> oPaths = null;
string[] aPaths = Util.GetFiles(sDir);
//if (sMatch == "") oPaths = LbcVB.GetFiles(sDir, sFilter);
//else oPaths = LbcVB.FindInFiles(sDir, sMatch, sFilter);
//if (oPaths.Count == 0) {
AddMessage("No files found!");
return;
}
//foreach (string s in oPaths) {
foreach (string s in aPaths) {
sPaths += s + "\n";
sName = s.Substring(sDir.Length + 1);
sNames += sName + "\n";
}
aPaths = sPaths.Trim().Split('\n');
aNames = sNames.Trim().Split('\n');
iIndex = 0;
}

App.WriteData("FileFindMatch", sMatch);
App.WriteData("FileFindFilter", sFilter);
sName = Dialog.Pick("Pick", aNames, true, iIndex);
if (sName.Length == 0) return;

int iName = Array.IndexOf(aNames, sName);
App.WriteData("FileFindDir", sDir);
App.aFileFind = aNames;
App.iFileFind = iName;
sPath = aPaths[iName];
sDir = Path.GetDirectoryName(sPath);
if (sDir.Length == 0) return;
if (Directory.Exists(sDir)) {
OpenOrActivateWindow(sFile, 1);
}
else AddMessage("Folder " + sDir + " not found!");
*/
} //FileFind method

public void CurrentWindows() {
object[] children = this.MdiChildren;
string sTitles = "";
foreach (MdiChild child in children) {
sTitles += child.Text + "\n";
}
string[] aTitles = sTitles.Trim().Split('\n');
string sTitle = Dialog.Pick("Current Windows", aTitles, true, 0);
if (sTitle.Length == 0) return;

int iTitle = Array.IndexOf(aTitles, sTitle);
((MdiChild) children[iTitle]).Activate();
} // CurrentWindows method

public void ExplorerFolder(string sDir) {
string sCommand = sDir;
try {
Process.Start(sCommand);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}
} // ExplorerFolder method

public void CommandPrompt(string sDir) {
string sCommand = Environment.GetEnvironmentVariable("COMSPEC");
sDir = Util.Quote(sDir);
string sParams = "/k cd " + sDir;
try {
Process.Start(sCommand, sParams);
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
return;
}
} // CommandPrompt method

public void BurnToCD() {
string[] aPaths = GetPathsFromDocument();
// MessageBox.Show(String.Join("\n", aPaths), "Paths");
if (aPaths.Length == 0) return;
string sPathList = String.Join("\r\n", aPaths).Trim() + "\r\n";
Util.DeleteFile(App.TempFile, false);
Util.String2File(sPathList, App.TempFile);
string sExe = Path.Combine(App.ProgramDir, "Burn2CD.exe");
Util.Run(sExe + " " + App.TempFile);
} // BurnToCD method

public string[] GetPathsFromDocument() {
HomerRichTextBox rtb = App.Frame.Child.RTB;
List<string> list = new List<string>();
string[] aResults = rtb.Lines;
string sDir = Directory.GetCurrentDirectory();
string sTempDir = "";
for (int i = 0; i < aResults.Length; i++) {
string s = aResults[i].Trim();
if (s.Length == 0) continue;
sTempDir = Path.GetDirectoryName(s);
if (Directory.Exists(sTempDir)) sDir = sTempDir;
if (!File.Exists(s) && !Directory.Exists(s)) s = Path.Combine(sDir, s);
if (File.Exists(s) || Directory.Exists(s)) list.Add(s);
}

aResults = list.ToArray();
if (aResults.Length == 0) {
AddMessage("No files found!");
return aResults;
}

string sText = Util.GetExtensions(aResults);
string sResult = Dialog.Input("Filter", "Extensions", sText, "FilterExtensions").Trim();
if (sResult.Length == 0) return new string[] {};

aResults = Util.GetPathsWithExtensions(aResults, sResult);
if (aResults.Length == 0) AddMessage("No files!");
return aResults;
} // GetPathsFromDocument method

public void AlternateMenu() {
int iChoice = -1;
List<ToolStripMenuItem> items = new List<ToolStripMenuItem>();
string sItems = "";
StringBuilder sb = new StringBuilder();
foreach (ToolStripMenuItem menu in menuMain.Items) {
foreach (object o in menu.DropDownItems) {
ToolStripMenuItem item = o as ToolStripMenuItem;
if (item == null) continue;
if (item == menuHelpAlternateMenu) continue;
// string sText = item.Text.Replace("&", "") + "\t" + item.ShortcutKeyDisplayString;
// if ("1234567890".Contains(sText.Substring(0, 1))) continue;
if (item.IsMdiWindowListEntry) continue;
string[] aSummary = GetKeySummary(item);
string sText = aSummary[0] + " = " + aSummary[1] + ", " + aSummary[2];
sb.Append(sText + "\n");
items.Add(item);
}
}
sItems = sb.ToString();
string[] aItems = sItems.Trim().Split('\n');
string sItem = Dialog.Pick("Alternate Menu", aItems, true, 0);
if (sItem.Length == 0) return;

foreach (ToolStripMenuItem item in items) {
//if (sItem == item.Text.Replace("&", "")) {
// if (sItem == item.Text.Replace("&", "") + "\t" + item.ShortcutKeyDisplayString) {
string[] aSummary = GetKeySummary(item);
string sText = aSummary[0] + " = " + aSummary[1] + ", " + aSummary[2];
if (sItem == sText) {
iChoice = items.IndexOf(item);
break;
}
}
items[iChoice].PerformClick();
} // AlternateMenu method

new void ContextMenu(string sFile) {
MdiChild child = this.Child;

string[] aVerbs = COM.Verbs(sFile);
bool bFound = false;
foreach (string s in aVerbs) {
if (s.Contains("pen Wit")) bFound = true;
if (bFound) break;
} // foreach s
if (!bFound) {
Array.Resize(ref aVerbs, aVerbs.Length + 1);
aVerbs[aVerbs.Length - 1] = "Open With...";
}

string[] aNames = new string[aVerbs.Length];
for (int iVerb = 0; iVerb < aVerbs.Length; iVerb++) aNames[iVerb] = aVerbs[iVerb].Replace("&", "");

string sName = Dialog.Pick("Context Menu", aNames, true, 0);
if (sName.Length == 0) return;

int i = Array.IndexOf(aNames, sName);
string sVerb = aVerbs[i];

// Clipboard.SetText(sVerb);
// if (sVerb.Replace("&", "") == "Open With...") Win32.OpenWith(sFile);
// if (sVerb.Replace("&", "") == "Open With...") Process.Start("Rundll32.exe", "shell32.dll, OpenAs_RunDLL " + Util.Quote(sFile));
if (sVerb.Replace("&", "") == "Open With...") Util.Run("Rundll32.exe shell32.dll, OpenAs_RunDLL " + sFile);
else COM.InvokeVerb(sFile, sVerb);
} // ContextMenu method

public void SendToMenu(string sFile) {
MdiChild child = this.Child;
string sDir = Environment.GetFolderPath(Environment.SpecialFolder.SendTo);
string[]aLinks = Directory.GetFiles(sDir);
string sNameList = "";
foreach (string s in aLinks) sNameList += Path.GetFileNameWithoutExtension(s) + "\n";
string[]aNames = sNameList.Trim().Split('\n');
string sName = Dialog.Pick("SendTo Menu", aNames, true, 0);
if (sName.Length == 0) return;

int i = Array.IndexOf(aNames, sName);
string sLink = aLinks[i];

Process.Start(sLink, sFile);
} // SendToMenu method

public void ListBox_KeyUp(Object sender, KeyEventArgs e) {
ListBox lst = (ListBox) sender;
bool bChecked = false;
if (lst is CheckedListBox) bChecked = true;

if(e.KeyCode == Keys.Space && !e.Alt && !e.Control && e.Shift) {
if (bChecked) {
foreach (int i in ((CheckedListBox) sender).CheckedIndices) {
//Util.Say(lst.Items[i].ToString());
Util.Say(i);
}
e.Handled = true;
}
}
else if(e.KeyCode == Keys.J && ((e.Alt && !e.Control) || (!e.Alt && e.Control)) && !e.Shift) {
string sText = Dialog.Jump;
if (e.Control) {
sText = Dialog.Input("Jump", "Text", sText, "Jump");
if (sText.Length == 0) return;
}

int iIndex = lst.SelectedIndex;
if (e.Alt || sText == Dialog.Jump) iIndex++;
else iIndex = 0;
Dialog.Jump = sText;

int iCount = lst.Items.Count;
//while (iIndex < iCount && lst.Items[iIndex].ToString().ToLower().IndexOf(sText) == -1) iIndex ++;
while (iIndex < iCount && lst.Items[iIndex].ToString().ToLower().IndexOf(sText) == -1) {
//Util.Say(iIndex);
iIndex ++;
}
if (iIndex < iCount) lst.SelectedIndex = iIndex;
else AddMessage("Not found!");
//lst.Update();
e.Handled = true;
}
else e.Handled = false;
} // ListBox_KeyUp handler

} // MdiFrame class

public class HomerRichTextBox : RichTextBox {
public int OldIndex = -1;
public int OldTextLength = -1;
public static string CR = "\r";
public static string LF = "\n";
public static string LB = LF;
public static string LineBreak = Environment.NewLine;
public static string FF = "\f";
public static string SB = FF + LB;
public static string DD = "----------";
public static string SectionBreak = LB + DD + LB + SB;
public static string EOD = LB + DD + LB + "End of Document" + LB;

public bool IndentMode = false;
public int IndentLevels = 0;
public int Index {
get {
return this.SelectionStart + this.SelectionLength;
}
set {
this.DeselectAll();
this.SelectionStart = value;
this.ScrollToCaret();
this.Update();
this.Refresh();
Application.DoEvents();
System.Threading.Thread.Sleep(100);
//this.OnNotifyMessage();
//this.OnSelectionChanged();
}
} // Index property

public int Row {
get {
return this.GetLineFromCharIndex(this.Index);
}
set {
int iIndex = this.GetFirstCharIndexFromLine(value);
this.DeselectAll();
this.SelectionStart = iIndex;
}
} // Row property

public int Col {
get {
return this.Index - this.GetFirstCharIndexOfCurrentLine();
}
set {
this.Index = GetFirstCharIndexOfCurrentLine() + value;
}
} // Col property

public int RowStart {
get {
return this.GetFirstCharIndexOfCurrentLine();
}
set {
}
} // RowStart property

public string RowText {
get {
return this.GetRowText(this.Row);
}
set {
}
} // RowText property

public int RowEnd {
get {
return this.RowStart + this.RowText.Length;
}
set {
}
} // RowEnd property

public int Line {
get {
return this.Row + 1;
}
set {
this.Row = value - 1;
}
} // Line property

public int Column {
get {
return this.Col + 1;
}
set {
this.Col = value - 1;
}
} // Column property

public double Percent {
get {
if (this.Text.Length == 0) return 0;
else return Math.Round((double) ((100.0 * this.Index) / this.Text.Length), 1);
}
set {
int iIndex = (int) ((this.Text.Length * value) / 100.0);
this.DeselectAll();
this.SelectionStart = iIndex;
}
} // Percent property

public void SetRowAndCol(int iRow, int iCol) {
int iRowStart = this.GetFirstCharIndexFromLine(iRow);
int iIndex = iRowStart + iCol;
this.DeselectAll();
this.SelectionStart = iIndex;
} // SetRowAndCol method

public void SetLineAndColumn(int iLine, int iColumn) {
int iRow = iLine - 1;
int iCol = iColumn - 1;
this.SetRowAndCol(iRow, iCol);
} // SetLineAndColumn method

public string GetRange(int iStart, int iEnd) {
int iLength = iEnd - iStart ;
string sText = this.Text;
return sText.Substring(iStart, iLength);
} // GetRange method

public void ReplaceRange(int iStart, int iEnd, string sText) {
this.DeselectAll();
this.Select(iStart, iEnd - iStart);
this.SelectedText = sText;
this.Index = iStart + sText.Length;
} // ReplaceRange method

public void SelectRange(int iStart, int iEnd) {
this.DeselectAll();
this.Select(iStart, iEnd - iStart);
} // SelectRange method

private int iStartSelection;
public int StartSelection {
get {
return iStartSelection;
}
set {
iStartSelection = value;
}
} // StartSelection property

private int iBookmark;
public int Bookmark {
get {
return iBookmark;
}
set {
iBookmark = value;
}
} // Bookmark property

private string sFindText;
public string FindText {
get {
return sFindText;
}
set {
sFindText = value;
}
} // FindText property

private string sMatchText;
public string MatchText {
get {
return sMatchText;
}
set {
sMatchText = value;
}
} // MatchText property

private string sReplaceText;
public string ReplaceText {
get {
return sReplaceText;
}
set {
sReplaceText = value;
}
} // ReplaceText property

private string sPatternText;
public string PatternText {
get {
return sPatternText;
}
set {
sPatternText = value;
}
} // PatternText property

private string sSubstituteText;
public string SubstituteText {
get {
return sSubstituteText;
}
set {
sSubstituteText = value;
}
} // SubstituteText property

private string sJumpLine;
public string JumpLine {
get {
return sJumpLine;
}
set {
sJumpLine = value;
}
} // JumpLine property

private string sGoPercent;
public string GoPercent {
get {
return sGoPercent;
}
set {
sGoPercent = value;
}
} // GoPercent property

private string sSearchTopic;
public string SearchTopic {
get {
return sSearchTopic;
}
set {
sSearchTopic = value;
}
} // SearchTopic property

private int iOldSelectionStart;
public int OldSelectionStart {
get {
return iOldSelectionStart;
}
set {
iOldSelectionStart = value;
}
} // OldSelectionStart property

private int iOldSelectionLength;
public int OldSelectionLength {
get {
return iOldSelectionLength;
}
set {
iOldSelectionLength = value;
}
} // OldSelectionLength property

public void StoreSelection() {
this.OldSelectionStart = this.SelectionStart;
this.OldSelectionLength = this.SelectionLength;
this.DeselectAll();
this.Index = this.OldSelectionStart + this.OldSelectionLength;
} // StoreSelection method

public void Reselect() {
this.DeselectAll();
this.Select(this.OldSelectionStart, this.OldSelectionLength);
} // Reselect method

public bool IsBottomRow {
get {
int iIndex = GetFirstCharIndexFromLine(this.Row + 1);
return iIndex < 0;
}
set {
}
} // IsBottomRow property

public int BottomRow {
get {
return this.Text.Split('\n').Length - 1;
}
set {
}
} // BottomRow property

public int RowLength {
get {
int iRow = this.Row;
int iStart = GetFirstCharIndexFromLine(iRow);
int iEnd = GetFirstCharIndexFromLine(iRow + 1);
//if (iEnd <= 0) iEnd = iStart;
if (iEnd <= 0) iEnd = this.TextLength;
;
int iLength = iEnd - iStart;
/*
int iLength = this.Lines[this.Row].Length;
if (!this.IsBottomRow) iLength++;
*/
return iLength;
}
set {
}
} // RowLength property

public HomerRichTextBox() {
SectionBreak = App.ReadOption("SectionBreak", SectionBreak);
string s = App.ReadOption("UseIndentModeDefault", "N").Trim().ToUpper();
if (s == "Y" || s == "YES") this.IndentMode = true;
Ini.WriteValue(App.IniFile, "Data", "IndentMode", (this.IndentMode ? "1" : "0"), false);
} // HomerRichTextBox constructor

public int GetIndexRow(int iIndex) {
return this.GetLineFromCharIndex(iIndex);
} // GetIndexRow method

public int GetRowStart(int iRow) {
return this.GetFirstCharIndexFromLine(iRow);
} // GetRowStart method

public string GetRowText(int iRow) {
int iStart = this.GetFirstCharIndexFromLine(iRow);
int iEnd = this.GetFirstCharIndexFromLine(iRow + 1);
// Dialog.Show("iStart " + iStart, "iEnd " + iEnd);
if (iEnd == -1) iEnd = this.Text.Length;
else iEnd --;
return this.GetRange(iStart, iEnd);
} // GetRowText method

public bool SetWrap(bool bWrap) {
bool bOldWrap = this.WordWrap;
bool bModified = this.Modified;
this.WordWrap = bWrap;
this.Modified = bModified;
return bOldWrap;
} // SetWrap method

public bool SetGuard(bool bGuard) {
bool bOldGuard = this.ReadOnly;
bool bModified = this.Modified;
this.ReadOnly = bGuard;
this.Modified = bModified;
return bOldGuard;
} // SetGuard method

protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
return App.Frame.ProcessCmdKey_Helper(ref msg, keyData);
} // ProcessCmdKey handler

} // HomerRichTextBox class

public class ListForm : Form {

public ListBox lst;
public DataTable tbl;
public BindingSource bs ;
public string Filter;
public DataTable tblDefault = null;
public int CheckFirst = -1;
public int CheckLast = -1;

protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
ListBox lst = this.lst;
bool bChecked = false;
if (lst is CheckedListBox) bChecked = true;

switch (keyData) {
case Keys.Alt | Keys.A :
App.Frame.AddMessage("Alpha order");
bs.Sort = "Item asc";
bs.Position = 0;
return true;
case Keys.Alt | Keys.Shift | Keys.A :
App.Frame.AddMessage("Reverse alpha order");
bs.Sort = "Item desc";
bs.Position = 0;
return true;
case Keys.Alt | Keys.D :
App.Frame.AddMessage("Default order");
if (this.tblDefault == null) {
this.tblDefault = new DataTable();
this.tblDefault.Columns.Add("Item", typeof(string));
this.tblDefault.Columns.Add("Value", typeof(string));
for (int i = 0; i < tbl.Rows.Count; i++)  this.tblDefault.Rows.Add(tbl.Rows[i][0].ToString(), tbl.Rows[i][1].ToString());
}

tbl = this.tblDefault;
bs.Sort = "";
bs.Position = 0;
return true;
case Keys.Alt | Keys.Shift | Keys.D :
App.Frame.AddMessage("Reverse default order");
if (this.tblDefault == null) {
this.tblDefault = new DataTable();
this.tblDefault.Columns.Add("Item", typeof(string));
this.tblDefault.Columns.Add("Value", typeof(string));
for (int i = 0; i < tbl.Rows.Count; i++)  this.tblDefault.Rows.Add(tbl.Rows[i][0].ToString(), tbl.Rows[i][1].ToString());
}

DataTable tblNew = new DataTable();
tblNew.Columns.Add("Item", typeof(string));
tblNew.Columns.Add("Value", typeof(string));
for (int i = this.tblDefault.Rows.Count -1; i >= 0; i--) tblNew.Rows.Add(this.tblDefault.Rows[i][0].ToString(), tblDefault.Rows[i][1].ToString());
tbl = tblNew;
//bs = new BindingSource();
bs.DataSource = tbl;
//this.BS = bs;
//bs.ResetBindings();
bs.Sort = "";
bs.Position = 0;
return true;
case Keys.Alt | Keys.Delete :
App.Frame.AddMessage((bs.Position + 1) + " of " + tbl.DefaultView.Count);
return true;
case Keys.Shift | Keys.Space :
if (bChecked) {
int iChecked = ((CheckedListBox) lst).CheckedItems.Count;
if (iChecked == 0) App.Frame.AddMessage("No items checked!");
else App.Frame.AddMessage("Checked" + iChecked);
List<int> listChecked = new List<int>();
foreach (int i in ((CheckedListBox) lst).CheckedIndices) listChecked.Add(i);
listChecked.Sort();
foreach (int i in listChecked) App.Frame.AddMessage(tbl.DefaultView[i][0].ToString());
}
else {
App.Frame.AddMessage("Selected");
foreach (int i in lst.SelectedIndices) App.Frame.AddMessage(tbl.DefaultView[i][0].ToString());
}
return true;
case Keys.Space :
//if (!bChecked || this.ActiveControl is Button) return base.ProcessCmdKey (ref msg, keyData);
if (!bChecked || !(this.ActiveControl is ListBox)) return base.ProcessCmdKey (ref msg, keyData);

{
int i = bs.Position;
bool b = ((CheckedListBox) lst).GetItemChecked(i);
((CheckedListBox) lst).SetItemChecked(i, !b);
return true;
}
case Keys.Control | Keys.Home :
if (!bChecked || !(this.ActiveControl is ListBox)) return base.ProcessCmdKey (ref msg, keyData);

int iStart = -1;
for (int i = 0; i < tbl.DefaultView.Count; i++) {
if (((CheckedListBox) lst).GetItemChecked(i)) {
iStart = i;
break;
}
}

if (iStart >= 0) bs.Position = iStart;
else App.Frame.AddMessage("Not found!");
return true;
case Keys.Control | Keys.End :
if (!bChecked || !(this.ActiveControl is ListBox)) return base.ProcessCmdKey (ref msg, keyData);

int iEnd = -1;
for (int i = tbl.DefaultView.Count - 1; i >= 0; i--) {
if (((CheckedListBox) lst).GetItemChecked(i)) {
iEnd = i;
break;
}
}

if (iEnd >= 0) bs.Position = iEnd;
else App.Frame.AddMessage("Not found!");
return true;
case Keys.Control | Keys.Down :
if (!bChecked || !(this.ActiveControl is ListBox)) return base.ProcessCmdKey (ref msg, keyData);

int iNext = -1;
for (int i = bs.Position + 1; i < tbl.DefaultView.Count; i++) {
if (((CheckedListBox) lst).GetItemChecked(i)) {
iNext = i;
break;
}
}

if (iNext >= 0) bs.Position = iNext;
else App.Frame.AddMessage("Not found!");
return true;
case Keys.F8 :
case Keys.Shift | Keys.F8 :
case Keys.Alt | Keys.Shift | Keys.F8 :
case Keys.Shift | Keys.Clear :
case Keys.Alt | Keys.Shift | Keys.Clear :
case Keys.Shift | Keys.Down :
case Keys.Alt | Keys.Shift | Keys.Down :
case Keys.Shift | Keys.Up :
case Keys.Alt | Keys.Shift | Keys.Up :
case Keys.Shift | Keys.End :
case Keys.Alt | Keys.Shift | Keys.End :
case Keys.Shift | Keys.Home :
case Keys.Alt | Keys.Shift | Keys.Home :
if (!bChecked || !(this.ActiveControl is ListBox)) return base.ProcessCmdKey (ref msg, keyData);

bool bState;
int iFirst, iLast;
int iAfter = bs.Position;
string sKey = Util.Key2String(keyData);

if (keyData == Keys.F8) {
App.Frame.AddMessage("Start Check or Uncheck");
this.CheckFirst = iAfter;
return true;
}
else if (keyData == (Keys.Shift | Keys.F8)) {
App.Frame.AddMessage("Complete Check");
bState = true;
iFirst = this.CheckFirst;
iLast = iAfter;
}
else if (keyData == (Keys.Alt | Keys.Shift | Keys.F8)) {
App.Frame.AddMessage("Complete Uncheck");
bState = false;
iFirst = this.CheckFirst;
iLast = iAfter;
}
else {
if (sKey.IndexOf("Alt+") >= 0) bState = false;
else bState = true;

if (sKey.IndexOf("+End") >= 0) {
iLast = tbl.DefaultView.Count - 1;
iAfter = iLast;
}
else iLast = iAfter;

if (sKey.IndexOf("+Home") >= 0) {
iFirst = 0;
iAfter = iFirst;
}
else iFirst = iAfter;

if (sKey.IndexOf("+Up") >= 0) iAfter--;
if (sKey.IndexOf("+Down") >= 0) iAfter++;

}

if (iFirst > iLast) Util.Swap(ref iFirst, ref iLast);
for (int iPosition = iFirst; iPosition <= iLast; iPosition ++) ((CheckedListBox) lst).SetItemChecked(iPosition, bState);
if (iAfter != bs.Position && iAfter >=0 && iAfter < tbl.DefaultView.Count) bs.Position = iAfter;
return true;
case Keys.Control | Keys.Up :
if (!bChecked || !(this.ActiveControl is ListBox)) return base.ProcessCmdKey (ref msg, keyData);

int iPrevious = -1;
for (int i = bs.Position - 1; i >= 0; i--) {
if (((CheckedListBox) lst).GetItemChecked(i)) {
iPrevious = i;
break;
}
}

if (iPrevious >= 0) bs.Position = iPrevious;
else App.Frame.AddMessage("Not found!");
return true;
case Keys.Control | Keys.A :
if (!bChecked || !(this.ActiveControl is ListBox)) return base.ProcessCmdKey (ref msg, keyData);

if (bChecked) {
App.Frame.AddMessage("Check All");
for (int i = 0; i < tbl.DefaultView.Count; i++) ((CheckedListBox) lst).SetItemChecked(i, true);
}
return true;
case Keys.Control | Keys.Shift | Keys.A :
if (!bChecked || !(this.ActiveControl is ListBox)) return base.ProcessCmdKey (ref msg, keyData);

if (bChecked) {
App.Frame.AddMessage("Uncheck All");
for (int i = 0; i < tbl.DefaultView.Count; i++) ((CheckedListBox) lst).SetItemChecked(i, false);
}
return true;
case Keys.Control | Keys.F :
case Keys.Control | Keys.Shift | Keys.F :
string sFilterSql = "";
string sFilter = "";
if (keyData == (Keys.Control | Keys.Shift | Keys.F)) App.Frame.AddMessage("Clear filter");
else {
Dialog.hashFilter.TryGetValue(this.Text, out sFilter);
sFilter = Dialog.Input("Filter", "Text", sFilter);
if (sFilter.Length == 0) return true;
//sFilterSql = "Item like '" + sFilter + "'";
sFilterSql = GetFilterSql(sFilter);
}

string sTemp = bs.Filter;
try {
bs.Filter = sFilterSql;
this.Filter = sFilter;
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
bs.Filter = sTemp;
return true;
}

bs.Position = 0;
App.Frame.AddMessage(Util.Pluralize(bs.Count, "item"));

//if (keyData == (Keys.Control | Keys.F)) {
if (Dialog.hashFilter.ContainsKey(this.Text)) Dialog.hashFilter.Remove(this.Text);
if (sFilter.Trim().Length > 0) Dialog.hashFilter.Add(this.Text, sFilter);
//}
return true;
case Keys.Control | Keys.J :
case Keys.Alt | Keys.J :
string sTitle = this.Text;
string sJump = "";
Dialog.hashJump.TryGetValue(sTitle, out sJump);
if (keyData == (Keys.Control | Keys.J)) {
sJump = Dialog.Input("Jump", "Text", sJump);
if (sJump.Length == 0) return true;
}

int iIndex = bs.Position;
if (keyData == (Keys.Alt | Keys.J) || sJump == Dialog.Jump) iIndex++;
else iIndex = 0;
if (Dialog.hashJump.ContainsKey(sTitle)) Dialog.hashJump.Remove(sTitle);
Dialog.hashJump.Add(sTitle, sJump);

int iCount = tbl.DefaultView.Count;
//while (iIndex < iCount && tbl.DefaultView[iIndex].ToString().ToLower().IndexOf(sJump) == -1) iIndex ++;
/*
while (iIndex < iCount && tbl.DefaultView[iIndex].ToString().ToLower().IndexOf(sJump) == -1) {
//App.Frame.AddMessage(iIndex);
iIndex ++;
}
*/
while (iIndex < iCount && tbl.DefaultView[iIndex][0].ToString().ToLower().IndexOf(sJump) == -1) {
iIndex ++;
}
//if (iIndex < iCount) bs.Position = iIndex;
if (iIndex < iCount) bs.Position = iIndex;
else App.Frame.AddMessage("Not found!");
//lst.Update();
return true;
}

return base.ProcessCmdKey (ref msg, keyData);
} // ProcessCmdKey handler

public string GetFilterSql(string sText) {
if (sText == null) sText = "";
sText = sText.Trim();
if (sText == "" || sText == "*") return "";
string[] aFilters = sText.Split('|');
string s = "";
for (int i =0; i < aFilters.Length; i++) {
if (i == 0) s += "(";
string[] a = aFilters[i].Split('*');
for (int j = 0; j < a.Length; j++) {
string sPrefix = "";
string sSuffix = "";
if (j == 0) s += " (";
if (a[j].Length > 0) {
if (j > 0) sPrefix = "*";
if (j < a.Length - 1) sSuffix = "*";
s += "Item like '" + sPrefix + a[j] + sSuffix + "'";
}

if (j == a.Length - 1) s += ") ";
else s += " and ";
}
if (i == aFilters.Length - 1) s+=")";
else s += " or ";
}

s = s.Replace("( and ", "(");
s = s.Replace(" and )", ")");
s = s.Replace("**", "*");
s = s.Replace("  ", " ");
s = s.Replace("( ", "(");
s = s.Replace(" )", ")");
s = s.Trim();
return s;
} // GetFilterSql method

} // ListForm class

public class Dialog {
public static string Jump = "";
public static Dictionary<string, string> hashItem = new Dictionary<string, string>();
public static Dictionary<string, string> hashFilter = new Dictionary<string, string>();
public static Dictionary<string, string> hashSort = new Dictionary<string, string>();
public static Dictionary<string, string> hashJump = new Dictionary<string, string>();

public static int PickEncoding(string sTitle, int iDefault) {
EncodingInfo[] eis = Encoding.GetEncodings();
int iCount = eis.Length;
string[] aNames = new string[iCount];
int[] aCodes = new int[iCount];
for (int i = 0; i < iCount; i++) {
EncodingInfo ei = eis[i];
Encoding en = ei.GetEncoding();
aNames[i] = en.EncodingName + " = " + en.CodePage;
aCodes[i] = en.CodePage;
}

Array.Sort(aNames, aCodes);
int iPosition = Array.IndexOf(aCodes, Encoding.Default.CodePage);
if (iPosition == -1) iPosition = 0;

if (sTitle.Length == 0) sTitle = "Pick Encoding";
string sItem = "";
if (hashItem.TryGetValue(sTitle, out sItem)) iPosition = 0;

string sName = Dialog.Pick(sTitle, aNames, false, iPosition);
if (sName.Length == 0) return -1;

iPosition = Array.IndexOf(aNames, sName);
int iCodePage = aCodes[iPosition];
return iCodePage;
} // PickEncoding method

public static string OpenFile(string sTitle, string sPath) {
string sReturn = "";
string sDir;

OpenFileDialog dlg = new OpenFileDialog();
if (sTitle.Length > 0) dlg.Title = sTitle;
if (File.Exists(sPath)) {
dlg.FileName = sPath;
sDir = Path.GetDirectoryName(sPath);
}
else sDir = sPath;

if (!Directory.Exists(sDir)) sDir = Directory.GetCurrentDirectory();
dlg.InitialDirectory = sDir;

string sFilter = "All files (*.*)|*.*|Text files (*.txt)|*.txt|Rich Text Format files (*.rtf)|*.rtf";
string sCompiler = App.ReadData("Compiler", "Default");
string sExtensionDefault = App.ReadOption("ExtensionDefault", "");
if (sCompiler != "Default") sFilter = sCompiler + " files (*." + sExtensionDefault + ")|*." + sExtensionDefault + "|" + sFilter;
dlg.Filter = sFilter;
dlg.FilterIndex = 1;
dlg.ValidateNames = true;
dlg.CheckPathExists = true;

if (dlg.ShowDialog() == DialogResult.OK) sReturn = dlg.FileName;
dlg.Dispose();
return sReturn;
} // OpenFile method

public static string SaveFile(string sTitle, string sPath) {
string sReturn = "";
string sDir;

SaveFileDialog dlg = new SaveFileDialog();
if (sTitle.Length > 0) dlg.Title = sTitle;
if (Directory.Exists(sPath)) sDir = sPath;
else {
dlg.FileName = sPath;
sDir = Path.GetDirectoryName(sPath);
}

if (Directory.Exists(sDir)) dlg.InitialDirectory = sDir;

string sFilter = "All files (*.*)|*.*|Text files (*.txt)|*.txt|Rich Text Format files (*.rtf)|*.rtf";
string sCompiler = App.ReadData("Compiler", "Default");
string sExtensionDefault = App.ReadOption("ExtensionDefault", "");
if (sCompiler != "Default") sFilter = sCompiler + " files (*." + sExtensionDefault + ")|*." + sExtensionDefault + "|" + sFilter;
dlg.Filter = sFilter;
dlg.FilterIndex = 1;
dlg.CheckPathExists = true;
dlg.SupportMultiDottedExtensions = true;

dlg.CreatePrompt = false;
dlg.ValidateNames = true;
dlg.AddExtension = true;
//dlg.AddExtension = false;
//dlg.DefaultExt = "txt";
dlg.DefaultExt = App.ReadOption("ExtensionDefault", "");

if (dlg.ShowDialog() == DialogResult.OK) sReturn = dlg.FileName;
dlg.Dispose();
return sReturn;
} // SaveFile method

public static string OldInput(string sTitle, string sLabel, string sValue) {
return Input(sTitle, sLabel, sValue);
} // Input method

public static string Input(string sTitle, string sLabel, string sValue) {
string[] aLabel = new string[] {sLabel};
string[] aValue = new string[] {sValue};
string[] aReturn = MultiInput(sTitle, aLabel, aValue);
//string sReturn = aReturn[0];
string sReturn = "";
if (aReturn != null && aReturn.Length > 0) sReturn = aReturn[0];
return sReturn;
} // Input method

public static string[] MultiInput(string sTitle, string[] aLabel, string[] aValue) {
LbcDialog dlg = new LbcDialog(sTitle, App.Frame);
List<TextBox> lBoxes = new List<TextBox>();
for (int i = 0; i < aLabel.Length; i++) {
string sVal = (aValue != null && i < aValue.Length && aValue[i] != null) ? aValue[i] : "";
lBoxes.Add(dlg.addInputBox(aLabel[i], sVal));
}
List<string> lReturn = new List<string>();
if (dlg.runOkCancel()) foreach (TextBox txt in lBoxes) lReturn.Add(txt.Text);
dlg.Dispose();
return lReturn.ToArray();
} // MultiInput method

// Input with input history: when sHistoryKey is given (for example "Find"),
// the field is an editable combo box whose dropdown offers up to
// historyCount recent entries for the command, newest first. Delegates to
// the history-aware MultiInput.
public static string Input(string sTitle, string sLabel, string sValue, string sHistoryKey) {
if (String.IsNullOrEmpty(sHistoryKey)) return Input(sTitle, sLabel, sValue);
string[] aLabel = new string[] {sLabel};
string[] aValue = new string[] {sValue};
string[] aKey = new string[] {sHistoryKey};
string[] aReturn = MultiInput(sTitle, aLabel, aValue, aKey);
string sReturn = "";
if (aReturn != null && aReturn.Length > 0) sReturn = aReturn[0];
return sReturn;
} // Input method (history overload)

// MultiInput with input history: aHistoryKey parallels aLabel; where a key
// is given, that field is an editable combo box whose dropdown offers up to
// historyCount recent entries for the command, newest first, persisted in
// section [Recent<key>] of EdSharp.ini as slot keys term1, term2, and so
// on. This is the same layout FileDir and DbDo use. [General] historyCount
// sets the depth; default 10, ceiling 100. Fields with a null or empty key
// keep the plain text box, so nothing is recorded for them.
public static string[] MultiInput(string sTitle, string[] aLabel, string[] aValue, string[] aHistoryKey) {
if (aHistoryKey == null) return MultiInput(sTitle, aLabel, aValue);
int iCount = Homer.InputHistory.clampCount(App.ReadValue("General", "historyCount", ""));
LbcDialog dlg = new LbcDialog(sTitle, App.Frame);
List<Control> lBoxes = new List<Control>();
for (int i = 0; i < aLabel.Length; i++) {
string sVal = (aValue != null && i < aValue.Length && aValue[i] != null) ? aValue[i] : "";
string sKey = (i < aHistoryKey.Length) ? aHistoryKey[i] : null;
if (String.IsNullOrEmpty(sKey)) { lBoxes.Add(dlg.addInputBox(aLabel[i], sVal)); continue; }
string sSection = "Recent" + sKey;
List<string> lRecent = Homer.InputHistory.load(delegate(string sSlot) { return App.ReadValue(sSection, sSlot, ""); }, iCount);
lBoxes.Add(dlg.addComboHistoryBox(aLabel[i], lRecent, sVal, ""));
}
List<string> lReturn = new List<string>();
if (dlg.runOkCancel()) {
for (int i = 0; i < lBoxes.Count; i++) {
string sText = lBoxes[i].Text;
lReturn.Add(sText);
string sKey = (i < aHistoryKey.Length) ? aHistoryKey[i] : null;
if (String.IsNullOrEmpty(sKey) || sText == null || sText.Trim().Length == 0) continue;
string sSection = "Recent" + sKey;
List<string> lRecent = Homer.InputHistory.load(delegate(string sSlot) { return App.ReadValue(sSection, sSlot, ""); }, iCount);
lRecent = Homer.InputHistory.push(lRecent, sText.Trim(), iCount);
Homer.InputHistory.store(lRecent, delegate(string sSlot, string sSlotValue) { App.WriteValue(sSection, sSlot, sSlotValue); }, iCount);
}
}
dlg.Dispose();
return lReturn.ToArray();
} // MultiInput method (history overload)

public static string Pick(string sTitle, string[] aValue, bool bSort) {
string[] aDisplay = null;
int iIndex = 0;
return Pick(sTitle, aValue, aDisplay, bSort, iIndex);
} // Pick method

public static string[] MultiPick(string sTitle, string[] aValues, int[] aSelect, bool bSort) {
List<string> lSelectedValues = new List<string>();
if (aSelect != null) foreach (int i in aSelect) if (i >= 0 && i < aValues.Length) lSelectedValues.Add(aValues[i]);
string[] aItems = (string[]) aValues.Clone();
if (bSort) Array.Sort(aItems, new CaseInsensitiveComparer());
List<int> lChecked = new List<int>();
foreach (string sVal in lSelectedValues) { int idx = Array.IndexOf(aItems, sVal); if (idx >= 0) lChecked.Add(idx); }
List<string> lNames = new List<string>(aItems);
LbcDialog dlg = new LbcDialog(sTitle, App.Frame);
CheckedListBox clb = dlg.addCheckListBox(lNames, lChecked, "");
List<string> lReturn = new List<string>();
if (dlg.runOkCancel()) foreach (int i in clb.CheckedIndices) lReturn.Add(aItems[i]);
dlg.Dispose();
return lReturn.ToArray();
} // MultiPick method

public static string[] MultiCheck(string sTitle, string[] aValues, int[] aSelect, bool bSort, int iIndex) {
string[] aDisplay = null;
return MultiCheck(sTitle, aDisplay, aValues, aSelect, bSort, iIndex);
} // MultiCheck method

public static string[] MultiCheck(string sTitle, string[] aDisplay, string[] aValues, int[] aSelect, bool bSort, int iIndex) {
string[] aVal = (string[]) aValues.Clone();
string[] aDisp = (aDisplay == null) ? (string[]) aVal.Clone() : (string[]) aDisplay.Clone();
if (bSort) {
if (aDisplay == null) { Array.Sort(aVal, new CaseInsensitiveComparer()); aDisp = (string[]) aVal.Clone(); }
else Array.Sort(aDisp, aVal);
}
List<int> lChecked = new List<int>();
if (aSelect != null) foreach (int i in aSelect) if (i >= 0 && i < aValues.Length) { int idx = Array.IndexOf(aVal, aValues[i]); if (idx >= 0) lChecked.Add(idx); }
List<string> lNames = new List<string>(aDisp);
LbcDialog dlg = new LbcDialog(sTitle, App.Frame);
CheckedListBox clb = dlg.addCheckListBox(lNames, lChecked, "");
List<string> lReturn = new List<string>();
if (dlg.runOkCancel()) foreach (int i in clb.CheckedIndices) if (i >= 0 && i < aVal.Length) lReturn.Add(aVal[i]);
dlg.Dispose();
return lReturn.ToArray();
} // MultiCheck method

public static string Pick(string sTitle, string[] aValue, bool bSort, int iIndex) {
string[] aDisplay = null;
return Pick(sTitle, aValue, aDisplay, bSort, iIndex);
} // Pick method

public static string Pick(string sTitle, string[] aValue, string[] aDisplay, bool bSort, int iIndex) {
string[] aVal = (string[]) aValue.Clone();
string[] aDisp = (aDisplay == null) ? (string[]) aVal.Clone() : (string[]) aDisplay.Clone();
if (bSort) {
if (aDisplay == null) { Array.Sort(aVal, new CaseInsensitiveComparer()); aDisp = (string[]) aVal.Clone(); }
else Array.Sort(aDisp, aVal);
}
List<string> lNames = new List<string>(aDisp);
LbcDialog dlg = new LbcDialog(sTitle, App.Frame);
ListBox lst = dlg.addListBox(lNames, "", "");
if (iIndex >= 0 && iIndex < lNames.Count) lst.SelectedIndex = iIndex;
string sReturn = "";
if (dlg.runOkCancel()) {
int i = lst.SelectedIndex;
if (i >= 0 && i < aVal.Length) sReturn = aVal[i];
}
dlg.Dispose();
return sReturn;
} // Pick method

public static string Confirm(string sTitle, string sText, string sDefault) {
MessageBoxDefaultButton defaultButton;
if (sDefault.ToLower() == "n") defaultButton = MessageBoxDefaultButton.Button2;
else defaultButton = MessageBoxDefaultButton.Button1;

switch (MessageBox.Show(sText, sTitle, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, defaultButton)) {
case DialogResult.Yes :
//Util.Say("Yes");
return "Y";
case DialogResult.No :
//Util.Say("No");
return "N";
}
/*Util.Say("Cancel");*/
return "";
} // Confirm method

public static void Show(object oText) {
Show("Show", oText);
} // Show method

public static void Show(object oTitle, object oText) {
string sTitle = oTitle.ToString();
string sText = oText.ToString();
if (oTitle is bool) sTitle = ((bool) oTitle) ? "true" : "false";
if (oText is bool) sText = (bool) oText ? "true" : "false";
MessageBox.Show(oText.ToString(), oTitle.ToString());
} // Show method

public static void Properties(string sPath) {
COM.InvokeVerb(sPath, "P&roperties");
COM.InvokeVerb(sPath, "Properties");
// Win32.ShellExecute("Properties", sPath);
} // Properties method

// The spell check dialog: a heading that names and spells the word, the
// sentence it sits in, one editable combo box of suggestions, and four
// plainly named buttons. Returns the button clicked and the text in the
// combo box. Escape returns Cancel.
// A multiline prompt: a labelled box that keeps its lines, with OK and
// Cancel after it in the tab order. Enter inside the box inserts a line
// break rather than submitting, which is what a person typing several
// sentences expects; Control+Enter submits from anywhere, the Homer
// convention. Returns an empty string when cancelled.
public static string Prompt(string sTitle, string sLabel, string sValue) {
LbcDialog dlg = new LbcDialog(sTitle, App.Frame);
TextBox tbPrompt = dlg.addMemoBox(sLabel, sValue, "Enter starts a new line; Control+Enter submits");
string sClicked = dlg.runWithButtons(new string[] {"&OK", "&Cancel"});
string sText = (tbPrompt == null) ? "" : tbPrompt.Text;
dlg.Dispose();
if (sClicked == null || sClicked.Replace("&", "") != "OK") return "";
return sText;
} // Prompt method

public static object[] SpellChoice(string sTitle, string sPrompt, string sContext, List<string> lsSuggestions) {
LbcDialog dlg = new LbcDialog(sTitle, App.Frame);
dlg.addLabel(sPrompt);
if (sContext.Length > 0) dlg.addTextLine("Context", sContext);
ComboBox cbSuggestions = dlg.addComboHistoryBox("Replace with", lsSuggestions, (lsSuggestions.Count > 0) ? lsSuggestions[0] : "", "Down and Up walk the suggestions; type to correct the word yourself");
// Homer Tools rule: every button has its OWN access key, preferably the
// initial letter of its first word. Change and Cancel would both claim
// C, so the action is named Replace instead -- as familiar a word for
// this and equally plain -- giving R, S, A, C, and the dialog's own
// Help on H, all distinct.
string sClicked = dlg.runWithButtons(new string[] {"&Replace", "&Skip", "&Add to Dictionary", "&Cancel"});
string sReplacement = (cbSuggestions == null) ? "" : cbSuggestions.Text;
dlg.Dispose();
string sButton;
if (sClicked == null || sClicked.Length == 0) sButton = "Cancel";
else sButton = sClicked.Replace("&", "");
return new object[] {sButton, sReplacement};
} // SpellChoice method

public static string Choose (string sTitle, string sText, string[] aButtons, int iDefault) {
LbcDialog dlg = new LbcDialog(sTitle, App.Frame);
if (sText != "") dlg.addLabel(sText);
List<string> lButtons = new List<string>(aButtons);
bool bHasCancel = false;
foreach (string sButton in aButtons) if (sButton.Replace("&", "").Equals("Cancel", StringComparison.OrdinalIgnoreCase)) bHasCancel = true;
if (!bHasCancel) lButtons.Add("Cancel");
string sClicked = dlg.runWithButtons(lButtons.ToArray());
dlg.Dispose();
string sResult = "";
foreach (string sButton in aButtons) if (sButton.Replace("&", "") == sClicked) { sResult = sButton; break; }
Util.Say(sResult.Replace("&", ""));
return sResult;
} // Choose method

public static object[] PickAndChoose(string sTitle, object[] aValue, string[] aDisplay, string[] aButton, bool bSort, int iIndex) {
object[] aVal = (object[]) aValue.Clone();
string[] aDisp;
if (aDisplay == null) {
aDisp = new string[aVal.Length];
for (int i = 0; i < aVal.Length; i++) aDisp[i] = (aVal[i] == null) ? "" : aVal[i].ToString();
}
else aDisp = (string[]) aDisplay.Clone();
if (bSort) Array.Sort(aDisp, aVal);
List<string> lNames = new List<string>(aDisp);
LbcDialog dlg = new LbcDialog(sTitle, App.Frame);
ListBox lst = dlg.addListBox(lNames, "", "");
if (iIndex >= 0 && iIndex < lNames.Count) lst.SelectedIndex = iIndex;
else if (lNames.Count > 0) lst.SelectedIndex = 0;
List<string> lButtons = new List<string>(aButton);
bool bHasCancel = false;
foreach (string sButton in aButton) if (sButton.Replace("&", "").Equals("Cancel", StringComparison.OrdinalIgnoreCase)) bHasCancel = true;
if (!bHasCancel) lButtons.Add("Cancel");
string sClicked = dlg.runWithButtons(lButtons.ToArray());
int iPicked = lst.SelectedIndex;
dlg.Dispose();
string sButtonResult = "";
foreach (string sButton in aButton) if (sButton.Replace("&", "") == sClicked) { sButtonResult = sButton; break; }
object[] aResult = {};
if (sButtonResult != "" && iPicked >= 0 && iPicked < aVal.Length) aResult = new object[] {aVal[iPicked], sButtonResult};
return aResult;
} // PickAndChoose method

public static object[] GetFont(Font font, Color color) {
//ColorDialog d = new ColorDialog();
//d.ShowDialog();
FontDialog dlg = new FontDialog();
dlg.FontMustExist = true;
dlg.ShowColor = true;
dlg.Font = font;
dlg.Color = color;
object[] aReturn = {};
if(dlg.ShowDialog() == DialogResult.OK) aReturn = new object[] {dlg.Font, dlg.Color};
dlg.Dispose();
return aReturn;
} // GetFont method

public static string OpenFolder(string sTitle, string sLabel, string sValue) {
string sResult = "";

Form frm = new Form();
frm.SuspendLayout();
frm.AutoSize = true;
frm.AutoSizeMode = AutoSizeMode.GrowAndShrink;

FlowLayoutPanel flpMain = new FlowLayoutPanel();
flpMain.SuspendLayout();
flpMain.AutoSize = true;
flpMain.AutoSizeMode  = AutoSizeMode.GrowAndShrink;
flpMain.FlowDirection = FlowDirection.TopDown;

FlowLayoutPanel flpInput = new FlowLayoutPanel();
flpInput.SuspendLayout();
flpInput.Anchor = AnchorStyles.None;
flpInput.AutoSize = true;
flpInput.AutoSizeMode  = AutoSizeMode.GrowAndShrink;
flpInput.FlowDirection = FlowDirection.LeftToRight;

Label lbl = new Label();
lbl.AutoSize = true;
lbl.Text = sLabel + ":";
TextBox txt = new TextBox();
//txt.ScrollBars = ScrollBars.None;
txt.Width *= 2;
txt.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
txt.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;
txt.Text = sValue;
txt.AccessibleName = lbl.Text.Replace("&", "");
txt.GotFocus += delegate(object o, EventArgs e) {
txt.SelectAll();
};

Button btnBrowse = new Button();
btnBrowse.Click += delegate(object o, EventArgs e) {
txt.Text = Dialog.BrowseForFolder("", sValue, false);
txt.Select();
};
btnBrowse.Text = "&Browse";
btnBrowse.AccessibleName = btnBrowse.Text.Replace("&", "");

flpInput.Controls.AddRange(new Control[] {lbl, txt, btnBrowse});
flpInput.ResumeLayout();

FlowLayoutPanel flpButtons = new FlowLayoutPanel();
flpButtons.SuspendLayout();
flpButtons.Anchor = AnchorStyles.None;
flpButtons.AutoSize = true;
flpButtons.AutoSizeMode  = AutoSizeMode.GrowAndShrink;
flpButtons.FlowDirection = FlowDirection.LeftToRight;

Button btnOK = new Button();
btnOK.Click += delegate(object o, EventArgs e) {
sResult = txt.Text.Trim();
if (sResult != "" && !Directory.Exists(sResult)) {
string sChoice = Dialog.Confirm("Confirm", "Cannot find folder\n" + sResult + "\nCreate it?", "Y");
if (sChoice == "Y") {
try {
DirectoryInfo di = new DirectoryInfo(sResult);
di.Create();
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
}
}
}
if (Directory.Exists(sResult)) frm.Close();
else {
txt.SelectAll();
txt.Select();
}
};

btnOK.Text = "OK";
btnOK.AccessibleName = btnOK.Text;

Button btnCancel = new Button();
btnCancel.Click += delegate(object o, EventArgs e) {
/*Util.Say("Cancel");*/ sResult = "";
frm.Close();
};
btnCancel.Text = "Cancel";
btnCancel.AccessibleName = btnCancel.Text;

flpButtons.Controls.AddRange(new Control[] {btnOK, btnCancel});
flpButtons.ResumeLayout();

flpMain.Controls.AddRange(new Control[] {flpInput, flpButtons});
flpMain.ResumeLayout();

frm.AcceptButton = btnOK;
frm.CancelButton = btnCancel;
frm.StartPosition = FormStartPosition.CenterParent;
frm.Text = sTitle;
frm.Controls.Add(flpMain);
frm.ResumeLayout();
frm.Shown += delegate(object sender, EventArgs e) {
Win32.SetForegroundWindow(frm.Handle);
};
frm.ShowDialog();
frm.Dispose();
return sResult;
} // GetDirectory method

public static string BrowseForFolder(string sTitle, string sDir) {
bool bNewFolder = false;
return BrowseForFolder(sTitle, sDir, bNewFolder);
} // BrowseForFolder method

public static string BrowseForFolder(string sTitle, string sDir, bool bNewFolder) {
string sReturn = "";
FolderBrowserDialog dlg = new FolderBrowserDialog();
dlg.Description = sTitle;
dlg.ShowNewFolderButton = bNewFolder;
//dlg.RootFolder = sRootFolder;
dlg.SelectedPath = sDir;

if (dlg.ShowDialog() == DialogResult.OK) sReturn = dlg.SelectedPath;
dlg.Dispose();
return sReturn;
} // BrowseForFolder method

public static string[] PickAndInputDialog(string sTitle, string sLblList, string[] aValues, string sLblInput, string sValue, bool bSort, int iIndex) {
string[] aResults = {};

Form frm = new Form();
frm.SuspendLayout();
frm.AutoSize = true;
frm.AutoSizeMode = AutoSizeMode.GrowAndShrink;

FlowLayoutPanel flpMain = new FlowLayoutPanel();
flpMain.SuspendLayout();
flpMain.AutoSize = true;
flpMain.AutoSizeMode  = AutoSizeMode.GrowAndShrink;
flpMain.FlowDirection = FlowDirection.TopDown;

FlowLayoutPanel flpInput = new FlowLayoutPanel();
flpInput.SuspendLayout();
flpInput.Anchor = AnchorStyles.None;
flpInput.AutoSize = true;
flpInput.AutoSizeMode  = AutoSizeMode.GrowAndShrink;
flpInput.FlowDirection = FlowDirection.LeftToRight;

Label lblList = new Label();
lblList.Text = sLblList + ":";
lblList.AccessibleName = lblList.Text.Replace("&", "");

ListBox lst = new ListBox();
if (bSort) lst.Sorted = true;
lst.Items.AddRange(aValues);
lst.SelectedIndex = iIndex;

Label lblInput = new Label();
lblInput.Text = sLblInput + ":";
lblInput.AccessibleName = lblInput.Text.Replace("&", "");
TextBox txt = new TextBox();
txt.Width *= 2;
txt.AccessibleName = lblInput.AccessibleName;
if (lblInput.Text.Contains("Password:")) txt.UseSystemPasswordChar = true;
txt.Text = sValue;

flpInput.Controls.AddRange(new Control[] {lblList, lst, lblInput, txt});
flpInput.ResumeLayout();

FlowLayoutPanel flpButtons = new FlowLayoutPanel();
flpButtons.SuspendLayout();
flpButtons.Anchor = AnchorStyles.None;
flpButtons.AutoSize = true;
flpButtons.AutoSizeMode  = AutoSizeMode.GrowAndShrink;
flpButtons.FlowDirection = FlowDirection.LeftToRight;

Button btnOK = new Button();
btnOK.Click += delegate(object o, EventArgs e) {
aResults = new string[] {
lst.Text, txt.Text
};
frm.Close();
};

btnOK.Text = "OK";
btnOK.AccessibleName = btnOK.Text;

Button btnCancel = new Button();
btnCancel.Click += delegate(object o, EventArgs e) {
/*Util.Say("Cancel");*/ frm.Close();
};
btnCancel.Text = "Cancel";
btnCancel.AccessibleName = btnCancel.Text;

flpButtons.Controls.AddRange(new Control[] {btnOK, btnCancel});
flpButtons.ResumeLayout();

flpMain.Controls.AddRange(new Control[] {flpInput, flpButtons});
flpMain.ResumeLayout();

frm.AcceptButton = btnOK;
frm.CancelButton = btnCancel;
frm.StartPosition = FormStartPosition.CenterParent;
frm.Text = sTitle;
frm.Controls.Add(flpMain);
frm.ResumeLayout();
frm.Shown += delegate(object sender, EventArgs e) {
Win32.SetForegroundWindow(frm.Handle);
};
frm.ShowDialog();
frm.Dispose();
return aResults;
} // PickAndInput method

} // Dialog class

// Script: late-bound bridge to EdSharp.dll, the JScript .NET host built
// from EdSharp.js. Loaded by path (not /reference) so the exe and the
// same-named dll do not collide at load time. The MethodInfo is cached
// after first use. run returns the script result string, or text that
// begins "ERROR: " on a compile or runtime fault in the snippet.
// PickItem: carries a display string together with the index of the value
// it represents, so a sorted pick-list can map the selected row back to the
// original value array. Replaces the former VB6 ListBox ItemData shim.
// ===== Speech subsystem (Say + UIA) moved to Say.cs (namespace Homer) =

// ===== Lbc dialog classes moved to Lbc.cs (portable, shared with DbDuo) =====

// ===== JAWS script installer (DbDo-style) ==================================
// Invoked by the installer's Finish-page option as
// "EdSharp.exe --install-jaws-settings". For every installed JAWS version
// (per-user %APPDATA%\Freedom Scientific\JAWS\<ver>\Settings\<lang>), copies
// EdSharp's JAWS settings family in and compiles homer.jss then EdSharp.jss
// (Homer first, since EdSharp.jss does Use "Homer.jsb"). scompile.exe for each
// version is found via HKLM\Software\Freedom Scientific\JAWS\<ver>\Target.
public static class JawsScripts {

static string findScompilePath(string sVersion) {
try {
using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Freedom Scientific\JAWS\" + sVersion)) {
if (key != null) {
string sTarget = key.GetValue("Target") as string;
if (!String.IsNullOrEmpty(sTarget)) {
string sCompile = Path.Combine(sTarget, "scompile.exe");
if (File.Exists(sCompile)) return sCompile;
}
}
}
}
catch {}
string sPf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
string sFallback = Path.Combine(sPf, @"Freedom Scientific\JAWS\" + sVersion + @"\scompile.exe");
if (File.Exists(sFallback)) return sFallback;
return null;
} // findScompilePath method

static void compileOne(string sScompile, string sLangPath, string sJss, StringBuilder sb, ref int iCompiled) {
if (String.IsNullOrEmpty(sScompile)) return;
try {
ProcessStartInfo psi = new ProcessStartInfo(sScompile, "\"" + sJss + "\"");
psi.WorkingDirectory = sLangPath;
psi.UseShellExecute = false;
psi.CreateNoWindow = true;
using (Process proc = Process.Start(psi)) {
proc.WaitForExit(15000);
string sJsb = Path.Combine(sLangPath, Path.GetFileNameWithoutExtension(sJss) + ".jsb");
if (proc.HasExited && File.Exists(sJsb)) iCompiled++;
else sb.AppendLine("WARN: compile may have failed: " + sJss + " in " + sLangPath);
}
}
catch (Exception ex) { sb.AppendLine("FAIL: compile " + sJss + ": " + ex.Message); }
} // compileOne method

public static string install(string sScriptsFolder, out int iCopied, out int iCompiled) {
iCopied = 0;
iCompiled = 0;
StringBuilder sb = new StringBuilder();
string sJawsRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Freedom Scientific\JAWS");
if (!Directory.Exists(sJawsRoot)) {
sb.AppendLine("JAWS does not appear to be installed for the current user.");
sb.AppendLine("(No folder at " + sJawsRoot + ")");
return sb.ToString();
}
// Settings family to place in each JAWS settings folder: {source, target}.
string[][] aFiles = new string[][] {
new string[] {"EdSharp.JSS", "EdSharp.jss"},
new string[] {"edsharp.jkm", "EdSharp.jkm"},
new string[] {"EdSharp.JCF", "EdSharp.jcf"},
new string[] {"EdSharp.jsd", "EdSharp.jsd"},
new string[] {"Homer.jsh", "Homer.jsh"},
new string[] {"MSAA.jsh", "MSAA.jsh"},
new string[] {"homer.jss", "homer.jss"},
new string[] {"homer.jsd", "homer.jsd"}
};
foreach (string sVersionPath in Directory.GetDirectories(sJawsRoot)) {
string sVersion = Path.GetFileName(sVersionPath);
string sSettingsPath = Path.Combine(sVersionPath, "Settings");
if (!Directory.Exists(sSettingsPath)) continue;
string sScompile = findScompilePath(sVersion);
foreach (string sLangPath in Directory.GetDirectories(sSettingsPath)) {
foreach (string[] aPair in aFiles) {
string sSrc = Path.Combine(sScriptsFolder, aPair[0]);
string sDst = Path.Combine(sLangPath, aPair[1]);
if (!File.Exists(sSrc)) continue;
try { File.Copy(sSrc, sDst, true); iCopied++; }
catch (Exception ex) { sb.AppendLine("FAIL: copy " + aPair[1] + ": " + ex.Message); }
}
compileOne(sScompile, sLangPath, "homer.jss", sb, ref iCompiled);
compileOne(sScompile, sLangPath, "EdSharp.jss", sb, ref iCompiled);
if (String.IsNullOrEmpty(sScompile)) sb.AppendLine("WARN: scompile.exe not found for JAWS " + sVersion + "; files placed but not compiled.");
else sb.AppendLine("JAWS " + sVersion + " / " + Path.GetFileName(sLangPath) + ": done");
}
}
return sb.ToString();
} // install method

} // JawsScripts class

// InixCodec: order-preserving INI/INIX reader-writer ported from DbDo.
// ===== InixCodec moved to Inix.cs (namespace Homer) =========================

public class PickItem {
public int iValue;
public string sText;

public PickItem(string sText, int iValue) {
this.sText = sText;
this.iValue = iValue;
} // PickItem constructor

public override string ToString() { return sText; }
} // PickItem class

// VB: convenience methods ported from the former VB.vb support module,
// translated to C# with late-bound COM through dynamic so no separate
// VB.dll is needed. The Office text extractors (Xls2Txt, Ppt2Txt) and the
// Word automation are legacy COM, and are candidates to be replaced by
// Pandoc in a later stage. LookupTerm and GetLinks drive Internet Explorer,
// which is removed from current Windows; they are kept only for source
// compatibility and will likely be retired.
public class VB {
private const int iMsoTextEffect = 15; // MsoShapeType.msoTextEffect
private const int iWordFormatText = 2; // WdSaveFormat.wdFormatText
private const int iXlTextFormat = 21; // XlFileFormat current-region text
private const string sFormFeed = "\f";

// ---- Excel ----
public static object ExcelOpen(object oXlss, string sFile) {
dynamic xlss = oXlss;
try { return xlss.Open(sFile, false, true); } // UpdateLinks, ReadOnly
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); return null; }
} // ExcelOpen method

public static void ExcelSaveAs(object oXls, string sFile, int iFileFormat) {
dynamic xls = oXls;
try { xls.SaveAs(sFile, iFileFormat); }
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); }
} // ExcelSaveAs method

public static void ExcelClose(object oXls) {
dynamic xls = oXls;
try { xls.Close(0); }
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); }
} // ExcelClose method

public static void ExcelQuit(object oApp) {
dynamic app = oApp;
try { app.Quit(); }
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); }
} // ExcelQuit method

// ---- PowerPoint ----
public static object PowerPointOpen(object oPpts, string sFile) {
dynamic ppts = oPpts;
try { return ppts.Open(sFile, true); } // ReadOnly
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); return null; }
} // PowerPointOpen method

public static void PowerPointSaveAs(object oPpt, string sFile, int iFileFormat) {
dynamic ppt = oPpt;
try { ppt.SaveAs(sFile); }
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); }
} // PowerPointSaveAs method

public static void PowerPointClose(object oPpt) {
dynamic ppt = oPpt;
try { ppt.Close(); }
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); }
} // PowerPointClose method

public static void PowerPointQuit(object oApp) {
dynamic app = oApp;
try { app.Quit(); }
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); }
} // PowerPointQuit method

// ---- Word ----
public static object WordOpen(object oDocs, string sFile, bool bAppVisible) {
dynamic docs = oDocs;
try { return docs.Open(sFile, false, false, false); } // ConfirmConversions, ReadOnly, AddToRecentFiles
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); return null; }
} // WordOpen method

public static void WordSaveAs(object oDoc, string sFile, int iFileFormat) {
dynamic doc = oDoc;
try { doc.SaveAs(sFile, iFileFormat, false, "", false); } // LockComments, Password, AddToRecentFiles
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); }
} // WordSaveAs method

public static void WordClose(object oDoc) {
dynamic doc = oDoc;
ClearNormalTemplate(doc.Application);
try { doc.Close(0); }
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); }
} // WordClose method

public static void ClearNormalTemplate(object oApp) {
dynamic app = oApp;
dynamic template = app.NormalTemplate;
template.Saved = true;
Marshal.ReleaseComObject(template);
} // ClearNormalTemplate method

public static void WordQuit(object oApp) {
dynamic app = oApp;
ClearNormalTemplate(app);
try { app.Quit(0); }
catch (Exception ex) { MessageBox.Show(ex.Message, "Error!"); }
} // WordQuit method

// ---- Office text extractors (legacy COM) ----
public static object Ppt2Txt(string sSource, string sTarget) {
dynamic app = null;
bool bGet = false;
int iAlerts = 0;
try { app = Marshal.GetActiveObject("PowerPoint.Application"); iAlerts = (int) app.DisplayAlerts; bGet = true; }
catch { app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application")); bGet = false; }
app.Visible = true; // must be visible for COM automation
app.DisplayAlerts = false;
dynamic ppts = app.Presentations;
dynamic ppt = PowerPointOpen(ppts, sSource);

string s = ppt.Name;
string sText = Path.GetFileNameWithoutExtension(s);
dynamic slides = ppt.Slides;
int iSlideCount = (int) slides.Count;
sText = sText + "\r\n" + iSlideCount.ToString() + " Slide" + (iSlideCount == 1 ? "" : "s");

int iSlide = 1;
while (iSlide <= iSlideCount) {
dynamic slide = slides.Item(iSlide);
sText = sText + "\r\n" + "\r\n" + "----------" + "\r\n" + sFormFeed + "\r\n" + "Slide " + iSlide.ToString();

dynamic notes = slide.NotesPage;
int iNoteCount = (int) notes.Count;
bool bNoteLabel = true;
int iNote = 1;
while (iNote <= iNoteCount) {
dynamic note = notes.Item(iNote);
dynamic ships = note.Shapes;
int iShipCount = (int) ships.Count;
int iShip = 1;
while (iShip <= iShipCount) {
dynamic ship = ships.Item(iShip);
if ((int) ship.HasTextFrame != 0) {
dynamic frame = ship.TextFrame;
dynamic text = frame.TextRange;
s = text.Text;
if (s != "") {
if (bNoteLabel) { sText = sText + "\r\n" + "Notes:" + "\r\n" + s; bNoteLabel = false; }
else sText = sText + "\r\n" + s;
}
}
iShip = iShip + 1;
}
sText = sText.Trim();
iNote = iNote + 1;
}

dynamic shapes = slide.Shapes;
int iShapeCount = (int) shapes.Count;
bool bOutlineLabel = true;
int iShape = 1;
while (iShape <= iShapeCount) {
dynamic shape = shapes.Item(iShape);
s = "";
if ((int) shape.HasTextFrame != 0) {
dynamic textFrame = shape.TextFrame;
dynamic textRange = textFrame.TextRange;
s = textRange.Text;
if (s != "" && s.ToLower() != "outline") {
if (bOutlineLabel) { sText = sText + "\r\n" + "Outline:" + "\r\n" + s; bOutlineLabel = false; }
else sText = sText + "\r\n" + s;
}
}
if ((int) shape.HasTextFrame == 0 || s == "") {
s = shape.AlternativeText;
if (s != "") sText = sText + "\r\n" + s;
int iType = (int) shape.Type;
if (iType == iMsoTextEffect) {
dynamic textEffect = shape.TextEffect;
s = textEffect.Text;
if (s != "" && s != (string) shape.AlternativeText) sText = sText + "\r\n" + "Text Effect: " + s;
}
}
sText = sText.Trim();
iShape = iShape + 1;
}
sText = sText.Trim();
iSlide = iSlide + 1;
}

File.WriteAllText(sTarget, sText);
ppt.Saved = true;
PowerPointClose(ppt);
Marshal.ReleaseComObject(ppt);
Marshal.ReleaseComObject(ppts);
if (bGet) { app.DisplayAlerts = iAlerts; }
else PowerPointQuit(app);
Marshal.ReleaseComObject(app);
return File.Exists(sTarget);
} // Ppt2Txt method

public static bool Xls2Txt(string sSource, string sTarget) {
dynamic app = null;
bool bGet = false;
int iAlerts = 0, iUpdating = 0, iVisible = 0;
try {
app = Marshal.GetActiveObject("Excel.Application");
bGet = true;
iVisible = (int) app.Visible;
iAlerts = (int) app.DisplayAlerts;
iUpdating = (int) app.ScreenUpdating;
}
catch { app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")); bGet = false; }
app.Visible = false;
app.DisplayAlerts = false;
app.ScreenUpdating = false;
dynamic xlss = app.Workbooks;
dynamic xls = ExcelOpen(xlss, sSource);
dynamic sheets = xls.Sheets;
int iSheetCount = (int) sheets.Count;
int iSheet = 1;
string sBook = "";
while (iSheet <= iSheetCount) {
dynamic sheet = sheets.Item(iSheet);
if (File.Exists(sTarget)) File.Delete(sTarget);
ExcelSaveAs(xls, sTarget, iXlTextFormat);
string sName = "Sheet " + iSheet.ToString();
if (((string) sheet.Name).Length > 0) sName = sName + ": " + sheet.Name;
ExcelClose(xls);
string s = sName + "\r\n" + File.ReadAllText(sTarget).Trim();
sBook = sBook + (iSheet > 1 ? "\r\n" + "----------" + "\r\n" + sFormFeed + "\r\n" : "") + s;
xls = ExcelOpen(xlss, sSource);
sheets = xls.Sheets;
iSheet = iSheet + 1;
}
if (File.Exists(sTarget)) File.Delete(sTarget);
File.WriteAllText(sTarget, sBook);
ExcelClose(xls);
Marshal.ReleaseComObject(xls);
Marshal.ReleaseComObject(xlss);
if (bGet) { app.Visible = iVisible; app.DisplayAlerts = iAlerts; app.ScreenUpdating = iUpdating; }
else ExcelQuit(app);
Marshal.ReleaseComObject(app);
return File.Exists(sTarget);
} // Xls2Txt method

// ---- Web helpers ----
// DownloadFile: simple authenticated download. Modern WebClient replaces
// the former My.Computer.Network.DownloadFile.
public static void DownloadFile(string sUrl, string sFile, string sUserName, string sPassword) {
Homer.Web.configure();
using (WebClient web = new WebClient()) {
web.Headers[HttpRequestHeader.UserAgent] = Homer.Web.userAgent();
if (sUserName.Length > 0) web.Credentials = new NetworkCredential(sUserName, sPassword);
web.DownloadFile(sUrl, sFile);
}
} // DownloadFile method

// LookupTerm: legacy dictionary lookup via Internet Explorer automation.
// IE is removed from current Windows; retained for source compatibility.
public static string LookupTerm(string sWord) {
if (sWord.Length == 0) return "";
string sUrl = "http://dictionary.reference.com/browse/";
string sDivider = "\r\n" + "----------" + "\r\n" + sFormFeed + "\r\n";
string sText = sWord + "\r\n" + "\r\n" + "Contents" + "\r\n";
sText = sText + "dictionary.com" + "\r\n" + "thesaurus.com" + "\r\n" + "wikipedia.org";

dynamic ie = Activator.CreateInstance(Type.GetTypeFromProgID("InternetExplorer.Application"));
sUrl = Uri.EscapeUriString(sUrl + sWord);
ie.Navigate(sUrl);
while ((int) ie.ReadyState != 4) System.Threading.Thread.Sleep(100);
dynamic doc = ie.Document;
dynamic tables = doc.GetElementByTagName("table");
sText = sText + sDivider + "dictionary.com" + "\r\n" + "\r\n";
foreach (dynamic table in tables) sText = sText + table.InnerText + "\r\n" + "\r\n";
sText = sText.Trim() + "\r\n";
try { doc.Close(); ie.Quit(); Marshal.ReleaseComObject(ie); } catch { }
return sText;
} // LookupTerm method

// GetLinks: legacy link harvest via Internet Explorer automation.
public static List<string[]> GetLinks(string sUrl) {
List<string[]> listLinks = new List<string[]>();
List<string> listRefs = new List<string>();
dynamic ie = Activator.CreateInstance(Type.GetTypeFromProgID("InternetExplorer.Application"));
ie.Visible = false;
ie.Silent = true;
ie.Navigate(sUrl);
while ((int) ie.ReadyState != 4) System.Threading.Thread.Sleep(100);
dynamic doc = ie.Document;
dynamic links = doc.Links;
foreach (dynamic link in links) {
string sRef = link.HRef;
if (sRef.ToLower().StartsWith("mailto:")) continue;
try { Uri uri = new Uri(sRef); } catch { sRef = ""; }
if (sRef.Length == 0 || listRefs.Contains(sRef)) continue;
listRefs.Add(sRef);
listLinks.Add(new string[] {sRef, (string) link.InnerText});
}
try { doc.Close(); ie.Quit(); Marshal.ReleaseComObject(ie); } catch { }
return listLinks;
} // GetLinks method
} // VB class

public class Script {
private static MethodInfo miRun;

private static MethodInfo GetRunMethod() {
if (miRun != null) return miRun;
string sDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
string sDll = Path.Combine(sDir, "EdSharp.dll");
Assembly asmHost = Assembly.LoadFrom(sDll);
Type typeJs = asmHost.GetType("EdSharp.JS");
miRun = typeJs.GetMethod("runScript", new Type[] {typeof(string), typeof(object), typeof(object)});
return miRun;
} // GetRunMethod method

// run: evaluate sCode with the active editor window as frm and its
// RichTextBox as rtb, both visible to the snippet. Either may be null
// for host-internal expressions that need no document context.
public static string run(string sCode) {
object frm = (App.Frame != null) ? App.Frame.Child : null;
object rtb = (App.Frame != null && App.Frame.Child != null) ? (object) App.Frame.Child.RTB : null;
return (string) GetRunMethod().Invoke(null, new object[] {sCode, frm, rtb});
} // run method
} // Script class

public class COM {
public static object CreateObject(string sProgID) {
Type t = Type.GetTypeFromProgID(sProgID);
object oResult = Activator.CreateInstance(t);
return oResult;
} // CreateObject method

public static object GetObject(string sProgID) {
object oResult = Marshal.GetActiveObject(sProgID);
return oResult;
} // GetObject method

public static object GetOrCreateObject(string sProgID, out bool bCreate, string sMessage) {
object oResult;
try {
oResult = GetObject(sProgID);
bCreate = false;
}
catch {
Util.Say(sMessage);
oResult = CreateObject(sProgID);
bCreate = true;
}
return oResult;
} // GetOrCreateObject method

public static object CallMethod(object o, string sMethod) {
object[] args = {};
return CallMethod(o, sMethod, args);
} // CallMethod method

public static object CallMethod(object o, string sMethod, string sValue) {
object[] args = {sValue};
return CallMethod(o, sMethod, args);
} // CallMethod method

public static object CallMethod(object o, string sMethod, int iValue) {
object[] args = {iValue};
return CallMethod(o, sMethod, args);
} // CallMethod method

public static object CallMethod(object o, string sMethod, object[] args) {
Type t = o.GetType();
object oResult = t.InvokeMember(sMethod, BindingFlags.InvokeMethod, null, o, args);
return oResult;
} // CallMethod method

public static object SetProperty(object o, string sProperty, string sValue) {
object[] args = {sValue};
return SetProperty(o, sProperty, args);
} // SetProperty method

public static object SetProperty(object o, string sProperty, int iValue) {
object[] args = {iValue};
return SetProperty(o, sProperty, args);
} // SetProperty method

public static object SetProperty(object o, string sProperty, bool bValue) {
object[] args = {bValue};
return SetProperty(o, sProperty, args);
} // SetProperty method

public static object SetProperty(object o, string sProperty, object[] args) {
Type t = o.GetType();
object oResult = t.InvokeMember(sProperty, BindingFlags.SetProperty, null, o, args);
return oResult;
} // SetProperty method

public static object GetProperty(object o, string sProperty) {
object[] args = new object[] {};
return GetProperty(o, sProperty, args);
} // GetProperty method

public static object GetProperty(object o, string sProperty, object[] args) {
Type t = o.GetType();
object oResult = t.InvokeMember(sProperty, BindingFlags.GetProperty, null, o, args);
return oResult;
} // GetProperty method

public static bool JFWRunFunction(string sText) {
return JFWRunFunction(sText, ref App.JAWS);
} // JFWRunFunction method

public static bool JFWRunFunction(string sText, ref object oJFW) {
try {
if (oJFW == null) oJFW = CreateObject("FreedomSci.JawsApi");
bool bResult = (bool) CallMethod(oJFW, "RunFunction", new object[] {sText});
return bResult;
}
catch {
return false;
}
} // JFWRunFunction method

public static void InvokeVerb(string sPath, string sVerb) {
// Dialog.Show(sPath, sVerb);
object o = COM.CreateObject("Shell.Application");
string sDir = Path.GetDirectoryName(sPath);
string sName = Path.GetFileName(sPath);
o = COM.CallMethod(o, "Namespace", new string[] {sDir});
// o = COM.GetProperty(o, "Self");
o = COM.CallMethod(o, "ParseName", new string[] {sName});
o = COM.CallMethod(o, "InvokeVerb", new string[] {sVerb});
} // InvokeVerb method

public static string[] Verbs(string sPath) {
object o = COM.CreateObject("Shell.Application");
string sDir = Path.GetDirectoryName(sPath);
string sName = Path.GetFileName(sPath);
o = COM.CallMethod(o, "Namespace", new string[] {sDir});
o = COM.CallMethod(o, "ParseName", new string[] {sName});
try {
o = COM.CallMethod(o, "Verbs", new object[] {});
}
catch {
return new string[] {};
}
int iCount = (int) COM.GetProperty(o, "Count");
StringBuilder sb = new StringBuilder();
for (int i = 0; i < iCount; i++) {
object oVerb = COM.CallMethod(o, "Item", new object[] {(int) i});
string sVerb = (string) COM.GetProperty(oVerb, "Name");
if (sVerb.Trim() != "") sb.Append(sVerb + "\n");
}
string[] aVerbs = sb.ToString().Trim().Split('\n');
return aVerbs;
} // Verbs method

public static string ConvertFile2String(string sSource) {
int iConvert = 2;
return ConvertFile2String(sSource, ref iConvert);
} // ConvertFile2String method

public static string ConvertFile2String(string sSource, ref int iConvert) {
string sTargetExt = "txt";
return ConvertFile2String(sSource, ref iConvert, ref sTargetExt);
} // ConvertFile2String method

public static string ConvertFile2String(string sSource, ref int iConvert, ref string sTargetExt) {
bool bTextOnly = false;
return ConvertFile2String(sSource, ref iConvert, ref sTargetExt, bTextOnly);
} // Convert File2String method

public static string ConvertFile2String(string sSource, ref int iConvert, ref string sTargetExt, bool bTextOnly) {
string sText = "";
if (iConvert == 0) sText = Util.File2String(sSource);
else {
//string sTarget = App.TempFile;
string sTarget = Path.GetTempFileName();
App.TempFiles.Add(sTarget);
//sTarget = Win32.GetShortPath(sTarget);
string sResult;
string[] aResults = App.ReadSectionKeys("Import");
HomerList hl = new HomerList(aResults);
hl.AddUniqueRange("rtf|htm|html|xhtml");
hl.ToLower();
string sExt = Path.GetExtension(sSource).ToLower().TrimStart('.');
string sMatch = "^" + sExt + @"(2\w+)?$";
if (bTextOnly) sMatch = "^" + sExt + @"2txt$";
hl.KeepLike(sMatch);
// The bare extension entry and a real <ext>2txt converter both display
// as "txt", and choosing the bare one opens the file RAW -- a duplicated
// row whose twin quietly does something different (the Open Other Format
// confusion found in the 22 August 2026 audit). When the real converter
// exists it serves the txt row alone; when none does, the bare entry
// stays, and the raw-read fallback below is what its txt offer means.
if (hl.Contains(sExt) && hl.Contains(sExt + "2txt")) hl.Remove(sExt);
// The source format is never offered as a target: converting a
// document to its own format is Control+O's plain Open, and an "epub"
// row in an epub's target list is noise. (A leftover line here used to
// push exactly that entry, against its own comment.) And only TEXT
// formats belong in an import target list, since importing means
// bringing the document into the editor as text.
string[] aTextTargets = new string[] {"txt", "md", "mdx", "htm", "html", "xhtml", "rtf", "brf", "csv"};
HomerList hlText = new HomerList();
foreach (string sEntry in hl.ToArray()) {
string sTargetPart = Util.RegExpReplaceCase(sEntry, @"^\w+2", "");
bool bText = false;
foreach (string sAllowed in aTextTargets) if (sTargetPart == sAllowed) { bText = true; break; }
if (bText && sTargetPart != sExt) hlText.Add(sEntry);
}
hl = hlText;
aResults = hl.ToArray();
hl.ReplaceLike("^" + sExt + "$", sExt + "2txt");
hl.ReplaceLike(@"^\w+2", "");
string[] aDisplay = hl.ToArray();
Array.Sort(aDisplay, aResults);
// Solved above instead
// Solve TextConvert with brf
// if (bTextOnly) aResults = new string[] { sExt + "2txt"};
// if (bTextOnly) aResults = new string[] { sExt, sExt + "2txt"};
if (aResults.Length == 0) {
//Dialog.Show("Alert", "No import options for " + sExt);
//return "";
sResult = "";
}
else if (aResults.Length == 1) sResult = aResults[0];
else {
//sResult = Dialog.Pick("Import Format", aResults, true, 0);
string sTitle = "Import " + sExt + " to ";
// The list opens on the target format used LAST time, when this
// source can offer it -- open an epub as txt, and the next pdf's list
// starts on txt too. When the remembered target is absent, the first
// item is selected, as before.
string sLastTarget = App.ReadData("ImportTarget", "");
int iDefaultPick = 0;
for (int iPick = 0; iPick < aDisplay.Length; iPick++) {
if (aDisplay[iPick] == sLastTarget) { iDefaultPick = iPick; break; }
}
sResult = Dialog.Pick(sTitle, aResults, aDisplay, true, iDefaultPick);
//Dialog.Show(sResult);
if (sResult.Length == 0) return "";

}
string sTempExt = Util.RegExpReplaceCase(sResult, @"^\w+2", "");
//Dialog.Show(s, sResult);
if (sTempExt != sResult) sTargetExt = sTempExt;
else sResult = sExt;
// Remember the chosen target so the next import's list starts there.
if (sTargetExt != null && sTargetExt.Length > 0) App.WriteData("ImportTarget", sTargetExt);
string sCommand = Ini.ReadValue(App.IniFile, "Import", sResult, "");
if (sCommand.Length > 0) {
// Dialog.Show(sTargetExt, "target extension");
string s = Path.ChangeExtension(sTarget, sTargetExt);
if (!Util.Equiv(sTarget, s)) {
if (File.Exists(s)) File.Delete(s);
System.IO.File.Move(sTarget, s);
sTarget = s;
}

// Middleware policy for Markdown imports: the outside tool's only job
// is reading the source format -- 2htm turns Word documents, PDF
// (through Word's PDF Reflow), slides, and spreadsheets into HTML --
// and the HTML becomes Markdown INSIDE the binary with the same
// ReverseMarkdown converter the HTML paste command uses. Pandoc is not
// involved: it cannot read PDF anyway, and the fewer outside legs a
// conversion has, the fewer ways it can fail. The [Import] commands
// for a Markdown target therefore call any2htm.cmd, and this flag
// finishes the second leg here.
// Markdown is the one target no outside reader emits from Office and
// PDF sources, so it alone takes two legs: 2htm to HTML, then the
// binary's ReverseMarkdown. Plain text goes single-step through
// 2htm's own -p mode in any2txt.cmd, per the conversion-path review
// of 25 August 2026.
bool bFinishMarkdownInBinary = (sTargetExt == "md" && sCommand.IndexOf("any2htm.cmd", StringComparison.OrdinalIgnoreCase) >= 0);
if (bFinishMarkdownInBinary) sTarget = Path.ChangeExtension(sTarget, ".htm");
// The PDF path produces RICH Markdown for every target: headings, lists,
// tables and emphasis, never a plain text dump. When the target is HTML
// or plain text, the outside step still writes Markdown and the last
// step happens here -- Markdig for HTML, and HTML to text for plain --
// so one good conversion serves all three.
bool bFinishFromRichMarkdown = ((sTargetExt == "htm" || sTargetExt == "html" || sTargetExt == "txt") && sCommand.IndexOf("pdf2md.cmd", StringComparison.OrdinalIgnoreCase) >= 0);
if (bFinishFromRichMarkdown) sTarget = Path.ChangeExtension(sTarget, ".md");
sCommand = Util.ExpandCommandLine(sCommand, sSource, sTarget);
// Dialog.Show(sTarget, "target file");
// Dialog.Show(sCommand);
//Clipboard.SetText(sCommand);
//Clipboard.SetText(sTarget);
App.Frame.AddMessage("Converting");
if (File.Exists(sTarget)) sTarget = Win32.GetShortPath(sTarget);
if (File.Exists(sTarget)) File.Delete(sTarget);
Util.RunHideWait(sCommand);
if (!File.Exists(sTarget)) {
//Util.RunHide(sCommand);
sCommand = "cmd.exe /c " + sCommand;
//Util.RunHide(sCommand);
Util.RunHideWait(sCommand);
/*
int iLoop = 20;
while (iLoop > 0 && !File.Exists(sTarget)) {
System.Threading.Thread.Sleep(100);
iLoop--;
}
*/
}
if (bFinishFromRichMarkdown && File.Exists(sTarget)) {
string sRich = Util.ConvertedFile2String(sTarget);
string sFinished;
if (sTargetExt == "txt") sFinished = Util.Markdown2Text(sRich);
else sFinished = Util.Markdown2Html(sRich, Path.GetFileNameWithoutExtension(sSource));
try { File.Delete(sTarget); } catch (Exception) {}
sTarget = Path.ChangeExtension(sTarget, "." + sTargetExt);
Util.String2File(sFinished, sTarget);
Util.Log("finished rich PDF conversion as " + sTargetExt + " in the binary: " + sTarget);
}
if (bFinishMarkdownInBinary && File.Exists(sTarget)) {
// Second leg, in the binary: HTML to Markdown with ReverseMarkdown,
// and for a plain-text target, on to text -- the same conversions the
// editor's own commands use, with no second outside process. This
// covers pdf2txt and doc2txt too, which used to run a separate
// any2txt script with its own copy of the quoting defect.
string sHtml = Util.ConvertedFile2String(sTarget);
string sConverted = Util.Html2Markdown(sHtml);
try { File.Delete(sTarget); } catch (Exception) {}
sTarget = Path.ChangeExtension(sTarget, "." + sTargetExt);
Util.String2File(sConverted, sTarget);
Util.Log("converted to " + sTargetExt + " in the binary: " + sTarget);
}
// Read the converted target. File2String detects its encoding (byte-order
// mark first, then content detection) and decodes it correctly, so the old
// re-encode pass through Convert\EasyEncode\utf8b.exe is no longer needed.
// Dropping it removes that external tool from the conversion path.
if (File.Exists(sTarget)) sText = Util.ConvertedFile2String(sTarget);
// Last line of defense against the far-eastern gibberish: whatever
// decided the encoding, text that is mostly CJK characters when the
// source document was not is the signature of single-byte text read as
// UTF-16. Re-read it the plain way and log the rescue.
// A genuinely Chinese or Japanese document is also mostly CJK, so the
// rescue is accepted only when reading the file plainly makes the CJK
// share go away -- proof that the characters were an artifact of the
// decode rather than the document's own script.
if (sText.Length > 0 && Util.LooksLikeMisreadUtf16(sText)) {
string sPlain = Util.PlainFile2String(sTarget);
if (sPlain.Length > 0 && !Util.LooksLikeMisreadUtf16(sPlain)) {
sText = sPlain;
Util.Log("re-read as single-byte text after a wide-encoding misread: " + sTarget);
}
}

if (sText.Length == 0) {
// The conversion scripts capture their tool's console output beside
// the target as <target>.log; its last lines usually name the real
// problem, so show them right in the dialog.
string sToolLog = sTarget + ".log";
string sToolWords = "";
try {
if (File.Exists(sToolLog)) {
string[] aToolLines = File.ReadAllLines(sToolLog);
int iFrom = Math.Max(0, aToolLines.Length - 6);
sToolWords = "\n\nThe converter said:\n" + String.Join("\n", aToolLines, iFrom, aToolLines.Length - iFrom);
}
}
catch (Exception) {}
Dialog.Show("Error", "The conversion produced no output.\nCommand line:\n" + sCommand + sToolWords + "\n\nThe run log records each step and its exit code:\n" + App.LogFile);
}
}
else {
if (sTargetExt == sExt) {
sExt = "";
iConvert = 1;
}

switch (sExt) {
case "rtf" :
if (iConvert > 0) iConvert = -1;
break;
//Use OfficeConvert utilities
/*
case "doc" :
case "docx" :
App.Frame.AddMessage("Converting");
sText = WordFile2String(sSource);
break;
case "ppt" :
case "pptx" :
App.Frame.AddMessage("Converting");
VB.Ppt2Txt(sSource, sTarget);
sText = Util.File2String(sTarget);
break;
case "xls" :
case "xlsx" :
App.Frame.AddMessage("Converting");
VB.Xls2Txt(sSource, sTarget);
sText = Util.File2String(sTarget);
break;
*/
default :
// Disable Word conversions of unknown extensions
// if (iConvert == 1) {
if (iConvert != -1) {
sText = Util.File2String(sSource);
iConvert = 0;
}
else {
App.Frame.AddMessage("Converting");
sText = WordFile2String(sSource);
}
break;
}
}
}
App.Frame.Activate();
sText = Util.Convert2UnixLineBreak(sText);
return sText;
} // ConvertFile2String method

public static string WordFile2String(string sSource) {
bool bCreate, bVisible;
int iDisplayAlerts;
bool bAppVisible = false;
//object oApp = COM.GetOrCreateObject("Word.Application", out bCreate);
object oApp = COM.WordAccess(out bCreate);
bVisible = (bool) COM.GetProperty(oApp, "Visible");
iDisplayAlerts = (int) COM.GetProperty(oApp, "DisplayAlerts");
COM.SetProperty(oApp, "Visible", bAppVisible);
COM.SetProperty(oApp, "DisplayAlerts", 0);
object oDocs =COM.GetProperty(oApp, "Documents");
object oDoc = VB.WordOpen(oDocs, sSource, bAppVisible);
string sTarget = Path.GetTempFileName();
if (File.Exists(sTarget)) File.Delete(sTarget);
object oSelection = COM.GetProperty(oApp, "Selection");
int iLength = (int) COM.GetProperty(oSelection, "StoryLength");
COM.CallMethod(oSelection, "SetRange", new object[] {0, iLength});
string sText = (string) COM.GetProperty(oSelection, "Text");
COM.Release(ref oSelection);
sText = sText.Trim();
sText = Util.RegExpReplaceCase(sText, "\r\f", "\f\r");
sText = Util.Convert2UnixLineBreak(sText);
sText = Util.RegExpReplaceCase(sText, MdiFrame.SB, MdiFrame.SectionBreak);
//sText = Util.Convert2WinLineBreak(sText);
Util.String2File(sText, sTarget);
//VB.WordSaveAs(oDoc, sTarget, 2);
VB.WordClose(oDoc);
COM.Release(ref oDoc);
COM.Release(ref oDocs);

if (bCreate) {
//VB.WordQuit(oApp);
}
else {
COM.SetProperty(oApp, "Visible", bVisible);
COM.SetProperty(oApp, "DisplayAlerts", iDisplayAlerts);
}

COM.Release(ref oApp);
if (File.Exists(sTarget)) {
string sReturn = Util.File2String(sTarget);
File.Delete(sTarget);
return sReturn;
}
else return "";
} // WordFile2String();

public static object WordOpen(object oDocs, string sFile, bool bAppVisible) {
bool bConfirmConversions = false;
bool bReadOnly = false;
bool bAddToRecentFiles = false;
object sPasswordDocument = Missing.Value;
object sPasswordTemplate = Missing.Value;
bool bRevert = true;
object sWritePasswordDocument = Missing.Value;
object sWritePasswordTemplate = Missing.Value;
object iFormat = Missing.Value;
object iEncoding = Missing.Value;
bool bVisible = bAppVisible;
object oOpenConflictDocument = Missing.Value;
bool bOpenAndRepair = true;
object iDocumentDirection = Missing.Value;
bool bNoEncodingDialog = true;

object[] oParams = {sFile, bConfirmConversions, bReadOnly, bAddToRecentFiles, sPasswordDocument, sPasswordTemplate, bRevert, sWritePasswordDocument, sWritePasswordTemplate, iFormat, iEncoding, bVisible, oOpenConflictDocument, bOpenAndRepair, iDocumentDirection, bNoEncodingDialog};
oParams = new object[] {sFile, bConfirmConversions, bReadOnly, bAddToRecentFiles};

object oDoc = null;
try {
oDoc = CallMethod(oDocs, "Open", oParams);
}
catch (COMException ex) {
Dialog.Show("Error", ex.Message);
}
return oDoc;
} // OpenWordDocument method

public static void WordSaveAs(object oDoc, string sFile, int iSaveFormat) {
int iFileFormat = iSaveFormat;
bool bLockComments = false;
object oPassword = Type.Missing;
bool bAddToRecentFiles = false;
object oWritePassword = Type.Missing;
bool bReadOnlyRecommended = false;
bool bEmbedTrueTypeFonts = false;
bool bSaveNativePictureFormat = false;
bool bSaveFormsData = false;
bool bSaveAsAOCELetter = false;
object oEncoding= Type.Missing;
bool bInsertLineBreaks = false;
bool bAllowSubstitutions = false;
object sLineEnding = Type.Missing;
bool bAddBiDiMarks = false;
object[] oParams = {sFile, iFileFormat, bLockComments, oPassword, bAddToRecentFiles, oWritePassword, bReadOnlyRecommended, bEmbedTrueTypeFonts, bSaveNativePictureFormat,
bSaveFormsData, bSaveAsAOCELetter, oEncoding, bInsertLineBreaks, bAllowSubstitutions, sLineEnding, bAddBiDiMarks
};
try {
CallMethod(oDoc, "SaveAs", oParams);
}
catch (COMException ex) {
Dialog.Show("Error", ex.Message);
}
} // WordSaveAs method

public static void WordClose(object oDoc) {
object oApp = GetProperty(oDoc, "Application");
ClearNormalTemplate(oApp);

int iSaveChanges = 0;
object iOriginalFormat = Type.Missing;
bool bRouteDocument = false;
object[] oParams = {iSaveChanges, iOriginalFormat, bRouteDocument};
oParams = new object[] {iSaveChanges};

try {
CallMethod(oDoc, "Close", oParams);
}
catch (COMException ex) {
Dialog.Show("Error", ex.Message);
}
} // WordClose method

public static void ClearNormalTemplate(object oApp) {
object oTemplate = COM.GetProperty(oApp, "NormalTemplate");
COM.SetProperty(oTemplate, "Saved", true);
Release(ref oTemplate);
} // ClearNormalTemplate method

public static void WordQuit(object oApp) {
COM.ClearNormalTemplate(oApp);

int iSaveChanges = 0;
object iFormat = Type.Missing;
bool bRouteDocument = false;

object[] oParams = {iSaveChanges, iFormat, bRouteDocument};
oParams = new object[] {iSaveChanges};

try {
CallMethod(oApp, "Quit", oParams);
}
catch (COMException ex) {
Dialog.Show("Error", ex.Message);
}
} // WordQuit method

public static void Release(ref object o) {
Marshal.ReleaseComObject(o);
o = null;
} // Release method

public static object WordAccess(out bool bCreate) {
string sMessage = "Initializing Microsoft Word";
object oApp = GetOrCreateObject("Word.Application", out bCreate, sMessage);
if (bCreate) App.WordCreated = true;
return oApp;
} // WordAccess method

public static void WordExit() {
object oApp = null;
bool bLoop = true;
while (bLoop) {
try {
oApp = GetObject("Word.Application");
//Util.Say("quit");
//WordQuit(oApp);
//break;
VB.WordQuit(oApp);
Release(ref oApp);
}
catch {
break;
}
}
Util.TerminateProcess("WinWord");
} // WordExit method;

public static bool WordSource2TargetFormat(string sSource, string sTarget, string sFormat) {
int iFormat = 2; // text;
if (sFormat == "doc") iFormat = 0;
else if (sFormat == "htm") iFormat = 10;
else if (sFormat == "xml") iFormat = 11;
bool bAppVisible = false;
bool bCreate;
object oApp = COM.WordAccess(out bCreate);
bool bVisible = (bool) COM.GetProperty(oApp, "Visible");
int iDisplayAlerts = (int) COM.GetProperty(oApp, "DisplayAlerts");
COM.SetProperty(oApp, "Visible", bAppVisible);
COM.SetProperty(oApp, "DisplayAlerts", 0);
object oDocs = COM.GetProperty(oApp, "Documents");
object oDoc = VB.WordOpen(oDocs, sSource, bAppVisible);
object oSelection = COM.GetProperty(oApp, "Selection");
//object oRange = COM.GetProperty(oSelection, "Range");
int iLength = (int) COM.GetProperty(oSelection, "StoryLength");
object oRange = COM.CallMethod(oDoc, "Range", new object[] {0, iLength});
COM.CallMethod(oRange, "AutoFormat");
VB.WordSaveAs(oDoc, sTarget, iFormat);
COM.Release(ref oRange);
VB.WordClose(oDoc);
COM.Release(ref oDoc);
COM.Release(ref oDocs);
if (!bCreate) {
COM.SetProperty(oApp, "Visible", bVisible);
COM.SetProperty(oApp, "DisplayAlerts", iDisplayAlerts);
}
COM.Release(ref oApp);
App.Frame.Activate();
return File.Exists(sTarget);
} // WordSource2TargetFormat method

public static string GetUrl() {
string sUrl = "";
try {
object oShell = COM.CreateObject("Shell.Application");
object oWindows = COM.CallMethod(oShell, "Windows");
int iCount = (int) COM.GetProperty(oWindows, "Count");
if (iCount > 0) {
object oWindow = COM.CallMethod(oWindows, "Item", new object[] {iCount - 1});
sUrl = (string) COM.GetProperty(oWindow, "LocationURL");
}
}
catch {}
return sUrl;
} // GetUrl method

public static void ActivateTitle(string sTitle) {
object oShell = CreateObject("WScript.Shell");
CallMethod(oShell, "AppActivate", sTitle);
} // ActivateTitle method

} // COM class

// Inix: app-level accessor for an optional EdSharp.inix layered over the
// classic .ini. Reads are inix-first with .ini fallback (see Ini.ReadValue):
// when no EdSharp.inix exists, or a key is absent from it, lookups miss and the
// classic .ini path runs unchanged, so default behavior is preserved. Built on
// the portable Homer.InixCodec.
public static class Inix {
static List<Homer.InixCodec.Section> lSections = null;
static List<Homer.InixCodec.Section> lProgram = null;
static bool bLoaded = false;

static string getPath() {
string sIni = App.IniFile;
if (String.IsNullOrEmpty(sIni)) return null;
if (sIni.ToLower().EndsWith(".ini")) return sIni.Substring(0, sIni.Length - 4) + ".inix";
return sIni + ".inix";
} // getPath method

static string getProgramPath() {
// The program-folder EdSharp.inix is the shipped configuration layer:
// the installer refreshes it on every upgrade (ignoreversion), unlike
// the data-folder files, which belong to the user and are preserved.
if (String.IsNullOrEmpty(App.ProgramDir)) return null;
return Path.Combine(App.ProgramDir, App.GetAppName() + ".inix");
} // getProgramPath method

static void ensureLoaded() {
if (bLoaded) return;
bLoaded = true;
try {
string sPath = getPath();
if (sPath != null && File.Exists(sPath)) lSections = Homer.InixCodec.read(sPath);
}
catch { lSections = null; }
try {
string sProgramPath = getProgramPath();
if (sProgramPath != null && File.Exists(sProgramPath)) lProgram = Homer.InixCodec.read(sProgramPath);
}
catch { lProgram = null; }
} // ensureLoaded method

public static bool tryGet(string sSection, string sKey, out string sValue) {
// Layer order: the data-folder EdSharp.inix (the user's override
// layer) wins over the program-folder EdSharp.inix (the shipped
// configuration layer); both override the classic .ini.
sValue = null;
ensureLoaded();
List<List<Homer.InixCodec.Section>> lLayers = new List<List<Homer.InixCodec.Section>>();
if (lSections != null) lLayers.Add(lSections);
if (lProgram != null) lLayers.Add(lProgram);
foreach (List<Homer.InixCodec.Section> lLayer in lLayers) {
foreach (Homer.InixCodec.Section sec in lLayer) {
if (!Util.Equiv(sec.Name, sSection)) continue;
foreach (Homer.InixCodec.Pair pair in sec.Pairs) if (Util.Equiv(pair.Key, sKey)) { sValue = pair.Value; return true; }
}
}
return false;
} // tryGet method

// keyNames: every key name a section defines across both inix layers.
// Lets Ini.ReadSectionKeys list keys that exist only in an inix (new
// conversions, for example) and recognize tombstones: keys an inix
// defines with an empty value to retire a stale entry left in an old
// user .ini, without editing that file.
public static List<string> keyNames(string sSection) {
List<string> lReturn = new List<string>();
ensureLoaded();
List<List<Homer.InixCodec.Section>> lLayers = new List<List<Homer.InixCodec.Section>>();
if (lSections != null) lLayers.Add(lSections);
if (lProgram != null) lLayers.Add(lProgram);
foreach (List<Homer.InixCodec.Section> lLayer in lLayers) {
foreach (Homer.InixCodec.Section sec in lLayer) {
if (!Util.Equiv(sec.Name, sSection)) continue;
foreach (Homer.InixCodec.Pair pair in sec.Pairs) if (!String.IsNullOrEmpty(pair.Key)) lReturn.Add(pair.Key);
}
}
return lReturn;
} // keyNames method

// syncWrite: when an EdSharp.inix exists (the user opted into the layer),
// keep it consistent with writes to the main .ini so an override never masks
// a value the app just changed. A no-op when no EdSharp.inix is present, so
// default behavior is unchanged. Guarded and caught; never disturbs the .ini
// write that already happened.
public static void syncWrite(string sFile, string sSection, string sKey, string sValue) {
if (String.IsNullOrEmpty(sFile) || !Util.Equiv(sFile, App.IniFile)) return;
string sPath = getPath();
if (sPath == null || !File.Exists(sPath)) return;
try { Homer.InixCodec.writeValue(sPath, sSection, sKey, sValue); reload(); }
catch {}
} // syncWrite method

public static void reload() { bLoaded = false; lSections = null; lProgram = null; } // reload method
} // Inix class

public class Ini {
public static string RedirectFile(string sFile, string sSection) {
if(Util.Equiv(sFile, App.IniFile) && (Util.Equiv(sSection, "Favorites") || Util.Equiv(sSection, "Recent") || Util.Equiv(sFile, "Tokens"))) sFile = Path.Combine(App.DataDir, App.ReadData("Compiler", "Default") + ".ini");
return sFile;
} // RedirectFile method

[DllImport("kernel32.dll")]
public static extern int GetPrivateProfileString(string sSection, string sKey, string sDefault, StringBuilder sReturnString, int iLength, string sFile);
public static String ReadValue(String sFile, String sSection, String sKey, string sDefault) {
string sInix;
if (Util.Equiv(sFile, App.IniFile) && Inix.tryGet(sSection, sKey, out sInix)) return sInix;
sFile = RedirectFile(sFile, sSection);
StringBuilder sb = new StringBuilder(260);
if (GetPrivateProfileString(sSection, sKey, sDefault, sb, sb.Capacity, sFile) > 0) return sb.ToString();
else return sDefault;
} // ReadValue method

[DllImport("kernel32.dll")]
public static extern bool WritePrivateProfileString(string sSection, string sKey, string sValue, string sFile);
public static bool WriteQuote(String sFile, String sSection, String sKey, String sValue) {
bool bQuote = true;
return WriteValue(sFile, sSection, sKey, sValue, bQuote);
} // WriteQuote method

public static bool WriteValue(String sFile, String sSection, String sKey, String sValue) {
bool bQuote = false;
return WriteValue(sFile, sSection, sKey, sValue, bQuote);
} // WriteValue method

public static bool WriteValue(String sFile, String sSection, String sKey, String sValue, bool bQuote) {
sFile = RedirectFile(sFile, sSection);
string sRaw = sValue;
if (bQuote) sValue = "\"" + sValue + "\"";
bool bReturn = WritePrivateProfileString(sSection, sKey, sValue, sFile);
FlushFile(sFile);
Inix.syncWrite(sFile, sSection, sKey, sRaw);
return bReturn;
} // WriteValue method

[DllImport("kernel32.dll")]
public static extern bool WritePrivateProfileString(string sSection, string sKey, int iValue, string sFile);
public static bool DeleteKey(String sFile, String sSection, String sKey) {
sFile = RedirectFile(sFile, sSection);
int iValue = 0;
bool bReturn = WritePrivateProfileString(sSection, sKey, iValue, sFile);
FlushFile(sFile);
return bReturn;
} // DeleteKey method

[DllImport("kernel32.dll")]
public static extern bool WritePrivateProfileString(string sSection, int iKey, int iValue, string sFile);
public static bool DeleteSection(String sFile, String sSection) {
int iKey = 0;
int iValue = 0;
bool bReturn = WritePrivateProfileString(sSection, iKey, iValue, sFile);
FlushFile(sFile);
return bReturn;
} // DeleteSection method

[DllImport("kernel32.dll")]
public static extern bool WritePrivateProfileString(int iSection, int iKey, int iValue, string sFile);
public static bool FlushFile(String sFile) {
int iSection = 0;
int iKey = 0;
int iValue = 0;
return WritePrivateProfileString(iSection, iKey, iValue, sFile);
} // FlushFile method

public static string[] ReadSectionKeys(string sFile, string sSection) {
bool bIncludeComments = false;
return ReadSectionKeys(sFile, sSection, bIncludeComments);
} // ReadSectionKeys method

public static string[] ReadSectionKeys(string sFile, string sSection, bool bIncludeComments) {
string sFileOriginal = sFile;
sFile = RedirectFile(sFile, sSection);
string[] aDefault = new string[] {};
if (!File.Exists(sFile)) return mergeInixKeys(sFileOriginal, sSection, bIncludeComments, aDefault);

string sText = Util.File2String(sFile);
string sMatch = "^\\[" + sSection + "\\](.|\n)*?((\n\\[)|\\Z)";
object[] aResult = Util.RegExpContainsCase(sText, sMatch);
int iIndex = (int) aResult[0];
if (iIndex == -1) return mergeInixKeys(sFileOriginal, sSection, bIncludeComments, aDefault);

string sValue = (string) aResult[1];
string[] aLines = sValue.Split('\n');
StringBuilder sb = new StringBuilder();
foreach (string sLine in aLines) {
string s = sLine.TrimStart();
if (s.Length == 0 || (!bIncludeComments && s.StartsWith(";")) || s.StartsWith("=") || !s.Contains("=")) continue;
int i = s.IndexOf("=");
string sKey = s.Substring(0, i).TrimEnd();
sb.Append(sKey + "\n");
}

string sKeys = sb.ToString().Trim();
string[] aReturn = (sKeys.Length == 0) ? aDefault : sKeys.Split('\n');
return mergeInixKeys(sFileOriginal, sSection, bIncludeComments, aReturn);
} // ReadSectionKeys method

// mergeInixKeys: for the main ini only, append key names that exist
// only in the inix layers, then drop keys whose effective layered
// value is empty. An empty inix value is a tombstone: it disables and
// hides a stale entry that an old user .ini still carries, without
// editing the user's file. Sections read with comments included are
// passed through untouched.
static string[] mergeInixKeys(string sFile, string sSection, bool bIncludeComments, string[] aKeys) {
if (bIncludeComments || !Util.Equiv(sFile, App.IniFile)) return aKeys;
List<string> lMerged = new List<string>(aKeys);
foreach (string sOne in Inix.keyNames(sSection)) {
bool bFound = false;
foreach (string sHave in lMerged) if (Util.Equiv(sHave, sOne)) { bFound = true; break; }
if (!bFound) lMerged.Add(sOne);
}
List<string> lReturn = new List<string>();
foreach (string sOne in lMerged) {
string sInix;
if (Inix.tryGet(sSection, sOne, out sInix) && sInix != null && sInix.Length == 0) continue;
lReturn.Add(sOne);
}
return lReturn.ToArray();
} // mergeInixKeys method

public static string[] ReadSections(string sFile) {
string[] aDefault = new string[] {};
if (!File.Exists(sFile)) return aDefault;

string sText = Util.File2String(sFile);
string sMatch = "^\\[.+?\\]\r\n";
string[] aResults = Util.RegExpExtractCase(sText, sMatch);
string sSections = String.Join("", aResults).Trim();
if (sSections.Length == 0) return aDefault;

sSections = sSections.Replace("[", "").Replace("]", "").Replace("\r", "");
string[] aReturn = sSections.Split('\n');
return aReturn;
} // ReadSections method

} // Ini class

public class Win32 {
[DllImport("user32.dll")]
public static extern int AttachThreadInput(int iThread1, int iThread2, int iAttach);

[DllImport("user32.dll")]
public static extern IntPtr GetActiveWindow();

[DllImport("user32.dll")]
public static extern int BringWindowToTop(IntPtr h);

[DllImport("user32.dll")]
public static extern int ShowWindow(IntPtr h, int iState);

[DllImport("kernel32.dll")]
public static extern int GetCurrentThreadId();

[DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(IntPtr h, int iProcess);

public static bool ForceWindow(IntPtr h) {
int iForegroundThread = GetWindowThreadProcessId(GetForegroundWindow(), 0);
int iAppThread = GetCurrentThreadId();
if (iForegroundThread == iAppThread) {
BringWindowToTop(h);
ShowWindow(h,3);
}
else {
AttachThreadInput(iForegroundThread, iAppThread, 1);
BringWindowToTop(h);
ShowWindow(h,3);
AttachThreadInput(iForegroundThread, iAppThread, 0);
}

return GetActiveWindow() == h;
} // ForceWindow method

[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
public static extern int GetShortPathName(string path, StringBuilder shortPath, int shortPathLength);
public static string GetShortPath(string sLongPath) {
StringBuilder sbShortPath = new StringBuilder(260);
GetShortPathName(sLongPath, sbShortPath, sbShortPath.Capacity);
string sReturn = sbShortPath.ToString().Trim();
if (sReturn.Length == 0) sReturn = Path.Combine(GetShortPath(Path.GetDirectoryName(sLongPath)), Path.GetFileName(sLongPath));
return sReturn;
} // GetShortPath method

[DllImport("user32.dll")]
static extern bool SystemParametersInfo(int iAction, int iParam, out bool bActive, int iUpdate);
public static bool IsScreenReaderActive() {
int iAction = 70; // SPI_GETSCREENREADER constant;
int iParam = 0;
bool bActive;
int iUpdate = 0;
bool bReturn = SystemParametersInfo(iAction, iParam, out bActive, iUpdate);
return bReturn && bActive;
} // IsScreenReaderActive method

[DllImport("user32.dll")]
public static extern int SendMessage(IntPtr h, int iMsg, int wParam, int lParam);

// Pointer-sized overload.  wParam and lParam are pointer-sized, so on a 64-bit
// build (x64 or ARM64) they must NOT be passed as int: a window handle above
// 2 GB truncates, and the message then goes to the wrong window or none at all.
// Used for the clipboard-viewer chain, whose parameters are window handles.
[DllImport("user32.dll", EntryPoint = "SendMessage")]
public static extern IntPtr SendMessagePtr(IntPtr h, int iMsg, IntPtr wParam, IntPtr lParam);

[DllImport("user32.dll")]
public static extern IntPtr GetForegroundWindow();

[DllImport("user32.dll")]
public static extern int SetForegroundWindow(IntPtr h);

[DllImport("user32.dll")]
public static extern IntPtr FindWindow(string sClass, string sTitle);

[DllImport("user32.dll")]
public static extern IntPtr FindWindow(int iClass, string sTitle);

[DllImport("user32.dll")]
public static extern IntPtr FindWindow(string sClass, int iTitle);

[DllImport("shell32.dll")]
public static extern int ShellExecute(int i1, string sVerb, string sFile, int i2, int i3, int i4);

public static int ShellExecute(string sVerb, string sFile) {
return ShellExecute(0, sVerb, sFile, 0, 0, 1);
} // ShellExecute method

[DllImport("shell32.dll")]
public static extern int ShellExecute(int i1, int i2, string sFile, int i3, int i4, int i5);

public static int ShellDefault(string sFile) {
return ShellExecute(0, 0, sFile, 0, 0, 1);
} // ShellDefault method

[DllImport("MSCorEE.dll", CharSet = CharSet.Auto)]
public static extern int GetCORSystemDirectory  (StringBuilder sbPath, int iSize, out int iLength);
public static string GetNetSdkDir() {
int iSize = 260;
StringBuilder sbPath = new StringBuilder(iSize);
int iLength;
GetCORSystemDirectory  (sbPath, iSize, out iLength);
return sbPath.ToString();
} // GetNetSdkDir method

[DllImport("MSCorEE.dll", CharSet = CharSet.Auto)]
// public static extern int GetRuntimeDirectory  (StringBuilder sbPath, int iSize, out int iLength);
public static extern int GetRuntimeDirectory  (StringBuilder sbPath, out int iLength);
public static string GetNetRuntimeDir() {
int iSize = 260;
StringBuilder sbPath = new StringBuilder(iSize);
// int iLength;
int iLength = 260;
// GetRuntimeDirectory  (sbPath, iSize, out iLength);
GetRuntimeDirectory  (sbPath, out iLength);
return sbPath.ToString();
} // GetNetRuntimeDir method


public static string GetJFWDir() {
RegistryKey key = Registry.LocalMachine;
string sSubKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\";
string sName = "Path";
string sPath = GetRegString(key, (sSubKey + "jfw.exe"), sName);

if (sPath == "") {
string[] sVersions = {"12", "11", "10", "90", "81", "80", "8", "71", "70", "7", "62", "61", "60", "6"};
sName = "";
foreach (string sVersion in sVersions) {
sPath = GetRegString(key, (sSubKey + "jaws" + sVersion + ".exe"), sName);
if (sPath != "") {
sPath = Path.GetDirectoryName(sPath);
break;
}
}
}
//if (sPath !="" && !sPath.EndsWith(@"\")) sPath = String.Concat(sPath, @"\");
sPath = sPath.TrimEnd('\\');
return sPath;
} // GetJFWDir method

public static string GetWEDir() {
RegistryKey key = Registry.LocalMachine;
string sSubKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinEyes.exe";
string sName = "Path";
string sPath = GetRegString(key, sSubKey, sName);
if (sPath !="" && !sPath.EndsWith(@"\")) sPath = String.Concat(sPath, @"\");
return sPath;
} // GetWEDir method

public static string GetRegString(RegistryKey key, string sSubKey, string sName) {
RegistryKey subkey = null;
string sData = "";

try {
subkey = key.OpenSubKey(sSubKey);
sData = subkey.GetValue(sName).ToString();
}
catch {
}
finally {
if (subkey != null) subkey.Close();
}
return sData;
} // GetRegString method

[DllImport("urlmon.dll")]
public static extern int URLDownloadToFile(int i1, string sUrl, string sFile, int i2, int i3, int i4);

public static bool Url2File(string sUrl, string sFile) {
int iResult = URLDownloadToFile(0, sUrl, sFile, 0, 0, 0);
return iResult == 0;
} // Url2File method

[Serializable]
public struct ShellExecuteInfo {
public int Size;
public uint Mask;
public IntPtr hwnd;
public string Verb;
public string File;
public string Parameters;
public string Directory;
public uint Show;
public IntPtr InstApp;
public IntPtr IDList;
public string Class;
public IntPtr hkeyClass;
public uint HotKey;
public IntPtr Icon;
public IntPtr Monitor;
}

[DllImport("shell32.dll", SetLastError = true)]
extern public static bool ShellExecuteEx(ref ShellExecuteInfo lpExecInfo);

public const uint SW_NORMAL = 1;

public static void OpenWith(string file) {
ShellExecuteInfo sei = new ShellExecuteInfo();
sei.Size = Marshal.SizeOf(sei);
sei.Verb = "openas";
sei.File = file;
sei.Show = SW_NORMAL;
if (!ShellExecuteEx(ref sei))
throw new System.ComponentModel.Win32Exception();
} //OpenAs method

} // Win32 class

public class Util {

public static string GetPortableExecutableKind() {
PortableExecutableKinds peKind  ;
ImageFileMachine machine  ;

// Module module = App.Shell.GetType().Module;
Module module = Assembly.GetExecutingAssembly().ManifestModule;
module.GetPEKind(out peKind, out machine);

if ((peKind & PortableExecutableKinds.ILOnly) != 0) // Assembly is platform independent.
{}
else { // assembly is platform dependent
switch (machine) {
case ImageFileMachine.I386: // i386, x86, IA-32, ... dependent.
break;
case ImageFileMachine.IA64: // IA-64 dependent.
break;
case ImageFileMachine.AMD64: // AMD-64, x64 dependent.
break;
} // switch
} // if

Dictionary<string, string> dFlags = new Dictionary<string, string>();
dFlags.Add("NotAPortableExecutableImage", "The file is not in portable executable (PE) file format.");
dFlags.Add("ILOnly", "The executable contains only Microsoft intermediate language (MSIL).");
dFlags.Add("Required32Bit", "The executable can be run on a 32-bit platform, or in the 32-bit Windows on Windows (WOW) environment on a 64-bit platform.");
dFlags.Add("PE32Plus", "The executable requires a 64-bit platform.");
dFlags.Add("Unmanaged32Bit", "The executable contains pure unmanaged code.");
dFlags.Add("I386", "Targets a 32-bit Intel processor.");
dFlags.Add("IA64", "Targets a 64-bit Intel processor.");
dFlags.Add("AMD64", "Targets a 64-bit AMD processor.");
string sReturn = "";
string sPEKind = peKind.ToString();
string[] aPEKind = sPEKind.Split(',');
string sMachine = machine.ToString();
foreach (string s in aPEKind) {
sPEKind = s.Trim();
if (dFlags.ContainsKey(sPEKind)) sReturn += dFlags[sPEKind] + "\n\n";
else sReturn += sPEKind + "\n\n";
} // foreach

// Not useful info
// if (dFlags.ContainsKey(sMachine)) sReturn += dFlags[sMachine] + "\n\n";
// else sReturn += sMachine + "\n\n";
sReturn += "Running in " + (IntPtr.Size == 8 ? "64" : "32") + "-bit mode.";
// sReturn = sReturn.Replace("\nTargets a ", "\nRunning on a ");
// Dialog.Show("Portable Executable Kind", sReturn);
return sReturn;
} // GetPortableExecutableKind method

public static string GetBomStringFromBytes(byte[] aBom) {
string sReturn = "";
foreach (byte b in aBom) {
if (sReturn.Length > 0) sReturn += "|";
sReturn += b;
}
return sReturn;
} // GetBomStringFromBytes method

public static string GetBomStringFromFile(string sFile) {
FileStream file = new FileStream(sFile, FileMode.Open, FileAccess.Read, FileShare.Read);
byte[] aBom = new byte[4];
int iCount = file.Read(aBom, 0, 4);
file.Close();
byte[] aReturn = new byte[iCount];
for (int i = 0; i < iCount; i++) aReturn[i] = aBom[i];
return GetBomStringFromBytes(aReturn);
} // GetBom method

public static Dictionary<string, int> GetBomDictionary() {
Dictionary<string, int> dCodes = new Dictionary<string, int>();
Dictionary<string, int> dBoms = new Dictionary<string, int>();
dCodes.Add("Unicode (Big-Endian)", 1201);
dCodes.Add("Unicode (UTF-32 Big-Endian)", 12001);
dCodes.Add("Unicode (UTF-32)", 12000);
// dCodes.Add("Unicode (UTF-7)", 65000);
dCodes.Add("Unicode (UTF-8)", 65001);
dCodes.Add("Unicode", 1200);

string sBody = "";
foreach (string sKey in dCodes.Keys) {
int iValue = dCodes[sKey];
Encoding en = Encoding.GetEncoding(iValue);
string sFile = Path.GetTempFileName();
File.WriteAllText(sFile, sBody, en);

string sBom = GetBomStringFromFile(sFile);
// MessageBox.Show(en.EncodingName, sBom);
// if (dBoms.ContainsKey(sBom)) MessageBox.Show(en.EncodingName, Encoding.GetEncoding(dBoms[sBom]).EncodingName);
dBoms.Add(sBom, iValue);
File.Delete(sFile);
}
return dBoms;
} // GetBomDictionary method

public static Encoding GetFileEncoding(string sFile) {
Dictionary<string, int> dBom = GetBomDictionary();
return GetFileEncoding(sFile, dBom);
} // GetFileEncoding method

public static Encoding GetFileEncoding(string sFile, Dictionary<string, int> dBom) {
// A byte-order mark is definitive, so it wins. Without one, fall back to
// content detection (DetectEncodingNoBom), which prefers UTF-8 with BOM.
string sBom = GetBomStringFromFile(sFile);
Encoding en = null;
foreach (string s in dBom.Keys) {
if (sBom.StartsWith(s)) {
en = Encoding.GetEncoding(dBom[s]);
break;
}
}
if (en == null) en = DetectEncodingNoBom(sFile);
return en;
} // GetFileEncoding method

public static Encoding DetectEncodingNoBom(string sFile) {
// Content-based detection for a file with no byte-order mark. EdSharp's
// default for text is UTF-8 with BOM ("utf8b"), so pure-ASCII, BOM-less
// UTF-8, undetectable, or empty content all resolve to utf8b. A clearly
// detected legacy or wide encoding (windows-1252, UTF-16 without BOM,
// Shift-JIS, ...) is honored so the file is read -- and later saved --
// without corruption. Detection uses the Ude charset detector when the
// Ude.dll library is present at build time (the HAVEUDE symbol); without
// it, detection degrades to the utf8b default.
Encoding enUtf8b = new UTF8Encoding(true);
#if HAVEUDE
try {
byte[] aBytes = System.IO.File.ReadAllBytes(sFile);
if (aBytes.Length == 0) return enUtf8b;
Ude.CharsetDetector charsetDetector = new Ude.CharsetDetector();
charsetDetector.Feed(aBytes, 0, aBytes.Length);
charsetDetector.DataEnd();
string sCharset = charsetDetector.Charset;
if (String.IsNullOrEmpty(sCharset)) return enUtf8b;
Encoding enDetected = CharsetName2Encoding(sCharset, enUtf8b);
// A sanity check on the heuristic, from Scott's plain-text conversion of
// 25 August 2026: Ude reported UTF-16 for ordinary single-byte text with
// no byte-order mark, and reading it that way fused every two letters
// into one far-eastern character. Detection is guesswork, but this part
// is arithmetic -- every Latin letter in UTF-16 carries a zero byte, so
// content without zero bytes CANNOT be UTF-16 or UTF-32, whatever the
// detector says. When the two disagree, the arithmetic wins.
if (enDetected == Encoding.Unicode || enDetected == Encoding.BigEndianUnicode || enDetected == Encoding.UTF32) {
int iSample = Math.Min(aBytes.Length, 4096);
bool bAnyZero = false;
for (int i = 0; i < iSample; i++) if (aBytes[i] == 0) { bAnyZero = true; break; }
if (!bAnyZero) return enUtf8b;
}
return enDetected;
}
catch { return enUtf8b; }
#else
return enUtf8b;
#endif
} // DetectEncodingNoBom method

public static Encoding CharsetName2Encoding(string sName, Encoding enDefault) {
// Map a detector charset name to a .NET Encoding. ASCII and BOM-less UTF-8
// fold to the caller's default (utf8b) for the certainty of a BOM; UTF-16
// and UTF-32 map to their .NET encodings; anything else is looked up by
// name so legacy code pages round-trip.
string sKey = sName.Trim().Replace("-", "").Replace("_", "").ToLower();
if (sKey == "ascii" || sKey == "usascii" || sKey == "utf8") return enDefault;
if (sKey == "utf16le" || sKey == "utf16" || sKey == "unicode") return Encoding.Unicode;
if (sKey == "utf16be") return Encoding.BigEndianUnicode;
if (sKey == "utf32" || sKey == "utf32le") return Encoding.UTF32;
try { return Encoding.GetEncoding(sName); }
catch { return enDefault; }
} // CharsetName2Encoding method

public static string FetchLatestReleaseTag(string sOwnerRepo) {
// Return the tag of the latest GitHub release for "owner/repo", e.g. "v5.0.0".
// The public REST API is tried first (no credentials needed); on any failure
// the releases/latest page is fetched and its post-redirect address, which
// ends in the tag, is used instead.  Returns "" if neither path yields a tag.
// Homer.Web supplies the User-Agent header and modern TLS that GitHub needs.
string sApiUrl = "https://api.github.com/repos/" + sOwnerRepo + "/releases/latest";
string sFinalUrl = "";
string sJson = Homer.Web.getPage(sApiUrl, out sFinalUrl);
if (sJson.Length > 0) {
Match matchTag = Regex.Match(sJson, "\"tag_name\"\\s*:\\s*\"([^\"]+)\"");
if (matchTag.Success) return matchTag.Groups[1].Value;
}
string sPageUrl = "https://github.com/" + sOwnerRepo + "/releases/latest";
Homer.Web.getPage(sPageUrl, out sFinalUrl);
string sRedirect = (sFinalUrl == null ? "" : sFinalUrl).TrimEnd('/');
int iSlash = sRedirect.LastIndexOf('/');
if (iSlash >= 0 && iSlash < sRedirect.Length - 1) {
string sTag = sRedirect.Substring(iSlash + 1);
if (!sTag.Equals("latest", StringComparison.OrdinalIgnoreCase)) return sTag;
}
return "";
} // FetchLatestReleaseTag method

public static int CompareVersions(string sA, string sB) {
// Compare two dotted-numeric version strings, e.g. "5.0.1" versus "5.0.0".
// Returns a negative value if sA is older, zero if equal, positive if newer.
// Missing trailing parts count as zero (so "5.0" equals "5.0.0"); any part
// that is not an integer falls back to ordinal string comparison.
string[] asA = (sA == null ? "" : sA).Split('.');
string[] asB = (sB == null ? "" : sB).Split('.');
int iCount = Math.Max(asA.Length, asB.Length);
int iIndex, iValA, iValB;
for (iIndex = 0; iIndex < iCount; iIndex++) {
string sPartA = iIndex < asA.Length ? asA[iIndex].Trim() : "0";
string sPartB = iIndex < asB.Length ? asB[iIndex].Trim() : "0";
bool bNumA = int.TryParse(sPartA, out iValA);
bool bNumB = int.TryParse(sPartB, out iValB);
if (bNumA && bNumB) { if (iValA != iValB) return iValA - iValB; }
else { int iCmp = string.CompareOrdinal(sPartA, sPartB); if (iCmp != 0) return iCmp; }
}
return 0;
} // CompareVersions method

public static string GetClipboardText() {
// Read text from the clipboard defensively.  The clipboard is a shared resource
// that another process can briefly lock; on Windows 11 (cloud clipboard and
// clipboard history) GetText throws an ExternalException intermittently when
// that happens.  Retry a few times, then give up quietly with "" rather than
// letting the exception crash EdSharp (the Append From Clipboard viewer reads
// the clipboard from inside WndProc, where an unhandled throw is fatal).
int iTry;
for (iTry = 0; iTry < 10; iTry++) {
try {
if (Clipboard.ContainsText()) return Clipboard.GetText();
return "";
}
catch (Exception) {
System.Threading.Thread.Sleep(40);
}
}
return "";
} // GetClipboardText method

public static void SetClipboardText(string sText) {
// Write text to the clipboard defensively, with the same retry-on-contention
// logic as GetClipboardText.  An empty or null string clears the clipboard
// instead of throwing (Clipboard.SetText rejects the empty string).
int iTry;
for (iTry = 0; iTry < 10; iTry++) {
try {
if (sText == null || sText.Length == 0) Clipboard.Clear();
else Clipboard.SetText(sText);
return;
}
catch (Exception) {
System.Threading.Thread.Sleep(40);
}
}
} // SetClipboardText method

// The official Python from python.org, never the Microsoft Store stub.
// The stub lives under WindowsApps and answers when asked, so any path
// there is rejected; the usual python.org locations are searched next,
// newest version first, and whatever is chosen must actually answer
// --version. Returns an empty string when nothing real is installed, so
// the caller can fall back to the bare name.
public static string FindPythonPath() {
try {
string sPath = Environment.GetEnvironmentVariable("PATH");
if (sPath != null) {
foreach (string sDir in sPath.Split(';')) {
if (sDir.Trim().Length == 0) continue;
if (sDir.IndexOf(@"\WindowsApps", StringComparison.OrdinalIgnoreCase) >= 0) continue;
string sTry = Path.Combine(sDir.Trim(), "python.exe");
if (File.Exists(sTry)) return sTry;
}
}
List<string> lsRoots = new List<string>();
lsRoots.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles));
lsRoots.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python"));
// The python.org installer's own default when it installs for all users
// is a folder straight off the drive root, such as C:\Python314.
lsRoots.Add(Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.System)));
List<string> lsFound = new List<string>();
foreach (string sRoot in lsRoots) {
if (!Directory.Exists(sRoot)) continue;
foreach (string sDir in Directory.GetDirectories(sRoot, "Python3*")) {
string sTry = Path.Combine(sDir, "python.exe");
if (File.Exists(sTry)) lsFound.Add(sTry);
}
}
lsFound.Sort();
lsFound.Reverse();
if (lsFound.Count > 0) return lsFound[0];
}
catch (Exception) {}
return "";
} // FindPythonPath method

public static string FindCscPath() {
// Locate a C# compiler: prefer the newest Roslyn csc (from VS Build Tools, for
// the latest C# language version), then fall back to the csc.exe that ships
// with the running .NET Framework, which is always present. Returns "" if none.
string sWin = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
// 64-bit first throughout, per the standing preference: the 64-bit
// Program Files copies of the Visual Studio Roslyn compiler come before
// their 32-bit twins, and the Framework64 compiler that ships with
// every Windows 11 comes before the 32-bit Framework one. The running
// runtime directory sits between them, since a 64-bit EdSharp is itself
// running on the 64-bit framework.
string[] aCandidates = new string[] {
@"C:\Program Files\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe",
@"C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\Roslyn\csc.exe",
@"C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\Roslyn\csc.exe",
@"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe",
Path.Combine(sWin, @"Microsoft.NET\Framework64\v4.0.30319\csc.exe"),
Path.Combine(System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "csc.exe"),
@"C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe",
@"C:\Program Files (x86)\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\Roslyn\csc.exe",
@"C:\Program Files (x86)\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\Roslyn\csc.exe",
@"C:\Program Files (x86)\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe",
Path.Combine(sWin, @"Microsoft.NET\Framework\v4.0.30319\csc.exe")
};
foreach (string s in aCandidates) {
try { if (File.Exists(s)) return s; } catch {}
}
return "";
} // FindCscPath method

public static void GetGoogleLanguages(out string[] aLanguageNames, out string[] aLanguageAbbreviations) {
List<string[]> lLanguages = new List<string[]>();
lLanguages.Add(new string[] {"AFRIKAANS", "af"});
lLanguages.Add(new string[] {"ALBANIAN", "sq"});
lLanguages.Add(new string[] {"AMHARIC", "am"});
lLanguages.Add(new string[] {"ARABIC", "ar"});
lLanguages.Add(new string[] {"ARMENIAN", "hy"});
lLanguages.Add(new string[] {"AZERBAIJANI", "az"});
lLanguages.Add(new string[] {"BASQUE", "eu"});
lLanguages.Add(new string[] {"BELARUSIAN", "be"});
lLanguages.Add(new string[] {"BENGALI", "bn"});
lLanguages.Add(new string[] {"BIHARI", "bh"});
lLanguages.Add(new string[] {"BULGARIAN", "bg"});
lLanguages.Add(new string[] {"BURMESE", "my"});
lLanguages.Add(new string[] {"CATALAN", "ca"});
lLanguages.Add(new string[] {"CHEROKEE", "chr"});
lLanguages.Add(new string[] {"CHINESE", "zh"});
lLanguages.Add(new string[] {"CHINESE_SIMPLIFIED", "zh-CN"});
lLanguages.Add(new string[] {"CHINESE_TRADITIONAL", "zh-TW"});
lLanguages.Add(new string[] {"CROATIAN", "hr"});
lLanguages.Add(new string[] {"CZECH", "cs"});
lLanguages.Add(new string[] {"DANISH", "da"});
lLanguages.Add(new string[] {"DHIVEHI", "dv"});
lLanguages.Add(new string[] {"DUTCH", "nl"});
lLanguages.Add(new string[] {"ENGLISH", "en"});
lLanguages.Add(new string[] {"ESPERANTO", "eo"});
lLanguages.Add(new string[] {"ESTONIAN", "et"});
lLanguages.Add(new string[] {"FILIPINO", "tl"});
lLanguages.Add(new string[] {"FINNISH", "fi"});
lLanguages.Add(new string[] {"FRENCH", "fr"});
lLanguages.Add(new string[] {"GALICIAN", "gl"});
lLanguages.Add(new string[] {"GEORGIAN", "ka"});
lLanguages.Add(new string[] {"GERMAN", "de"});
lLanguages.Add(new string[] {"GREEK", "el"});
lLanguages.Add(new string[] {"GUARANI", "gn"});
lLanguages.Add(new string[] {"GUJARATI", "gu"});
lLanguages.Add(new string[] {"HEBREW", "iw"});
lLanguages.Add(new string[] {"HINDI", "hi"});
lLanguages.Add(new string[] {"HUNGARIAN", "hu"});
lLanguages.Add(new string[] {"ICELANDIC", "is"});
lLanguages.Add(new string[] {"INDONESIAN", "id"});
lLanguages.Add(new string[] {"INUKTITUT", "iu"});
lLanguages.Add(new string[] {"ITALIAN", "it"});
lLanguages.Add(new string[] {"JAPANESE", "ja"});
lLanguages.Add(new string[] {"KANNADA", "kn"});
lLanguages.Add(new string[] {"KAZAKH", "kk"});
lLanguages.Add(new string[] {"KHMER", "km"});
lLanguages.Add(new string[] {"KOREAN", "ko"});
lLanguages.Add(new string[] {"KURDISH", "ku"});
lLanguages.Add(new string[] {"KYRGYZ", "ky"});
lLanguages.Add(new string[] {"LAOTHIAN", "lo"});
lLanguages.Add(new string[] {"LATVIAN", "lv"});
lLanguages.Add(new string[] {"LITHUANIAN", "lt"});
lLanguages.Add(new string[] {"MACEDONIAN", "mk"});
lLanguages.Add(new string[] {"MALAY", "ms"});
lLanguages.Add(new string[] {"MALAYALAM", "ml"});
lLanguages.Add(new string[] {"MALTESE", "mt"});
lLanguages.Add(new string[] {"MARATHI", "mr"});
lLanguages.Add(new string[] {"MONGOLIAN", "mn"});
lLanguages.Add(new string[] {"NEPALI", "ne"});
lLanguages.Add(new string[] {"NORWEGIAN", "no"});
lLanguages.Add(new string[] {"ORIYA", "or"});
lLanguages.Add(new string[] {"PASHTO", "ps"});
lLanguages.Add(new string[] {"PERSIAN", "fa"});
lLanguages.Add(new string[] {"POLISH", "pl"});
lLanguages.Add(new string[] {"PORTUGUESE", "pt-PT"});
lLanguages.Add(new string[] {"PUNJABI", "pa"});
lLanguages.Add(new string[] {"ROMANIAN", "ro"});
lLanguages.Add(new string[] {"RUSSIAN", "ru"});
lLanguages.Add(new string[] {"SANSKRIT", "sa"});
lLanguages.Add(new string[] {"SERBIAN", "sr"});
lLanguages.Add(new string[] {"SINDHI", "sd"});
lLanguages.Add(new string[] {"SINHALESE", "si"});
lLanguages.Add(new string[] {"SLOVAK", "sk"});
lLanguages.Add(new string[] {"SLOVENIAN", "sl"});
lLanguages.Add(new string[] {"SPANISH", "es"});
lLanguages.Add(new string[] {"SWAHILI", "sw"});
lLanguages.Add(new string[] {"SWEDISH", "sv"});
lLanguages.Add(new string[] {"TAJIK", "tg"});
lLanguages.Add(new string[] {"TAMIL", "ta"});
lLanguages.Add(new string[] {"TAGALOG", "tl"});
lLanguages.Add(new string[] {"TELUGU", "te"});
lLanguages.Add(new string[] {"THAI", "th"});
lLanguages.Add(new string[] {"TIBETAN", "bo"});
lLanguages.Add(new string[] {"TURKISH", "tr"});
lLanguages.Add(new string[] {"UKRAINIAN", "uk"});
lLanguages.Add(new string[] {"URDU", "ur"});
lLanguages.Add(new string[] {"UZBEK", "uz"});
lLanguages.Add(new string[] {"UIGHUR", "ug"});
lLanguages.Add(new string[] {"VIETNAMESE", "vi"});
lLanguages.Add(new string[] {"UNKNOWN", ""});

int iCount = lLanguages.Count;
aLanguageNames = new string[iCount];
aLanguageAbbreviations = new string[iCount];
for (int i = 0; i < iCount; i++) {
string[] a = lLanguages[i];
aLanguageNames[i] = a[0];
aLanguageAbbreviations[i] = a[1];
};
} // GetGoogleLanguages method

public static string[] OldGetGoogleLanguages() {
HomerList hl = new HomerList();
hl.Add("Arabic");
hl.Add("Bulgarian");
hl.Add("Chinese");
hl.Add("Catalan");
hl.Add("Croatian");
hl.Add("Czech");
hl.Add("Danish");
hl.Add("Dutch");
hl.Add("English");
hl.Add("Filipino");
hl.Add("Finnish");
hl.Add("French");
hl.Add("German");
hl.Add("Greek");
hl.Add("Hebrew");
hl.Add("Hindi");
hl.Add("Indonesian");
hl.Add("Italian");
hl.Add("Japanese");
hl.Add("Korean");
hl.Add("Latvian");
hl.Add("Lithuanian");
hl.Add("Norwegian");
hl.Add("Polish");
hl.Add("Portuguese");
hl.Add("Romanian");
hl.Add("Russian");
hl.Add("Spanish");
hl.Add("Serbian");
hl.Add("Slovak");
hl.Add("Slovenian");
hl.Add("Swedish");
hl.Add("Turkish");
hl.Add("Ukrainian");
hl.Add("Vietnamese");
hl.Add("Unknown");
return hl.ToArray();
} // OldGetGoogleLanguages method

public static bool MailMessage(string sRecipient, string sSubject, string sBody) {
sBody = Util.RegExpReplaceCase(sBody, "\r\n", "\r");
sBody = Util.RegExpReplaceCase(sBody, "\n", "\r");
sBody = Util.RegExpReplaceCase(sBody, "\r", "\r\n");
sBody = Util.RegExpReplaceCase(sBody, "\r\n", "%0D%0A");
sBody = Util.RegExpReplaceCase(sBody, " ", "%20");
sBody = Util.RegExpReplaceCase(sBody, "\t", "%09");
sBody = Util.RegExpReplaceCase(sBody, "\"", "%22");
sBody = Util.RegExpReplaceCase(sBody, "'", "%27");
sBody = Util.RegExpReplaceCase(sBody, "\\\\", "%5C");
// sBody = StringReplaceCase(sBody, "\\", "%5C");
// string sCommand = "mailto:?BODY=" + sBody;
string sCommand = "mailto:" + sRecipient + "?SUBJECT=" + sSubject + "&BODY=" + sBody;
bool bResult;
try {
Process.Start(sCommand);
bResult = true;
}
catch (Exception ex) {
Dialog.Show("Error", ex.Message);
bResult = false;
}
return bResult;
} // MailMessage method

public static Encoding OldGetFileEncoding(string sFile) {
System.Text.Encoding enc = null;
System.IO.FileStream file = new System.IO.FileStream(sFile,
FileMode.Open, FileAccess.Read, FileShare.Read);
if (file.CanSeek)
{
byte[] bom = new byte[4]; // Get the byte-order mark, if there is one
file.Read(bom, 0, 4);
if ((bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf) || // utf-8
(bom[0] == 0xff && bom[1] == 0xfe) || // ucs-2le, ucs-4le, and ucs-16le
(bom[0] == 0xfe && bom[1] == 0xff) || // utf-16 and ucs-2
(bom[0] == 0 && bom[1] == 0 && bom[2] == 0xfe && bom[3] == 0xff)) // ucs-4
{
enc = System.Text.Encoding.Unicode;
}
else
{
// enc = System.Text.Encoding.ASCII;
enc = System.Text.Encoding.Default;
}

// Now reposition the file cursor back to the start of the file
file.Seek(0, System.IO.SeekOrigin.Begin);
}
else
{
// The file cannot be randomly accessed, so you need to decide what to set the default to
// based on the data provided. If you're expecting data from a lot of older applications,
// default your encoding to Encoding.ASCII. If you're expecting data from a lot of newer
// applications, default your encoding to Encoding.Unicode. Also, since binary files are
// single byte-based, so you will want to use Encoding.ASCII, even though you'll probably
// never need to use the encoding then since the Encoding classes are really meant to get
// strings from the byte array that is the file.

// enc = System.Text.Encoding.ASCII;
enc = System.Text.Encoding.Default;
}
file.Close();
return enc;
} // OldGetFileEncoding method

[DllImport("user32.dll")]
public static extern IntPtr SetClipboardViewer(IntPtr h);

[DllImport("user32.dll")]
public static extern IntPtr     ChangeClipboardChain(IntPtr hCurrentClipboardViewer, IntPtr hNextClipboardViewer);

public static string GetTempFolder() {
object oSystem = COM.CreateObject("Scripting.FileSystemObject");
Object oDir = COM.CallMethod(oSystem, "GetSpecialFolder", new object[] {2});
string sPath = (string) COM.GetProperty(oDir, "Path");
return sPath;
} // GetTempFolder method

public static bool IsUnicode(string sText) {
foreach (char c in sText) {
// if (((int) c) > 255) Dialog.Show(c, (int) c);
if (((int) c) > 255) return true;
}
return false;
} // IsUnicode method

public static string Replicate(string sText, int iCount) {
string sReturn = sText;
for (int i = 1; i < iCount; i++) sReturn += sText;
return sReturn;
} // Replicate method

public static bool IsAppActiveWindow() {
IntPtr h = Win32.GetForegroundWindow();
foreach (Form frm in Application.OpenForms) if (frm.Handle == h) return true;
return false;
} // IsAppActiveWindow method

public static bool Spell(object oText) {
string sText = oText.ToString();
bool bReturn = false;
string sReturn = "";
for (int i = 0; i < sText.Length; i++) {
string s = sText.Substring(i, 1);
switch (s) {
case " " :
s = "Space";
sReturn += "Space\n";
break;
default :
sReturn += s + "\n";
break;
}
s = " " + s + " ";
bReturn = Say(s);
}
// return Say(sReturn, bGlobal);
return bReturn;
} // Spell method

public static bool Say(object oText) {
bool bGlobal = false;
return Say(oText, bGlobal);
} // Say method

// WHAT EDSHARP SPEAKS, AND WHAT IT LEAVES TO THE SCREEN READER.
//
// A screen reader already announces, on its own: the title of a window
// that opens or gains focus, and the name, role, state and value of the
// control that receives keyboard focus. JAWS, NVDA and Narrator all do
// this from the accessibility information Windows publishes, without
// being asked. So EdSharp must NOT speak those things itself: doing so
// produces the doubled announcements that make a program tiring to use.
//
// What EdSharp does speak is what only EdSharp knows: the ANSWER to a
// command. "3 changes, 2 skips", "Level 2", "In 1", the text a
// navigation command lands on, a count, a warning. Those are the
// command's own output, not a description of the interface, and they
// are exactly what this method is for.
//
// One message, one voice: Homer.Say tries JAWS, then NVDA, then a native
// UIA notification, and STOPS at the first that answers -- so a message
// is never delivered twice by two mechanisms.
public static bool Say(object oText, bool bGlobal) {
string sText = oText.ToString();
if (sText.Trim().Length == 0) sText = "Blank";
if (!App.ExtraSpeech) {
Util.StringAppend2File(sText + "\r\n", App.SpeechLog);
return false;
}

if (!bGlobal) {
if (!IsAppActiveWindow()) return false;

if ((Control.ModifierKeys & Keys.Alt) != 0 && (Control.ModifierKeys & Keys.Control) != 0) return false;
}

// Speech goes through Homer.Say: JAWS COM, then the NVDA controller client,
// then a native UIA notification reaching Narrator and any other UIA-listening
// reader.  The former per-reader COM/Win32 chain has been removed.
Homer.Say.sayForced(sText);
return true;
} // Say method

public static string Key2String(Keys keyData) {
return TypeDescriptor.GetConverter(typeof(Keys)).ConvertToString(keyData);
} // Key2String method

public static Keys String2Key(string sKey) {
// Keys.None for a blank key: menu-only commands are part of the design --
// KeyMap models them as unbound -- and the converter returns null for an
// empty string, which the old cast turned into the silent startup crash of
// 19 August 2026 (four keyless conversion menu items, added with the
// Markdown features, were the first blank keys ever to reach this method).
// The conversion itself is unchanged.
if (sKey == null) return Keys.None;
sKey = sKey.Trim();
if (sKey.Length == 0) return Keys.None;
object oKey = TypeDescriptor.GetConverter(typeof(Keys)).ConvertFromString(sKey);
if (oKey == null) return Keys.None;
return (Keys) oKey;
} // String2Key method

public static string Font2String(Font font) {
return TypeDescriptor.GetConverter(typeof(Font)).ConvertToString(font);
} // Font2String method

public static Font String2Font(string sFont) {
return (Font) TypeDescriptor.GetConverter(typeof(Font)).ConvertFromString(sFont);
} // String2Font method

public static string Color2String(Color color) {
return TypeDescriptor.GetConverter(typeof(Color)).ConvertToString(color);
} // Color2String method

public static Color String2Color(string sColor) {
return (Color) TypeDescriptor.GetConverter(typeof(Color)).ConvertFromString(sColor);
} // String2Color method

public static string GetFriendlyKeyName(string sKey) {
if (sKey.Contains("+OemQuotes")) sKey = sKey.Replace("+OemQuotes", "+Apostrophe");
if (sKey.Contains("+Back")) sKey = sKey.Replace("+Back", "+Backspace");
if (sKey.Contains("+Oem5")) sKey = sKey.Replace("+Oem5", "+Backslash");
if (sKey.Contains("+Oemplus")) sKey = sKey.Replace("+Oemplus", "+Equals");
if (sKey.Contains("+OemMinus")) sKey = sKey.Replace("+OemMinus", "+Dash");
if (sKey.Contains("+OemSemicolon")) sKey = sKey.Replace("+OemSemicolon", "+Semicolon");
if (sKey.Contains("+D6")) sKey = sKey.Replace("+D6", "+Caret");
if (sKey.Contains("+D0")) sKey = sKey.Replace("+D0", "+0");
if (sKey.Contains("+OemQuestion")) sKey = sKey.Replace("+OemQuestion", "+Slash");
if (sKey.Contains("+OemOpenBrackets")) sKey = sKey.Replace("+OemOpenBrackets", "+LeftBracket");
if (sKey.Contains("+OemCloseBrackets")) sKey = sKey.Replace("+OemCloseBrackets", "+RightBracket");
if (sKey.Contains("+Oemcomma")) sKey = sKey.Replace("+Oemcomma", "+Comma");
if (sKey.Contains("+OemPeriod")) sKey = sKey.Replace("+OemPeriod", "+Period");
return sKey;
} // GetFriendlyKeyName method

public static string RegExpReplaceEquiv(string sText, string sMatch, string sReplace) {
bool bCaseSensitive = false;
return RegExpReplace(sText, sMatch, sReplace, bCaseSensitive);
} // RegExpReplaceEquiv method

public static string RegExpReplaceCase(string sText, string sMatch, string sReplace) {
bool bCaseSensitive = true;
return RegExpReplace(sText, sMatch, sReplace, bCaseSensitive);
} // RegExpReplaceCase method

public static string RegExpReplace(string sText, string sMatch, string sReplace, bool bCaseSensitive) {
RegexOptions options = RegexOptions.Multiline;
if (!bCaseSensitive) options |= RegexOptions.IgnoreCase;
Regex rx = new Regex(sMatch, options);
string sReturn = rx.Replace(sText, sReplace);
return sReturn;
} // RegExpReplace method

public static int RegExpCountEquiv(string sText, string sMatch) {
bool bCaseSensitive = false;
return RegExpCount(sText, sMatch, bCaseSensitive);
} // RegExpCountEquiv method

public static int RegExpCountCase(string sText, string sMatch) {
bool bCaseSensitive = true;
return RegExpCount(sText, sMatch, bCaseSensitive);
} // RegExpCountCase method

public static int RegExpCount(string sText, string sMatch, bool bCaseSensitive) {
RegexOptions options = RegexOptions.Multiline;
if (!bCaseSensitive) options |= RegexOptions.IgnoreCase;
Regex rx = new Regex(sMatch, options);
MatchCollection matches = rx.Matches(sText);
int iReturn = matches.Count;
return iReturn;
} // RegExpCount method

public static string[] RegExpExtractEquiv(string sText, string sMatch) {
bool bCaseSensitive = false;
return RegExpExtract(sText, sMatch, bCaseSensitive);
} // RegExpExtractEquiv method

public static string[] RegExpExtractCase(string sText, string sMatch) {
bool bCaseSensitive = true;
return RegExpExtract(sText, sMatch, bCaseSensitive);
} // RegExpExtractCase method

public static string[] RegExpExtract(string sText, string sMatch, bool bCaseSensitive) {
RegexOptions options = RegexOptions.Multiline;
if (!bCaseSensitive) options |= RegexOptions.IgnoreCase;
Regex rx = new Regex(sMatch, options);
MatchCollection matches = rx.Matches(sText);
string[] aReturn = new string[matches.Count];
for (int i = 0; i < aReturn.Length; i++) aReturn[i] = matches[i].Value;
return aReturn;
} // RegExpExtract method

public static object[][] RegExpExtractWithIndex(string sText, string sMatch, bool bCaseSensitive) {
RegexOptions options = RegexOptions.Multiline;
if (!bCaseSensitive) options |= RegexOptions.IgnoreCase;
Regex rx = new Regex(sMatch, options);
MatchCollection matches = rx.Matches(sText);
object[][] aReturn = new object[matches.Count][];
for (int i = 0; i < aReturn.Length; i++) aReturn[i] = new object[] {matches[i].Index, matches[i].Value};
return aReturn;
} // RegExpExtractWithIndex method

public static object[] RegExpContainsEquiv(string sText, string sMatch) {
int iStart = 0;
return RegExpContainsEquiv(sText, sMatch, iStart);
} // RegExpContainsEquiv method

public static object[] RegExpContainsEquiv(string sText, string sMatch, int iStart) {
bool bCaseSensitive = false;
bool bLast = false;
return RegExpContains(sText, sMatch, bCaseSensitive, bLast, iStart);
} // RegExpContainsEquiv method

public static object[] RegExpContainsCase(string sText, string sMatch) {
int iStart = 0;
return RegExpContainsCase(sText, sMatch, iStart);
} // RegExpContainsCase method

public static object[] RegExpContainsCase(string sText, string sMatch, int iStart) {
bool bCaseSensitive = true;
bool bLast = false;
return RegExpContains(sText, sMatch, bCaseSensitive, bLast, iStart);
} // RegExpContainsCase method

public static object[] RegExpContainsLastEquiv(string sText, string sMatch) {
bool bCaseSensitive = false;
bool bLast = true;
return RegExpContains(sText, sMatch, bCaseSensitive, bLast);
} // RegExpContainsLastEquiv method

public static object[] RegExpContainsLastCase(string sText, string sMatch) {
bool bCaseSensitive = true;
bool bLast = true;
return RegExpContains(sText, sMatch, bCaseSensitive, bLast);
} // RegExpContainsLastCase method

public static object[] RegExpContains(string sText, string sMatch, bool bCaseSensitive, bool bLast) {
int iStart;
//if (bLast)  iStart = sText.Length - 1;
if (bLast)  iStart = sText.Length;
else iStart = 0;
return RegExpContains(sText, sMatch, bCaseSensitive, bLast, iStart);
} // RegExpContains method

public static object[] RegExpContains(string sText, string sMatch, bool bCaseSensitive, bool bLast, int iStart) {
RegexOptions options = RegexOptions.Multiline;
if (!bCaseSensitive) options |= RegexOptions.IgnoreCase;
if (bLast) options |= RegexOptions.RightToLeft;
Regex rx = new Regex(sMatch, options);

Match match = rx.Match(sText, iStart);
object[] aReturn = new object[] {-1, null};
// Dialog.Show(match.Success);
if (match.Success) aReturn = new object[] {match.Index, match.Value};
return aReturn;
} // RegExpContains method

public static bool Equiv(string s1, string s2) {
return String.Compare(s1, s2, true) == 0;
} // Equiv method

public static bool ToBool(string sValue) {
// Liberal truth test for .inix-style flags, matching Regexer's convention.
if (sValue == null) return false;
sValue = sValue.Trim().ToLower();
return (sValue == "true" || sValue == "yes" || sValue == "on" || sValue == "1" || sValue == "y");
} // ToBool method

public static RegexOptions RegexOptionsFromString(string sOptions) {
// Build a RegexOptions value from a comma-separated list of .NET option names,
// as Regexer does for each [Section]'s Options key. Unknown names are ignored;
// blank yields RegexOptions.None.
RegexOptions options = RegexOptions.None;
if (sOptions == null) return options;
string[] aOptions = sOptions.Replace(" ", "").ToLower().Split(',');
foreach (string sOption in aOptions) {
if (sOption == "compiled") options = options | RegexOptions.Compiled;
else if (sOption == "cultureinvariant") options = options | RegexOptions.CultureInvariant;
else if (sOption == "ecmascript") options = options | RegexOptions.ECMAScript;
else if (sOption == "explicitcapture") options = options | RegexOptions.ExplicitCapture;
else if (sOption == "ignorecase") options = options | RegexOptions.IgnoreCase;
else if (sOption == "ignorepatternwhitespace") options = options | RegexOptions.IgnorePatternWhitespace;
else if (sOption == "multiline") options = options | RegexOptions.Multiline;
else if (sOption == "righttoleft") options = options | RegexOptions.RightToLeft;
else if (sOption == "singleline") options = options | RegexOptions.Singleline;
}
return options;
} // RegexOptionsFromString method

public static string Pluralize(int iCount, string sSingular) {
string sPlural = null;
return Pluralize(iCount, sSingular, sPlural);
} // Pluralize method

public static string Pluralize(int iCount, string sSingular, string sPlural) {
if (sPlural == null) sPlural = sSingular + "s";
string sReturn = iCount.ToString() + " ";
if (iCount == 1) sReturn += sSingular;
else sReturn += sPlural;
return sReturn;
} // Pluralize method

public static string File2String(string sFile) {
Encoding en = null;
return File2String(sFile, ref en);
} // File2String method

public static string File2String(string sFile, ref Encoding en) {
//return System.IO.File.ReadAllText(sFile);
// Dialog.Show("Encoding", Util.GetFileEncoding(sFile));
if (en == null) en = Util.GetFileEncoding(sFile, App.BomDictionary);
// return System.IO.File.ReadAllText(sFile, System.Text.Encoding.Default);
// return System.IO.File.ReadAllText(sFile, encoding);
string sText = System.IO.File.ReadAllText(sFile, en);
return sText;
} // File2String method

public static string OldFile2String(string sFile) {
if (!File.Exists(sFile)) return "";
StreamReader textReader = new StreamReader(sFile);
string sBody = textReader.ReadToEnd();
textReader.Close();
return sBody;
} // OldFile2String method

public static void String2FileU(string sBody, string sFile) {
// bool bAppend = false;
System.IO.File.WriteAllText(sFile, sBody, Encoding.UTF8);
} // String2FileU method

public static void StringAppend2File(string sBody, string sFile) {
File.AppendAllText(sFile, sBody);
} // StringAppend2File method

public static void String2FileA(string sBody, string sFile) {
Encoding en = null;
String2FileA(sBody, sFile, en);
} // String2FileA method

public static void String2FileA(string sBody, string sFile, Encoding en) {
StreamWriter textWriter = new StreamWriter(sFile);
textWriter.Write(sBody);
textWriter.Close();
} // OldString2File method

public static void String2File(string sBody, string sFile) {
Encoding en = null;
String2File(sBody, sFile, ref en);
} // String2File method

public static void String2File(string sBody, string sFile, ref Encoding en) {
// Dialog.Show(IsUnicode(sBody));
if (en != null) {}
// Do nothing
else if (IsUnicode(sBody))en = Encoding.UTF8;
else en = Encoding.Default;

// sBody = Util.Convert2WinLineBreak(sBody);
File.WriteAllText(sFile, sBody, en);
} // String2File method

public static string Quote(string sText) {
return "\"" + Unquote(sText) + "\"";
} // Quote method

public static string Unquote(string sText) {
return sText.Trim('"');
} // Unquote method

public static string ConvertQuotes(string sText) {
string sReturn = sText.Replace(@"?", @"""");
sReturn  = sReturn.Replace(@"?", @"""");
sReturn  = sReturn.Replace(@"-", @"-");
sReturn  = sReturn.Replace(@"?", @"...");
sReturn  = sReturn.Replace(@"?", @"'");
sReturn  = sReturn.Replace(Util.Code2String(65533), @"'");
return sReturn;
} // ConvertQuotes method

public static string Convert2Ascii(string sText) {
int iLength = sText.Length;
for (int i = iLength - 1; i >= 0; i--) {
if ((int) sText[i] > 127) sText = sText.Remove(i, 1);
}
return sText;
} //Convert2Ascii method

public static string Convert2MacLineBreak(string sText) {
//Convert to Macintosh line break, \r;
string sMatch, sReplace;

sMatch = "\r\n";
sReplace = "\r";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
sMatch = "\n";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
return sText;
} // Convert2MacLineBreak metod

public static string Convert2UnixLineBreak(string sText) {
//Convert to Unix line break, \n;
string sMatch, sReplace;
sMatch = "\r\n";
sReplace = "\n";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
sMatch = "\r";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
return sText;
} // Convert2UnixLineBreak method

public static string Convert2WinLineBreak(string sText) {
//Convert to standard Windows line break, \r\nVar;
string sMatch, sReplace;
sMatch = "\r\n";
sReplace = "\n";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
sMatch = "\r";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
sMatch = "\n";
sReplace = "\r\n";
sText = Util.RegExpReplaceCase(sText, sMatch, sReplace);
return sText;
} // Convert2WinLineBreakMethod

// One line to the session log, timestamped. Never throws; a missing or
// locked log must not affect the work being logged.
public static void Log(string sMessage) {
if (App.LogFile == null || App.LogFile.Length == 0) return;
try { File.AppendAllText(App.LogFile, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + sMessage + "\r\n"); }
catch (Exception) {}
} // Log method

public static int RunHideWait(string sPath) {
return runShell(sPath, ProcessWindowStyle.Hidden, true);
} //RunHideWait method

public static int RunHide(string sPath) {
return runShell(sPath, ProcessWindowStyle.Hidden, false);
} //RunHide method

public static int Run(string sPath) {
return runShell(sPath, ProcessWindowStyle.Normal, false);
} //Run method

public static int RunWait(string sPath) {
return runShell(sPath, ProcessWindowStyle.Normal, true);
} //RunWait method

// runShell: launch a command line as a process (replaces VB Shell). The
// command line is split into executable and arguments so window style is
// honored on the child without a cmd wrapper. Returns the process id, or 0.
static int runShell(string sCommand, ProcessWindowStyle style, bool bWait) {
string sExe, sArgs;
sCommand = sCommand.Trim();
if (sCommand.StartsWith("\"")) {
int iEnd = sCommand.IndexOf('\"', 1);
sExe = sCommand.Substring(1, iEnd - 1);
sArgs = sCommand.Substring(iEnd + 1).Trim();
}
else {
int iSpace = sCommand.IndexOf(' ');
if (iSpace < 0) { sExe = sCommand; sArgs = ""; }
else { sExe = sCommand.Substring(0, iSpace); sArgs = sCommand.Substring(iSpace + 1).Trim(); }
}
ProcessStartInfo psi = new ProcessStartInfo(sExe, sArgs);
psi.UseShellExecute = false;
psi.WindowStyle = style;
psi.CreateNoWindow = (style == ProcessWindowStyle.Hidden);
Log((bWait ? "run and wait: " : "run: ") + sCommand);
Process p = Process.Start(psi);
if (p == null) Log("FAILED to start: " + sCommand);
if (bWait && p != null) {
// A converter that blocks -- a hidden Office dialog, a tool waiting for
// input it will never get -- used to freeze all of EdSharp forever,
// because this wait had no limit and runs on the interface thread (the
// hang suspected in the 22 August 2026 audit). Two minutes is generous
// for a big document. A process still running after that is ended, and
// the caller's own error dialog then shows the command line for
// diagnosis, which beats a frozen editor with no explanation.
const int c_iWaitMilliseconds = 120000;
if (!p.WaitForExit(c_iWaitMilliseconds)) {
Log("TIMEOUT after 120 seconds; process ended: " + sCommand);
try { p.Kill(); }
catch (Exception) {}
}
else Log("exit code " + p.ExitCode + ": " + sExe);
}
return (p != null) ? p.Id : 0;
} // runShell method

public static void ActivatePid(int iPid) {
Process p = Process.GetProcessById(iPid);
if (p != null && p.MainWindowHandle != IntPtr.Zero) Win32.SetForegroundWindow(p.MainWindowHandle);
} // ActivatePid method

public static bool ActivateProcess(string sProcess) {
Process[] processes = Process.GetProcessesByName(sProcess);
if (processes.Length == 0) return false;
Process process = processes[0];
//Dialog.Show(process.ProcessName, process.MainWindowTitle);

int iPid = processes[0].Id;
try {
ActivatePid(iPid);
return true;
}
catch {
return false;
}
} // ActivateProcess method

public static void ActivateTitle(string sTitle) {
IntPtr h = Win32.FindWindow(0, sTitle);
if ((int) h != 0) Win32.SetForegroundWindow(h);
} // ActivateTitle method

public static void Beep() {
System.Media.SystemSounds.Beep.Play();
} // Beep method

public static object If(bool bExp, object oTrue, object oFalse) {
return bExp ? oTrue : oFalse;
} // If method

public static int If(bool bExp, int iTrue, int iFalse) {
if (bExp) return iTrue;
else return iFalse;
} // If method

public static string If(bool bExp, string sTrue, string sFalse) {
if (bExp) return sTrue;
else return sFalse;
} // If method

public static string GetCommandLine() {
string[] aArgs = Environment.GetCommandLineArgs();
StringBuilder sb = new StringBuilder();
for (int i = 1; i < aArgs.Length; i++) { if (i > 1) sb.Append(" "); sb.Append(aArgs[i]); }
return sb.ToString();
} // GetCommandLine method

public static string[] OldGetFiles(string sDir, string sFilter, bool bSubDirs) {
string sFiles;
string[] a, aDirs, aFiles;
StringBuilder sb = new StringBuilder();

aDirs = Directory.GetDirectories(sDir);
if (bSubDirs) {
foreach (string s in aDirs) {
a = OldGetFiles(s, sFilter, bSubDirs);
sFiles = String.Join("\n", a);
if (sFiles.Length > 0) sb.Append(sFiles + "\n");
}
}

aFiles = Directory.GetFiles(sDir, sFilter);
sFiles = String.Join("\n", aFiles);
if (sFiles.Length > 0) sb.Append(sFiles + "\n");

sFiles = sb.ToString().TrimEnd();
if (sFiles.Length > 0) aFiles = sFiles.Split('\n');
else aFiles = new string[] {};
return aFiles;
} // OldGetFiles method

public static string[] FindInFiles(string sText, string sDir, string[] aFilters, bool bSubdirs) {
string[] aFiles = GetFiles(sDir, aFilters, bSubdirs);
List<string> list = new List<string>();
foreach (string sFile in aFiles) {
try { if (File.ReadAllText(sFile).IndexOf(sText, StringComparison.OrdinalIgnoreCase) >= 0) list.Add(sFile); }
catch { }
}
return list.ToArray();
} // FindInFiles method

public static string[] GetFiles(string sDir, string[] aFilters, bool bSubdirs) {
SearchOption searchOption = bSubdirs ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
List<string> list = new List<string>();
if (aFilters == null || aFilters.Length == 0) list.AddRange(Directory.GetFiles(sDir, "*", searchOption));
else foreach (string sFilter in aFilters) list.AddRange(Directory.GetFiles(sDir, sFilter, searchOption));
return list.ToArray();
} // GetFiles method

public static string[] GetPathsWithExtensions(string[] aFiles, string sExtensions) {
string sResult = "." + sExtensions.Trim().Replace(" ", " .");
sResult = sResult.Replace("..", ".");
string [] aResults = sResult.Split(' ');
List<string> list = new List<string>(aFiles);
for (int i = list.Count -1; i >=0; i--) {
string sFile = list[i];
string sExtension = Path.GetExtension(sFile).ToLower();
if (sExtension.Length == 0) sExtension = ".";
if (Array.IndexOf(aResults, sExtension) == -1) list.RemoveAt(i);
}
return list.ToArray();
} // GetPathsWithExtensions method

public static string GetExtensions(string sDir) {
return GetExtensions(Directory.GetFiles(sDir));
} // GetExtensions method

public static string GetExtensions(string[] aFiles) {
//string[] aFilters = new string[] {"*.*"};
//bool bSubdirs = false;
//string[] aFiles = GetFiles(sDir, aFilters, bSubdirs);
List<string> list = new List<string>(aFiles.Length);
for (int i = 0; i < aFiles.Length; i++) {
string s = aFiles[i];
s = Path.GetExtension(s);
//if (s.Length == 0) continue;
s = s.TrimStart('.');
s = s.ToLower();
if (s.Length == 0) s = ".";
if (!list.Contains(s)) list.Add(s);
}

list.Sort();
string[] aExtensions = list.ToArray();
return String.Join(" ", aExtensions);
} // GetExtensions method

public static bool PathExists(string sPath) {
return (Directory.Exists(sPath) || File.Exists(sPath));
} // PathExists method

public static void DeletePath(string sPath, bool bRecycle) {
FileAttributes attr = File.GetAttributes(sPath);
FileAttributes flag = FileAttributes.ReadOnly;
File.SetAttributes(sPath, (attr | flag) ^ flag);
if (Directory.Exists(sPath)) DeleteDirectory(sPath, bRecycle);
else if (File.Exists(sPath)) DeleteFile(sPath, bRecycle);
} // DeletePath method

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
struct ShFileOpStruct {
public IntPtr hwnd;
public uint wFunc;
[MarshalAs(UnmanagedType.LPWStr)] public string pFrom;
[MarshalAs(UnmanagedType.LPWStr)] public string pTo;
public ushort fFlags;
public int fAnyOperationsAborted;
public IntPtr hNameMappings;
[MarshalAs(UnmanagedType.LPWStr)] public string lpszProgressTitle;
} // ShFileOpStruct struct

[DllImport("shell32.dll", CharSet = CharSet.Unicode)]
static extern int SHFileOperation(ref ShFileOpStruct lpFileOp);

// recycle: send a file or folder to the Recycle Bin through the shell, the
// pure .NET Framework equivalent of the former VB SendToRecycleBin option.
static void recycle(string sPath) {
ShFileOpStruct op = new ShFileOpStruct();
op.wFunc = 3; // FO_DELETE
op.pFrom = sPath + "\0\0"; // list is double-null terminated
op.fFlags = (ushort) (0x0040 | 0x0010 | 0x0004 | 0x0400); // ALLOWUNDO|NOCONFIRMATION|SILENT|NOERRORUI
SHFileOperation(ref op);
} // recycle method

// copyDirectory: recursive directory copy (System.IO has no built-in one).
static void copyDirectory(string sSource, string sTarget) {
Directory.CreateDirectory(sTarget);
foreach (string sFile in Directory.GetFiles(sSource))
File.Copy(sFile, Path.Combine(sTarget, Path.GetFileName(sFile)), true);
foreach (string sDir in Directory.GetDirectories(sSource))
copyDirectory(sDir, Path.Combine(sTarget, Path.GetFileName(sDir)));
} // copyDirectory method

public static void DeleteDirectory(string sPath, bool bRecycle) {
if (!Directory.Exists(sPath)) return;
if (bRecycle) recycle(sPath);
else Directory.Delete(sPath, true);
} // DeleteDirectory method

public static void DeleteFile(string sPath, bool bRecycle) {
if (!File.Exists(sPath)) return;
if (bRecycle) recycle(sPath);
else File.Delete(sPath);
} // DeleteFile method

public static void CopyDirectory(string sSource, string sTarget, bool bRecycle) {
if (Directory.Exists(sTarget)) DeleteDirectory(sTarget, bRecycle);
else if (File.Exists(sTarget)) DeleteFile(sTarget, bRecycle);
copyDirectory(sSource, sTarget);
} // CopyDirectory method

public static void MoveDirectory(string sSource, string sTarget, bool bRecycle) {
if (Directory.Exists(sTarget)) DeleteDirectory(sTarget, bRecycle);
else if (File.Exists(sTarget)) DeleteFile(sTarget, bRecycle);
Directory.Move(sSource, sTarget);
} // MoveDirectory method

public static void CopyFile(string sSource, string sTarget, bool bRecycle) {
if (Directory.Exists(sTarget)) DeleteDirectory(sTarget, bRecycle);
else if (File.Exists(sTarget)) DeleteFile(sTarget, bRecycle);
File.Copy(sSource, sTarget, true);
} // CopyFile method

public static void MoveFile(string sSource, string sTarget, bool bRecycle) {
if (Directory.Exists(sTarget)) DeleteDirectory(sTarget, bRecycle);
else if (File.Exists(sTarget)) DeleteFile(sTarget, bRecycle);
File.Move(sSource, sTarget);
} // MoveFile method

public static void SendKeys(string sKeys) {
System.Windows.Forms.SendKeys.SendWait(sKeys);
} // SendKeys method

public static string ProperCase(string sText) {
return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(sText.ToLower());

/*
string[] aWords = sText.Split(' ');
for (int i = 0; i < aWords.Length; i++) {
string sWord = aWords[i];
string sInitial = sWord.Substring(0, 1).ToUpper();
string sRest = "";
if (sWord.Length > 1) sRest = sWord.Substring(1).ToLower();
sWord = sInitial + sRest;
aWords[i] = sWord;
}

string sReturn = String.Join(" ", aWords);
return sReturn;
*/
} // ProperCase method

public static string SwapCase(string sText) {
string sReturn = "";
StringBuilder sb = new StringBuilder(sText.Length);
for (int i = 0; i < sText.Length; i++) {
string s = sText.Substring(i, 1);
string sLower = s.ToLower();
string sUpper = s.ToUpper();
if (sLower == sUpper) sb.Append(s);
else if (s == sLower) sb.Append(sUpper);
else if (s == sUpper) sb.Append(sLower);
}

sReturn = sb.ToString();
return sReturn;
} // SwapCase method

public static int Month2Num(string sMonth) {
sMonth = Util.ProperCase(sMonth.Trim());
string[] aMonths = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"};
int iReturn = -1;
for (int i = 0; i < aMonths.Length; i++) {
string s = aMonths[i];
if (!s.StartsWith(sMonth)) continue;
iReturn = i + 1;
break;
}
return iReturn;
} // Month2Num method

public static int Day2Num(string sDay) {
sDay = Util.ProperCase(sDay.Trim());
string[] aDays = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"};
int iReturn = -1;
for (int i = 0; i < aDays.Length; i++) {
string s = aDays[i];
if (!s.StartsWith(sDay)) continue;
iReturn = i;
break;
}
return iReturn;
} // Day2Num method

public static string Type2String(object o) {
return TypeDescriptor.GetConverter(o.GetType()).ConvertToString(o);
} // Type2String method

public static object String2Type(string s, object o) {
return TypeDescriptor.GetConverter(o.GetType()).ConvertFromString(s);
} // String2Type method

public static void TerminateProcess(string sName) {
bool bLoop = true;
while (bLoop) {
Process[] processes = Process.GetProcessesByName(sName);
if (processes.Length == 0) break;

Process process = processes[0];
int iPid = process.Id;
process.CloseMainWindow();
System.Threading.Thread.Sleep(500);
try {
process = Process.GetProcessById(iPid);
process.Kill();
}
catch {
break;
}
}
} // TerminateProcess method

public static string GetLfn(string sPath) {
object oShell = COM.CreateObject("WScript.Shell");
object oShortcut = COM.CallMethod(oShell, "CreateShortcut", "temp.lnk");
COM.SetProperty(oShortcut, "TargetPath", sPath);
string sReturn = (string) COM.GetProperty(oShortcut, "TargetPath");
//COM.Release(ref oShortcut);
//COM.Release(ref oShell);
return sReturn;
} // GetLfn method

// mdPipeline: the Markdig pipeline used by the Markdown commands, built
// once on first use. UseAdvancedExtensions activates pipe and grid
// tables, footnotes, definition and task lists, auto identifiers, and
// autolinks, giving GitHub-flavored coverage without an external tool.
static Markdig.MarkdownPipeline mdPipeline = null;

static Markdig.MarkdownPipeline GetMarkdownPipeline() {
if (mdPipeline == null) mdPipeline = Markdig.MarkdownExtensions.UseAdvancedExtensions(new Markdig.MarkdownPipelineBuilder()).Build();
return mdPipeline;
} // GetMarkdownPipeline method

public static string Markdown2Html(string sSource, string sTitle) {
// Convert Markdown source to a complete (standalone) HTML document using the
// bundled Markdig library, returning the HTML. Returns "" if the conversion
// could not be done, after showing a message. Used by the HTML Format
// command, which assumes the current document is Markdown. Markdig replaced
// the former Pandoc-based conversion: it runs in process (no temp files, no
// child process, no Convert tools required) and is CommonMark compliant.
try {
App.Frame.AddMessage("Converting Markdown to HTML");
string sBody = Markdig.Markdown.ToHtml(sSource, GetMarkdownPipeline());
StringBuilder sb = new StringBuilder();
sb.Append("<!DOCTYPE html>\r\n");
sb.Append("<html>\r\n<head>\r\n");
sb.Append("<meta charset=\"utf-8\" />\r\n");
sb.Append("<title>" + String2Html(sTitle) + "</title>\r\n");
sb.Append("</head>\r\n<body>\r\n");
sb.Append(sBody);
// Mermaid diagrams: the Advanced extension set turns a fenced mermaid
// block into a div of class mermaid. When one is present, include the
// mermaid script so a real browser renders the diagram. The embedded
// preview's WebBrowser control uses the old Internet Explorer engine,
// which current mermaid does not support, so there the diagram appears
// as its readable source text -- itself the accessible view -- while
// Preview Markdown in Web Browser and any saved .htm render it drawn.
if (sBody.Contains("class=\"mermaid\"")) {
sb.Append("<script src=\"https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js\"></script>\r\n");
sb.Append("<script>if (window.mermaid) mermaid.initialize({startOnLoad:true});</script>\r\n");
}
sb.Append("</body>\r\n</html>\r\n");
return sb.ToString();
}
catch (Exception ex) {
Dialog.Show("HTML Format", "Could not convert Markdown to HTML.\n" + ex.Message);
return "";
}
} // Markdown2Html method

public static string Markdown2Text(string sSource) {
// Render Markdown source as plain text with markup stripped, via the
// Markdig library. Returns "" on failure after showing a message.
try {
return Markdig.Markdown.ToPlainText(sSource, GetMarkdownPipeline());
}
catch (Exception ex) {
Dialog.Show("Markdown to Plain Text", "Could not convert the document.\n" + ex.Message);
return "";
}
} // Markdown2Text method

// rmConverter: the ReverseMarkdown converter used by the HTML
// commands, built once on first use. GithubFlavored produces pipe
// tables and fenced code blocks that Markdig round-trips well;
// unknown tags are bypassed (contents kept, tags dropped) and HTML
// comments are removed.
static ReverseMarkdown.Converter rmConverter = null;

static ReverseMarkdown.Converter GetHtmlConverter() {
if (rmConverter == null) {
ReverseMarkdown.Config config = new ReverseMarkdown.Config();
config.GithubFlavored = true;
config.RemoveComments = true;
config.UnknownTags = ReverseMarkdown.Config.UnknownTagsOption.Bypass;
rmConverter = new ReverseMarkdown.Converter(config);
}
return rmConverter;
} // GetHtmlConverter method

// Read a file a converter just wrote. Byte-order marks are obeyed, and
// unmarked content whose bytes alternate with zeros is UTF-16 whose
// order those zeros reveal -- a certainty worth applying before any
// heuristic. Everything else goes to File2String, whose Ude-based
// detection now refuses to call zero-free content UTF-16 (see
// DetectEncodingNoBom).
// True when decoded text looks like single-byte content misread as a
// wide encoding: a large share of characters in the far-eastern blocks,
// which is what pairing Latin letters two at a time produces.
public static bool LooksLikeMisreadUtf16(string sText) {
int iSample = Math.Min(sText.Length, 2000);
if (iSample < 20) return false;
int iCjk = 0;
for (int i = 0; i < iSample; i++) {
char c = sText[i];
if (c >= '\u3000' && c <= '\u9FFF') iCjk++;
}
return iCjk * 2 > iSample;
} // LooksLikeMisreadUtf16 method

// Read a file as UTF-8, falling back to the system single-byte encoding
// when the bytes are not valid UTF-8. No detection, no guessing.
public static string PlainFile2String(string sFile) {
byte[] aBytes = File.ReadAllBytes(sFile);
try { return new UTF8Encoding(false, true).GetString(aBytes); }
catch (Exception) { return Encoding.Default.GetString(aBytes); }
} // PlainFile2String method

// The column marked by a caret line in tool output. Python answers a
// syntax error with an echo of the source line followed by a line of
// carets under the offending part; the column is the caret's position
// less the indentation the echo added. Returns 0 when there is no such
// marker line.
public static int CaretMarkerColumn(string sOutput) {
return CaretMarkerColumn(sOutput, "");
} // CaretMarkerColumn method

// The same, given the document's own text for that line. Tools differ in
// how they echo: Python adds four spaces of its own, Node repeats the
// line verbatim with its real indentation. Comparing the echo with the
// document line resolves the difference, so the column is right either
// way.
public static int CaretMarkerColumn(string sOutput, string sDocLine) {
string[] aLines = sOutput.Replace("\r\n", "\n").Split('\n');
for (int i = 1; i < aLines.Length; i++) {
string sTrim = aLines[i].Trim();
if (sTrim.Length == 0) continue;
bool bMarkerOnly = true;
foreach (char c in sTrim) if (c != '^' && c != '~') { bMarkerOnly = false; break; }
if (!bMarkerOnly) continue;
string sEcho = aLines[i - 1];
int iEchoIndent = sEcho.Length - sEcho.TrimStart().Length;
int iCaret = aLines[i].IndexOf('^');
if (iCaret < 0) continue;
int iDocIndent = 0;
if (sDocLine != null && sDocLine.Trim().Length > 0 && sDocLine.Trim() == sEcho.Trim()) iDocIndent = sDocLine.Length - sDocLine.TrimStart().Length;
// A marker that underlines the WHOLE line -- Python does this for an
// indentation error, where the fault is the line's position rather than
// any character in it -- says nothing useful about a column, so the
// cursor goes to the line's first real character instead.
if (sTrim.Length >= sEcho.Trim().Length && sEcho.Trim().Length > 0) return (sDocLine != null && sDocLine.Trim().Length > 0) ? iDocIndent + 1 : iEchoIndent + 1;
int iColumn = iCaret - iEchoIndent + iDocIndent + 1;
// Python marks a MISSING character -- the colon it wanted -- by pointing
// just past the end of the line. Landing there is right, but never
// further than one position past the last character.
if (sDocLine != null && sDocLine.Length > 0 && iColumn > sDocLine.Length + 1) iColumn = sDocLine.Length + 1;
return (iColumn > 0) ? iColumn : 0;
}
return 0;
} // CaretMarkerColumn method

public static string ConvertedFile2String(string sFile) {
try {
byte[] aBytes = File.ReadAllBytes(sFile);
// Say in the log which rule decided the encoding and what the opening
// bytes were. When a conversion comes back as gibberish, that one line
// separates a bad decode from a bad conversion without guesswork.
StringBuilder sbHead = new StringBuilder();
for (int i = 0; i < Math.Min(aBytes.Length, 8); i++) sbHead.Append(aBytes[i].ToString("x2")).Append(" ");
string sHead = Path.GetFileName(sFile) + ", " + aBytes.Length + " bytes, starts " + sbHead.ToString().Trim();
if (aBytes.Length >= 2) {
if (aBytes[0] == 0xFF && aBytes[1] == 0xFE) { Log("read as UTF-16 little endian by mark: " + sHead); return Encoding.Unicode.GetString(aBytes, 2, aBytes.Length - 2); }
if (aBytes[0] == 0xFE && aBytes[1] == 0xFF) { Log("read as UTF-16 big endian by mark: " + sHead); return Encoding.BigEndianUnicode.GetString(aBytes, 2, aBytes.Length - 2); }
}
if (aBytes.Length >= 3 && aBytes[0] == 0xEF && aBytes[1] == 0xBB && aBytes[2] == 0xBF) { Log("read as UTF-8 by mark: " + sHead); return Encoding.UTF8.GetString(aBytes, 3, aBytes.Length - 3); }
int iSample = Math.Min(aBytes.Length, 4096);
int iZeroEven = 0, iZeroOdd = 0;
for (int i = 0; i < iSample; i++) {
if (aBytes[i] != 0) continue;
if (i % 2 == 0) iZeroEven++;
else iZeroOdd++;
}
int iPairs = iSample / 2;
if (iPairs >= 8) {
if (iZeroOdd > iPairs / 2 && iZeroEven < iPairs / 8) { Log("read as UTF-16 little endian by zero pattern: " + sHead); return Encoding.Unicode.GetString(aBytes); }
if (iZeroEven > iPairs / 2 && iZeroOdd < iPairs / 8) { Log("read as UTF-16 big endian by zero pattern: " + sHead); return Encoding.BigEndianUnicode.GetString(aBytes); }
}
// No byte-order mark and no alternating zeros: single-byte text or
// UTF-8, never UTF-16. Decide here rather than passing the file to the
// general detector, whose heuristic produced the far-eastern gibberish
// in the first place.
try { string sUtf8 = new UTF8Encoding(false, true).GetString(aBytes); Log("read as UTF-8: " + sHead); return sUtf8; }
catch (Exception) { Log("read as system single-byte: " + sHead); return Encoding.Default.GetString(aBytes); }
}
catch (Exception) {}
return File2String(sFile);
} // ConvertedFile2String method

public static string Html2Markdown(string sHtml) {
// Convert HTML source to Markdown with the ReverseMarkdown library,
// which parses the HTML through HtmlAgilityPack. Returns "" on
// failure after showing a message.
try {
return GetHtmlConverter().Convert(sHtml);
}
catch (Exception ex) {
Dialog.Show("HTML to Markdown", "Could not convert the document.\n" + ex.Message);
return "";
}
} // Html2Markdown method

public static string Html2Text(string sHtml) {
// Render HTML source as plain text by chaining the two converters:
// HTML to Markdown (ReverseMarkdown), then Markdown to plain text
// (Markdig). Unlike a raw InnerText extraction, the chain preserves
// paragraph breaks, list structure, and reading order.
string sMarkdown = Html2Markdown(sHtml);
if (sMarkdown.Length == 0) return "";
try {
return Markdig.Markdown.ToPlainText(sMarkdown, GetMarkdownPipeline());
}
catch (Exception ex) {
Dialog.Show("HTML to Plain Text", "Could not convert the document.\n" + ex.Message);
return "";
}
} // Html2Text method

public static string String2Html(string sText) {
// Use System.Net.WebUtility.HtmlEncode (in System.dll, always loaded)
// rather than System.Web.HttpUtility, whose assembly fails to load at
// runtime in this desktop x64 process and crashed the Control+H HTML
// Format command. Same HTML entity-encoding result, no System.Web.
return System.Net.WebUtility.HtmlEncode(sText);
} // String2Html method

public static string ExpandCommandLine(string sCommand, string sSource, string sTarget ) {
// Dialog.Show(sTarget);

sCommand = sCommand.Replace("%NetDirLong%", App.NetDir);
sCommand = sCommand.Replace("%NetDir%", Win32.GetShortPath(App.NetDir));

sCommand = sCommand.Replace("%ProgDirLong%", App.ProgramDir);
sCommand = sCommand.Replace("%ProgDir%", Win32.GetShortPath(App.ProgramDir));

sCommand = sCommand.Replace("%DataDirLong%", App.DataDir);
sCommand = sCommand.Replace("%DataDir%", Win32.GetShortPath(App.DataDir));

if (sSource.Length > 0) {
sCommand = sCommand.Replace("%SourceLong%", sSource);
sCommand = sCommand.Replace("%Source%", Win32.GetShortPath(sSource));

sCommand = sCommand.Replace("%SourceDirLong%", Path.GetDirectoryName(sSource));
sCommand = sCommand.Replace("%SourceDir%", Win32.GetShortPath(Path.GetDirectoryName(sSource)));

sCommand = sCommand.Replace("%SourceNameLong%", Path.GetFileName(sSource));
sCommand = sCommand.Replace("%SourceName%", Path.GetFileName(Win32.GetShortPath(sSource)));

sCommand = sCommand.Replace("%SourceRootLong%", Path.GetFileNameWithoutExtension(sSource));
sCommand = sCommand.Replace("%SourceRoot%", Path.GetFileNameWithoutExtension(Win32.GetShortPath(sSource)));

sCommand = sCommand.Replace("%SourceExtLong%", Path.GetExtension(sSource));
sCommand = sCommand.Replace("%SourceExt%", Path.GetExtension(Win32.GetShortPath(sSource)));
}

if (sTarget.Length > 0) {
sCommand = sCommand.Replace("%TargetLong%", sTarget);
// Dialog.Show(sCommand, "");
sCommand = sCommand.Replace("%Target%", Win32.GetShortPath(sTarget));
// Dialog.Show(sCommand, "");

sCommand = sCommand.Replace("%TargetDirLong%", Path.GetDirectoryName(sTarget));
sCommand = sCommand.Replace("%TargetDir%", Win32.GetShortPath(Path.GetDirectoryName(sTarget)));

sCommand = sCommand.Replace("%TargetNameLong%", Path.GetFileName(sTarget));
sCommand = sCommand.Replace("%TargetName%", Path.GetFileName(Win32.GetShortPath(sTarget)));

sCommand = sCommand.Replace("%TargetRootLong%", Path.GetFileNameWithoutExtension(sTarget));
sCommand = sCommand.Replace("%TargetRoot%", Path.GetFileNameWithoutExtension(Win32.GetShortPath(sTarget)));

sCommand = sCommand.Replace("%TargetExtLong%", Path.GetExtension(sTarget));
sCommand = sCommand.Replace("%TargetExt%", Path.GetExtension(Win32.GetShortPath(sTarget)));
}

sCommand = sCommand.Replace("%TempFile%", App.TempFile);
sCommand = Environment.ExpandEnvironmentVariables(sCommand);
return sCommand.Trim();
} // ExpandCommandLine method

public static string GetProgramOutput(string sExe, string sParams) {
Process process = new Process();
ProcessStartInfo startInfo = new ProcessStartInfo(sExe, sParams);
startInfo.UseShellExecute = false;
//startInfo.RedirectStandardInput = true;
startInfo.RedirectStandardOutput = true;
startInfo.RedirectStandardError = true;
startInfo.WorkingDirectory = Path.GetDirectoryName(sExe);
startInfo.ErrorDialog = true;
startInfo.CreateNoWindow = true;
startInfo.WindowStyle = ProcessWindowStyle.Hidden;
process.StartInfo = startInfo;
process.Start();
StreamReader stream = process.StandardOutput;
process.WaitForExit();
string sText = stream.ReadToEnd();
stream.Close();
process.Close();
return sText;
} // GetProgramOutput method

public static bool ConvertString2FileFormat(string sText, string sTarget, string sTargetFormat) {
string sSourceFormat = "";
return ConvertString2FileFormat(sText, sSourceFormat, sTarget, sTargetFormat);
} // ConvertString2FileFormat method

public static bool ConvertString2FileFormat(string sText, string sSourceFormat, string sTarget, string sTargetFormat) {
/*
string sSource = Path.GetTempFileName();
if (sSourceFormat.Length > 0) {
string s = Path.ChangeExtension(sSource, sSourceFormat);
if (File.Exists(s)) File.Delete(s);
File.Move(sSource, s);
sSource = s;
}

//sSource = @"C:\edsharp\edsharp.htm";
App.TempFiles.Add(sSource);
//Util.String2File(sText, sSource);
//Util.StringAppend2File(sText, sSource);
*/

string sDir = Path.Combine(App.DataDir, "Temp");
if (Directory.Exists(sDir)) Util.DeleteDirectory(sDir, false);
Directory.CreateDirectory(sDir);
string sSource = Path.Combine(sDir, "Source.tmp");
if (sSourceFormat.Length > 0) sSource = Path.ChangeExtension(sSource, sSourceFormat);

Util.String2FileA(sText, sSource);
sSource = Win32.GetShortPath(sSource);
// Dialog.Show(sSource, Util.File2String(sSource));
string sCommand = Ini.ReadValue(App.IniFile, "Export", sTargetFormat, "");
//Dialog.Show(sCommand);
if (sCommand.Length > 0) {
sCommand = Util.ExpandCommandLine(sCommand, sSource, sTarget);
// Dialog.Show("show", sCommand);
App.Frame.AddMessage("Converting");
if (File.Exists(sTarget)) File.Delete(sTarget);
Util.RunHideWait(sCommand);
if (!File.Exists(sTarget)) {
sCommand = "cmd.exe /c " + sCommand;
Util.RunHideWait(sCommand);
}
sText = "";
if (File.Exists(sTarget)) sText = Util.ConvertedFile2String(sTarget);
if (sText.Length == 0) {
if (File.Exists(sTarget)) File.Delete(sTarget);
Dialog.Show("Error", "The conversion produced no output.\nCommand line:\n" + sCommand + "\n\nThe run log records each step and its exit code:\n" + App.LogFile);
}
}
else {
COM.WordSource2TargetFormat(sSource, sTarget, sTargetFormat);
}
//Dialog.Show("Error", "Command line:\n" + sCommand);
App.Frame.Activate();
return File.Exists(sTarget);
} // ConvertString2FileFormat method

public static string Literalize(string sText) {
bool bCheckPrefix = false;
return Literalize(sText, bCheckPrefix);
} // Literalize method

public static string Literalize(string sText, bool bCheckPrefix) {
if (bCheckPrefix) {
if (sText.StartsWith("@")) return sText.Substring(1);
else if (sText.StartsWith(@"\@")) sText = sText.Substring(1);
}
string sReturn = null;
try {
sReturn = Script.run("\"" + sText + "\"");
}
catch {}

//string sReturn = JS.Eval("\"" + sText + "\"").ToString();
//if (sReturn.Length == 0) sReturn = sText;
if (sReturn == null) sReturn = sText;
return sReturn;
} // Literalize method

public static string Reverse(string sText) {
/*
string[] a = sText.Split();
Array.Reverse(a);
sText = String.Join("", a);
*/
int iLength = sText.Length;
StringBuilder sb = new StringBuilder(iLength);
for (int i = iLength - 1; i >= 0; i--) sb.Append(sText.Substring(i, 1));
sText = sb.ToString();
return sText;
} // Reverse method

public static int Absolute(int i) {
if (i < 0) i = -1 * i;
return i;
} // Absolute method

public static bool IsNumeric(string sText) {
double dValue; return double.TryParse(sText, out dValue);
} // IsNumeric method

public static bool IsDate(string sText) {
DateTime dtValue; return DateTime.TryParse(sText, out dtValue);
} // IsDate method

public static bool IsNothing(string sText) {
return sText == null;
} // IsNothing method

public static string Left(string sText, int iChars) {
if (sText == null) return sText;
if (iChars < 0) iChars = 0;
return (iChars >= sText.Length) ? sText : sText.Substring(0, iChars);
} // Left method

public static string Right(string sText, int iChars) {
if (sText == null) return sText;
if (iChars < 0) iChars = 0;
return (iChars >= sText.Length) ? sText : sText.Substring(sText.Length - iChars);
} // Right method

public static Font SetBold(Font font, bool bState) {
return new Font(font, bState ? font.Style | FontStyle.Bold : font.Style & ~FontStyle.Bold);
} // SetBold method

public static Font SetItalic(Font font, bool bState) {
return new Font(font, bState ? font.Style | FontStyle.Italic : font.Style & ~FontStyle.Italic);
} // SetItalic method

public static Font SetUnderline(Font font, bool bState) {
return new Font(font, bState ? font.Style | FontStyle.Underline : font.Style & ~FontStyle.Underline);
} // SetUnderline method

public static string GetFileFromUri(string sUri) {
string sFile;
Uri oUri = new Uri(sUri);
//if (oUri.IsFile) {
sFile = oUri.LocalPath;
try {
sFile = Path.GetFileName(sFile);
}
catch {
sFile = "";
}
//else {
if (sFile.Length == 0) {
sFile = oUri.PathAndQuery;
sFile = Uri.UnescapeDataString(sFile);
StringBuilder sb = new StringBuilder();
for (int i = 0; i < sFile.Length; i++) {
if (Char.IsLetterOrDigit(sFile, i)) sb.Append(sFile.Substring(i, 1));
else sb.Append("_");
}
sFile = sb.ToString();
sFile = Util.RegExpReplaceCase(sFile, @"_+", "_");
sFile = sFile.Trim(new Char[] {'_', ' '});
if (sFile.Length == 0) sFile = "page";
if (!sFile.ToLower().EndsWith(".htm") && !sFile.ToLower().EndsWith(".html")) sFile += ".htm";
}
if (Path.GetExtension(sFile).Length == 0) sFile += ".htm";
return sFile;
} // GetFileFromUri method

public static string GetUniqueName(string sSource) {
if (!Directory.Exists(sSource) && !File.Exists(sSource)) return sSource;
string sTarget = "";
string sDir = Path.GetDirectoryName(sSource);
string sRoot = Path.GetFileNameWithoutExtension(sSource);
sRoot = Regex.Replace(sRoot, @"_\d\d$", "");
//Regex rx = new Regex(@"_\d\d$");
//sRoot = rx.Replace(sRoot, "");
string sExt = Path.GetExtension(sSource);
//for (int i = 1; i < 100; i++) {
for (int i = 1; i < 10000; i++) {
//string sNewName = sRoot + "_" + i.ToString().PadLeft(2, '0') + sExt;
string sNewName = sRoot + "_" + i.ToString().PadLeft(4, '0') + sExt;
sTarget = Path.Combine(sDir, sNewName);
if (!Directory.Exists(sTarget) && !File.Exists(sTarget)) break;
}
//if (Directory.Exists(sTarget) || File.Exists(sTarget)) sTarget = "";
return sTarget;
} // GetUniqueName method

public static void Swap(ref int i1, ref int i2) {
int i  = i1;
i1 = i2;
i2 = i;
} // Swap method

public static char Code2Char(int iCode) {
return (char) iCode;
} // Code2Char method

public static string Code2String(int iCode) {
return Code2Char(iCode).ToString();
} // Code2String method

} // Util class

public class HomerList : List<string> {

public char Delimiter = '|';
public bool CaseSensitive = false;

public int Max {
get {
return this.Count - 1;
}
}

public string Segments {
get {
string[] aSegments = this.ToArray();
string sSegments = String.Join(this.Delimiter.ToString(), aSegments);
return sSegments;
}
set {
string[] aSegments = value.Split(this.Delimiter);
this.Clear();
if (value.Length > 0) this.AddRange(aSegments);
}
} // Segments property

public HomerList() {
//new HomerList(this.Segments, this.Delimiter, this.CaseSensitive);
} // HomerList constructor

public HomerList(string sSegments) {
//new HomerList(sSegments, this.Delimiter, this.CaseSensitive);
this.Segments = sSegments;
//new HomerList();
} // HomerList constructor

public HomerList(string sSegments, char cDelimiter) {
this.Delimiter = cDelimiter;
this.Segments = sSegments;
} // HomerList constructor

public HomerList(string sSegments, char cDelimiter, bool bCaseSensitive) {
this.Delimiter = cDelimiter;
this.Segments = sSegments;
this.CaseSensitive = bCaseSensitive;
} // HomerList constructor

public HomerList(string[] aItems) {
this.AddRange(aItems);
} // HomerList constructor

public new int IndexOf(string sItem) {
if (this.CaseSensitive) return base.IndexOf(sItem);
else {
int iIndex = -1;
string sValue = sItem.ToLower();
for (int i = 0; i < this.Count; i++) {
if (this[i].ToLower() == sValue) {
iIndex = i;
break;
}
}
return iIndex;
}
} // IndexOf method

public new bool Contains(string sItem) {
return this.IndexOf(sItem) >= 0;
} // Contains method

public new void Sort() {
if (this.CaseSensitive) base.Sort();
else {
string[] a = this.ToArray();
Array.Sort(a, new CaseInsensitiveComparer());
this.Clear();
this.AddRange(a);
}
} // Sort method

public string GetSegments(char cDelimiter) {
this.Delimiter = cDelimiter;
return this.Segments;
} // GetSegments method

public void KeepUnique() {
for (int i = this.Count - 1; i >=0; i--) {
string s = this[i];
if (this.IndexOf(s) < i) this.RemoveAt(i);
}
} // KeepUnique method

public void RemoveLike(string sMatch) {
RegexOptions options = RegexOptions.Multiline;
if (!this.CaseSensitive) options |= RegexOptions.IgnoreCase;
Regex rx = new Regex(sMatch, options);

for (int i = this.Count - 1; i >= 0; i--) {
if (rx.IsMatch(this[i])) this.RemoveAt(i);
}
} // RemoveLike method

public void KeepLike(string sMatch) {
HomerList hl = this.FindLike(sMatch);
this.Clear();
this.AddRange(hl);
} // KeepLike method

public HomerList FindLike(string sMatch) {
RegexOptions options = RegexOptions.Multiline;
if (!this.CaseSensitive) options |= RegexOptions.IgnoreCase;
Regex rx = new Regex(sMatch, options);

HomerList hl = new HomerList();
foreach (string s in this) {
if (rx.IsMatch(s)) hl.Add(s);
}
return hl;
} // FindLike method

public void ReplaceLike(string sMatch, string sReplace) {
RegexOptions options = RegexOptions.Multiline;
if (!this.CaseSensitive) options |= RegexOptions.IgnoreCase;
Regex rx = new Regex(sMatch, options);

for (int i = 0; i < this.Count; i++)  this[i] = rx.Replace(this[i], sReplace);
} // ReplaceLike method

public void Push(string sItem) {
this.Insert(0, sItem);
} // Push method

public string Pop() {
int iUpper = this.Count - 1;
string sItem = this[iUpper];
this.RemoveAt(iUpper);
return sItem;
} // Pop method

public string Shift() {
int iLower = 0;
string sItem = this[iLower];
this.RemoveAt(iLower);
return sItem;
} // Shift method

public new void Remove(string sItem) {
bool bLoop = true;
while (bLoop) {
int iIndex = this.IndexOf(sItem);
if (iIndex == -1) break;
this.RemoveAt(iIndex);
}
} // Remove method

public void RemoveRange(HomerList hl) {
foreach (string sItem in hl) this.Remove(sItem);
} // RemoveRange method

public void RemoveRange(string[] aItems) {
HomerList hl = new HomerList(aItems);
this.RemoveRange(hl);
} // RemoveRange method

public void AddUnique(string sItem) {
if (!this.Contains(sItem)) this.Add(sItem);
} // AddUnique method

public void PushUnique(string sItem) {
if (!this.Contains(sItem)) this.Push(sItem);
} // PushUnique method

public void RemoveRange(string sSegments) {
Char cDelimiter = '|';
HomerList hl = new HomerList(sSegments, cDelimiter);
this.RemoveRange(hl);
} // RemoveRange method

public void RemoveRange(string sSegments, Char cDelimiter) {
HomerList hl = new HomerList(sSegments, cDelimiter);
this.RemoveRange(hl);
} // RemoveRange method

public void saveAddRange(HomerList hl) {
this.AddRange(hl);
} // AddRange method

public void OldAddRange(string[] aItems) {
HomerList hl = new HomerList(aItems);
this.AddRange(hl);
} // AddRange method

public void AddRange(string sSegments) {
Char cDelimiter = this.Delimiter;
HomerList hl = new HomerList(sSegments, cDelimiter);
this.AddRange(hl);
} // AddRange method

public void AddRange(string sSegments, Char cDelimiter) {
HomerList hl = new HomerList(sSegments, cDelimiter);
this.AddRange(hl);
} // AddRange method

public void AddUniqueRange(HomerList hl) {
foreach (string s in hl) if (!this.Contains(s)) this.Add(s);
} // AddUniqueRange method

public void AddUniqueRange(string[] aItems) {
HomerList hl = new HomerList(aItems);
this.AddUniqueRange(hl);
} // AddUniqueRange method

public void AddUniqueRange(string sSegments) {
Char cDelimiter = this.Delimiter;
HomerList hl = new HomerList(sSegments, cDelimiter);
this.AddUniqueRange(hl);
} // AddUniqueRange method

public void AddUniqueRange(string sSegments, Char cDelimiter) {
HomerList hl = new HomerList(sSegments, cDelimiter);
this.AddUniqueRange(hl);
} // AddUniqueRange method

public HomerList FindRange(HomerList hl) {
HomerList hlReturn = new HomerList();
foreach (string sItem in hl) if (this.Contains(sItem)) hlReturn.Add(sItem);
return hlReturn;
} // FindRange method

public void FindRange(string[] aItems) {
HomerList hl = new HomerList(aItems);
this.FindRange(hl);
} // FindRange method

public void FindRange(string sSegments) {
Char cDelimiter = '|';
HomerList hl = new HomerList(sSegments, cDelimiter);
this.FindRange(hl);
} // FindRange method

public void FindRange(string sSegments, Char cDelimiter) {
HomerList hl = new HomerList(sSegments, cDelimiter);
this.FindRange(hl);
} // FindRange method

public HomerList Clone() {
string[] aItems = this.ToArray();
HomerList hl = new HomerList(aItems);
return hl;
} // Clone method

public string MinValue() {
if (this.Count == 0) return "";

HomerList hl = this.Clone();
hl.Sort();
return hl[0];
} // MinValue method

public string MaxValue() {
if (this.Count == 0) return "";

HomerList hl = this.Clone();
hl.Sort();
return hl[hl.Count - 1];
} // MaxValue method

public int MinLength() {
int iLength = 2000000000;
foreach (string sItem in this) if (sItem.Length < iLength) iLength = sItem.Length;
if (iLength == 2000000000) iLength = 0;
return iLength;
} // MinLength method

public int MaxLength() {
int iLength = 0;
foreach (string sItem in this) if (sItem.Length > iLength) iLength = sItem.Length;
return iLength;
} // MaxLength method

public void SortLength() {
this.Sort(delegate(string s1, string s2) {
return s1.Length.CompareTo(s2.Length);
} );
} // SortLength method

public void ToLower() {
for (int i = 0; i < this.Count; i++) this[i] = this[i].ToLower();
} // ToLower method

public void ToUpper() {
for (int i = 0; i < this.Count; i++) this[i] = this[i].ToUpper();
} // ToUpper method

public void TrimStart() {
for (int i = 0; i < this.Count; i++) this[i] = this[i].TrimStart();
} // TrimStart method

public void TrimEnd() {
for (int i = 0; i < this.Count; i++) this[i] = this[i].TrimEnd();
} // TrimEnd method

public void TrimStart(char[] a) {
for (int i = 0; i < this.Count; i++) this[i] = this[i].TrimStart(a);
} // TrimStart method

public void TrimEnd(char[] a) {
for (int i = 0; i < this.Count; i++) this[i] = this[i].TrimEnd(a);
} // TrimEnd method

public void PadLeft(int iLength, char c) {
for (int i = 0; i < this.Count; i++) this[i] = this[i].PadLeft(iLength, c);
} // PadLeft method

public void PadRight(int iLength, char c) {
for (int i = 0; i < this.Count; i++) this[i] = this[i].PadRight(iLength, c);
} // PadRight method

public void PushRange(HomerList hl) {
this.InsertRange(0, hl);
} // PushRange method

public void PushRange(string sSegments) {
Char cDelimiter = this.Delimiter;
HomerList hl = new HomerList(sSegments, cDelimiter);
this.PushRange(hl);
} // PushRange method

public void PushRange(string sSegments, Char cDelimiter) {
HomerList hl = new HomerList(sSegments, cDelimiter);
this.PushRange(hl);
} // PushRange method

public void PushUniqueRange(HomerList hl) {
foreach (string s in hl) if (!this.Contains(s)) this.Push(s);
} // PushUniqueRange method

public void PushUniqueRange(string[] aItems) {
HomerList hl = new HomerList(aItems);
this.PushUniqueRange(hl);
} // PushUniqueRange method

public void PushUniqueRange(string sSegments) {
Char cDelimiter = this.Delimiter;
HomerList hl = new HomerList(sSegments, cDelimiter);
this.PushUniqueRange(hl);
} // PushUniqueRange method

public void PushUniqueRange(string sSegments, Char cDelimiter) {
HomerList hl = new HomerList(sSegments, cDelimiter);
this.PushUniqueRange(hl);
} // PushUniqueRange method

} // HomerList class

public class Segment {
public static bool CaseSensitive = false;
public static char Delimiter = '|';

public int Count(string sSegments) {
string[] aSegments = sSegments.Split(Segment.Delimiter);
return sSegments.Length == 0 ? 0 : aSegments.Length;
} // Count method

public static string Get(string sSegments, int iIndex) {
string[] aSegments = sSegments.Split(Segment.Delimiter);
return aSegments[iIndex];
} // Get method

public static int IndexOf(string sSegments, string sSegment) {
string[] aSegments = sSegments.Split(Segment.Delimiter);
List<string> listSegments = new List<string>(aSegments);
if (!Segment.CaseSensitive) for (int i = 0; i < listSegments.Count; i++) listSegments[i] = listSegments[i].ToLower();
return listSegments.IndexOf(sSegment);
} // IndexOf method

public static bool Contains(string sSegments, string sSegment) {
return IndexOf(sSegments, sSegment) >= 0;
} // Contains method

public static string RemoveAt(string sSegments, int iIndex) {
string[] aSegments = sSegments.Split(Segment.Delimiter);
List<string> listSegments = new List<string>(aSegments);
listSegments.RemoveAt(iIndex);
aSegments = listSegments.ToArray();
return String.Join(Segment.Delimiter.ToString(), aSegments);
} // RemoveAt method

public static string Remove(string sSegments, string sSegment) {
int iIndex = IndexOf(sSegments, sSegment);
return RemoveAt(sSegments, iIndex);
} // Remove method

public static string RemoveIfContains(string sSegments, string sSegment) {
int iIndex = IndexOf(sSegments, sSegment);
if (iIndex >= 0) sSegments = RemoveAt(sSegments, iIndex);
return sSegments;
} // RemoveIfContains method

public static string Insert(string sSegments, int iIndex, string sSegment) {
string[] aSegments = sSegments.Split(Segment.Delimiter);
List<string> listSegments = new List<string>(aSegments);
listSegments.Insert(iIndex, sSegment);
aSegments = listSegments.ToArray();
return String.Join(Segment.Delimiter.ToString(), aSegments);
} // Insert method

public static string Add(string sSegments, string sSegment) {
string[] aSegments = sSegments.Split(Segment.Delimiter);
List<string> listSegments = new List<string>(aSegments);
listSegments.Add(sSegment);
aSegments = listSegments.ToArray();
return String.Join(Segment.Delimiter.ToString(), aSegments);
} // Add method

public static string AddIfUnique(string sSegments, string sSegment) {
if (!Contains(sSegments, sSegment)) sSegments = Add(sSegments, sSegment);
return sSegments;
} // AddIfUnique method

public static string ReplaceAt(string sSegments, int iIndex, string sSegment) {
string[] aSegments = sSegments.Split(Segment.Delimiter);
List<string> listSegments = new List<string>(aSegments);
listSegments[iIndex] = sSegment;
aSegments = listSegments.ToArray();
return String.Join(Segment.Delimiter.ToString(), aSegments);
} // ReplaceAt method

public static string Replace(string sSegments, string sSegment) {
int iIndex = IndexOf(sSegments, sSegment);
return ReplaceAt(sSegments, iIndex, sSegment);
} // Replace method

public static string ReplaceIfContains(string sSegments, string sSegment) {
int iIndex = IndexOf(sSegments, sSegment);
if (iIndex >= 0) sSegments = ReplaceAt(sSegments, iIndex, sSegment);
return sSegments;
} // ReplaceIfContains method

public static string Sort(string sSegments) {
string[] aSegments = sSegments.Split(Segment.Delimiter);
List<string> listSegments = new List<string>(aSegments);
if (!Segment.CaseSensitive) for (int i = 0; i < listSegments.Count; i++) listSegments[i] = listSegments[i].ToLower();
string[] aKeys = listSegments.ToArray();
Array.Sort(aKeys, aSegments);
return String.Join(Segment.Delimiter.ToString(), aSegments);
} // Sort method

public static string Unique(string sSegments) {
string[] aSegments = sSegments.Split(Segment.Delimiter);
List<string> listSegments = new List<string>();
if (Segment.CaseSensitive) foreach (string sSegment in aSegments) if (!listSegments.Contains(sSegment)) listSegments.Add(sSegment);
else {
List<string> listLower = new List<string>();
foreach (string s in aSegments) {
if (!listLower.Contains(s.ToLower())) {
listSegments.Add(s);
listLower.Add(s.ToLower());
}
}
}

aSegments = listSegments.ToArray();
return String.Join(Segment.Delimiter.ToString(), aSegments);
} // Unique method

} // Segment class

// =====================================================================
// PreviewForm: renders HTML in an embedded WebBrowser control so the
// screen reader's virtual buffer applies -- JAWS and NVDA element
// navigation works exactly as on a web page. Used by Preview Markdown
// (Control+F9). Modeless and owned by the frame, so the editor remains
// reachable. Escape closes when the browser surface forwards the key
// (best effort with an embedded browser); Alt+F4 always closes. The
// WebBrowser control ships inside the .NET Framework, so this feature
// adds no dependency; if its legacy engine ever falls short, the
// upgrade path is swapping this one class to Microsoft's WebView2.
public class PreviewForm : Form {
WebBrowser browser;

public PreviewForm(string sTitle, string sHtml) {
this.Text = sTitle + " - Markdown Preview";
this.StartPosition = FormStartPosition.CenterScreen;
this.Width = 1000;
this.Height = 700;
this.KeyPreview = true;
this.KeyDown += delegate(object o, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) this.Close(); };
browser = new WebBrowser();
browser.Dock = DockStyle.Fill;
browser.ScriptErrorsSuppressed = true;
browser.WebBrowserShortcutsEnabled = true;
browser.PreviewKeyDown += delegate(object o, PreviewKeyDownEventArgs e) { if (e.KeyCode == Keys.Escape) this.Close(); };
browser.DocumentText = sHtml;
this.Controls.Add(browser);
this.Shown += delegate(object o, EventArgs e) { browser.Focus(); };
} // PreviewForm constructor

public static void ShowPreview(Form frmOwner, string sTitle, string sHtml) {
PreviewForm frm = new PreviewForm(sTitle, sHtml);
frm.Show(frmOwner);
} // ShowPreview method
} // PreviewForm class

} // EdSharp namespace
