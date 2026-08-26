/*
EdSharp.js -- JScript .NET scripting host for EdSharp, compiled by
jsc.exe into EdSharp.dll. EdSharp.exe loads this assembly at run time
via Assembly.LoadFrom(<full path>) and calls EdSharp.JS.runScript by
reflection, so EdSharp.cs takes no compile-time dependency on this
assembly. That matters because the exe and this dll share the base
name "EdSharp"; a compile-time reference would make Assembly.Load
ambiguous at load time. Late-bound dispatch resolves snippet member
access at run time with no compile-time type information.

This revives, in one file, the role formerly split across jsSupport.js
and Eval.js, following the DbDo.js model.

Globals visible to user .js snippets, inside the eval scope:
  frm -- the active editor window (MdiChild), passed in from C#.
  rtb -- shortcut for frm.RTB, the RichTextBox holding the document.

Camel Type: lower-camelCase functions and variables; frm and rtb are
the conventional short forms. The entry point is runScript rather than
eval, because eval is a JScript built-in we cannot shadow in our body.

Usage from C# (see the Script helper in EdSharp.cs):
  var asm = Assembly.LoadFrom(sDllPath);
  var jsType = asm.GetType("EdSharp.JS");
  var mi = jsType.GetMethod("runScript",
    new Type[] { typeof(string), typeof(object), typeof(object) });
  string sResult = (string)mi.Invoke(null,
    new object[] { sCode, frm, rtb });

The returned string is the script's last expression value via String(),
or "ERROR: " + message on any compile or runtime fault. The script does
NOT throw out to the host, so EdSharp's UI stays responsive.
*/

import System;
import System.Collections;
import System.Drawing;
import System.IO;
import System.Reflection;
import System.Text;
import System.Text.RegularExpressions;
import System.Windows.Forms;

package EdSharp
{
    public class JS
    {
        // runScript: evaluate sCode with frm and rtb in scope. The two
        // host objects are typed Object so this assembly needs no
        // reference to EdSharp.exe; JScript late-binds member access.
        public static function runScript(sCode : String, frm : Object, rtb : Object) : String
        {
            try
            {
                var oResult = eval(sCode, "unsafe");
                if (oResult == null) return "";
                return String(oResult);
            }
            catch (oError)
            {
                return "ERROR: " + oError.message;
            }
        }

        // ===== Interactive console ========================================
        //
        // A read-evaluate-print loop for writing EdSharp snippets. It is
        // the Interactive JScript program of the Homer.NET toolkit
        // (2009-2010), brought inside EdSharp: rewritten in Camel Type,
        // stripped of the screen readers and helpers that no longer
        // exist, and given what it lacked -- the editor itself. When it
        // starts, frm is the window you were editing and rtb is its text
        // box, already live, so a line typed here does to your document
        // exactly what the same line would do in a snippet.
        //
        // Everything is an expression unless it begins with one of the
        // words listed by help. Expressions print their value; statements
        // print nothing.

        static var frm : Object = null;
        static var rtb : Object = null;

        static function say(oText : Object) : void
        {
            Console.WriteLine((oText == null) ? "" : String(oText));
        }

        // The reminder shown at the start, and by help afterwards. It
        // names what is in scope before it names anything else, because
        // that is the reason this console exists.
        static function sayHelp() : void
        {
            say("Interactive JScript for EdSharp snippets.");
            say("");
            say("Already in scope:");
            say("  frm   the editor window you came from");
            say("  rtb   its text box: rtb.Text, rtb.SelectedText, rtb.Lines");
            say("");
            say("Type any JScript expression to see its value, or any");
            say("statement to run it. What works here works in a snippet.");
            say("");
            say("  rtb.SelectedText.toUpperCase()");
            say("  rtb.SelectedText = rtb.SelectedText.toUpperCase()");
            say("  frm.Text");
            say("");
            say("Commands, each a word with no quotes around its argument:");
            say("  help              this reminder");
            say("  quit              close the console and return to EdSharp");
            say("  cls               clear the screen");
            say("  load FileName     read a .js file and run it here");
            say("  save FileName     write this session's lines to a file");
            say("  members Thing     list what an object or type offers");
            say("  methods Thing     its methods only");
            say("  properties Thing  its properties only");
            say("  doc               open the EdSharp guide in your browser");
        }

        // What an object or a type can do. The old program had six
        // listing commands over reflection and COM type libraries; three
        // remain, because .NET reflection answers for both and the COM
        // ones described a world that has moved on.
        static function listMembers(oThing : Object, sKind : String) : void
        {
            try
            {
                var oType : Type = (oThing instanceof Type) ? Type(oThing) : oThing.GetType();
                say(oType.FullName);
                var aMembers : MemberInfo[] = oType.GetMembers();
                var lsNames : ArrayList = new ArrayList();
                for (var i : int = 0; i < aMembers.length; i++)
                {
                    var oMember : MemberInfo = aMembers[i];
                    var sType : String = String(oMember.MemberType);
                    if (sKind == "methods" && sType != "Method") continue;
                    if (sKind == "properties" && sType != "Property") continue;
                    var sLine : String = oMember.Name + "  (" + sType.toLowerCase() + ")";
                    if (!lsNames.Contains(sLine)) lsNames.Add(sLine);
                }
                lsNames.Sort();
                for (var j : int = 0; j < lsNames.Count; j++) say("  " + lsNames[j]);
                say(lsNames.Count + ((lsNames.Count == 1) ? " member" : " members"));
            }
            catch (oError)
            {
                say("Cannot list that: " + oError.message);
            }
        }

        // The loop itself. Returns when the person types quit, so the
        // caller can free the console window.
        public static function runConsole(oFrm : Object, oRtb : Object) : void
        {
            frm = oFrm;
            rtb = oRtb;
            var lsSession : ArrayList = new ArrayList();
            sayHelp();
            say("");

            while (true)
            {
                Console.Write("js> ");
                var sInput : String = Console.ReadLine();
                if (sInput == null) return;
                sInput = sInput.replace(/^\s+|\s+$/g, "");
                if (sInput.length == 0) continue;
                lsSession.Add(sInput);

                var sWord : String = sInput;
                var sRest : String = "";
                var iSpace : int = sInput.indexOf(" ");
                if (iSpace > 0)
                {
                    sWord = sInput.substring(0, iSpace);
                    sRest = sInput.substring(iSpace + 1).replace(/^\s+|\s+$/g, "");
                }

                try
                {
                    if (sWord == "quit") return;
                    else if (sWord == "help") sayHelp();
                    else if (sWord == "cls") Console.Clear();
                    else if (sWord == "doc")
                    {
                        var sGuide : String = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "EdSharp.htm");
                        if (File.Exists(sGuide)) System.Diagnostics.Process.Start(sGuide);
                        else say("The guide is not beside EdSharp.");
                    }
                    else if (sWord == "load")
                    {
                        if (!File.Exists(sRest)) say("No such file: " + sRest);
                        else
                        {
                            var sCode : String = File.ReadAllText(sRest);
                            var oLoaded = eval(sCode, "unsafe");
                            if (oLoaded != null) say(String(oLoaded));
                            say("Ran " + sRest);
                        }
                    }
                    else if (sWord == "save")
                    {
                        var sbLines : StringBuilder = new StringBuilder();
                        for (var k : int = 0; k < lsSession.Count - 1; k++) sbLines.AppendLine(String(lsSession[k]));
                        File.WriteAllText(sRest, sbLines.ToString());
                        say("Wrote " + (lsSession.Count - 1) + " lines to " + sRest);
                    }
                    else if (sWord == "members" || sWord == "methods" || sWord == "properties")
                    {
                        listMembers(eval(sRest, "unsafe"), sWord);
                    }
                    else
                    {
                        var oValue = eval(sInput, "unsafe");
                        if (oValue != null) say(oValue);
                    }
                }
                catch (oError)
                {
                    say("Error: " + oError.message);
                }
            }
        }
    }
}
