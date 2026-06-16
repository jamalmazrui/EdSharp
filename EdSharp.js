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
    }
}
