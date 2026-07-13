// Inix.cs (InixCodec: order-preserving .ini/.inix reader-writer) -- portable, reusable across EdSharp, DbDuo, and other C#
// projects. .inix is a superset of classic .ini (verbatim multi-line values, implicit [Global], order-preserving round-trip). Pure codec, no app dependencies.
// Namespace Homer is the shared toolkit namespace; reference it with
// `using Homer;`. To reuse elsewhere, copy this file as-is.

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
using System.Web;
using System.Windows.Forms;

namespace Homer {

public static class InixCodec
{
    // A single key/value pair, preserving order within a section.
    public class Pair
    {
        public string Key;
        public string Value;
        public Pair(string sK, string sV) { Key = sK; Value = sV; }
    }

    // A single section: name plus ordered list of pairs. Order of
    // pairs is preserved on round-trip.
    public class Section
    {
        public string Name;
        public List<Pair> Pairs = new List<Pair>();
        public Section(string sName) { Name = sName; }

        // Lookup is case-insensitive on the key (matches the .ini
        // convention). Returns null if the key is absent.
        public string get(string sKey)
        {
            foreach (Pair p in Pairs)
                if (string.Equals(p.Key, sKey, StringComparison.OrdinalIgnoreCase))
                    return p.Value;
            return null;
        }

        // getArray: the value read as an array of items -- the read side of the
        // .inix array convention.  A value that spans multiple lines yields one item
        // per line; a single-line value containing commas is split on commas; any
        // other non-blank value is a one-item array.  Surrounding whitespace and
        // blank items are dropped.  Returns an empty list when the key is absent or
        // blank.
        public List<string> getArray(string sKey)
        {
            List<string> lsItems = new List<string>();
            string sRaw = get(sKey);
            if (string.IsNullOrEmpty(sRaw)) return lsItems;
            string sNorm = sRaw.Replace("\r\n", "\n").Replace("\r", "\n");
            string[] aParts;
            if (sNorm.IndexOf('\n') >= 0)     aParts = sNorm.Split('\n');
            else if (sNorm.IndexOf(',') >= 0) aParts = sNorm.Split(',');
            else                              aParts = new string[] { sNorm };
            foreach (string sPart in aParts)
            {
                string sItem = (sPart == null ? "" : sPart).Trim();
                if (sItem.Length > 0) lsItems.Add(sItem);
            }
            return lsItems;
        }

        // Returns the full ordered list of keys.
        public List<string> keys()
        {
            List<string> l = new List<string>();
            foreach (Pair p in Pairs) l.Add(p.Key);
            return l;
        }
    }

    // Parse an .inix file. The returned list of sections preserves
    // file order. Implicit "[Global]" is created if the file starts
    // with key=value pairs before any explicit section header. The
    // returned list is empty if the file has no sections and no
    // top-level keys.
    public static List<Section> read(string sPath)
    {
        if (string.IsNullOrEmpty(sPath)) throw new ArgumentException("InixCodec.read requires a path.");
        if (!File.Exists(sPath)) throw new FileNotFoundException(".inix file not found.", sPath);
        string[] aLines = File.ReadAllLines(sPath, new UTF8Encoding(true));
        return parseLines(aLines);
    }

    // parseLines: the actual state machine. Separated from read so
    // unit tests can drive it without disk I/O.
    public static List<Section> parseLines(string[] aLines)
    {
        List<Section> lsSections = new List<Section>();
        Section secCurrent = null;       // null means "no section yet"
        string sPendingKey = null;       // multi-line value accumulator state
        StringBuilder sbValue = null;
        string sFenceToken = null;       // "`" or "\"\"\"" when inside a fenced value
        bool bSkipSection = false;       // section commented out via [;Name]

        // Helper to commit a pending multi-line value to the
        // current section. Strips the LAST trailing newline so that
        // the value doesn't carry an extra CRLF from the closing
        // line break.
        Action commitPending = delegate()
        {
            if (sPendingKey == null) return;
            string sFinal = (sbValue != null) ? sbValue.ToString() : "";
            // Drop one trailing newline if present.
            if (sFinal.EndsWith("\r\n")) sFinal = sFinal.Substring(0, sFinal.Length - 2);
            else if (sFinal.EndsWith("\n")) sFinal = sFinal.Substring(0, sFinal.Length - 1);
            if (secCurrent != null && !bSkipSection)
                secCurrent.Pairs.Add(new Pair(sPendingKey, sFinal));
            sPendingKey = null;
            sbValue = null;
            sFenceToken = null;
        };

        for (int i = 0; i < aLines.Length; i++)
        {
            string sRaw = aLines[i] ?? "";

            // Fenced multi-line value: accept the line VERBATIM
            // until a closing fence appears on a line by itself.
            if (sFenceToken != null)
            {
                if (sRaw.Trim() == sFenceToken)
                {
                    // Closing fence: commit and exit fenced mode.
                    commitPending();
                    continue;
                }
                if (sbValue.Length > 0) sbValue.Append("\r\n");
                sbValue.Append(sRaw);
                continue;
            }

            string sTrim = sRaw.Trim();

            // Plain multi-line value: accumulate lines until we see
            // a line that looks like a section header or a key=
            // line. The plain form requires that continuation lines
            // do NOT start with '[' or contain '=' as a meaningful
            // separator. Comment lines (starting with ';') in plain-
            // multi-line mode are treated as ordinary value content.
            if (sPendingKey != null)
            {
                bool bLooksLikeSection = sTrim.StartsWith("[") && sTrim.EndsWith("]");
                bool bLooksLikeKey = !sTrim.StartsWith(";")
                                     && sTrim.IndexOf('=') > 0
                                     && !looksLikePartOfValue(sTrim);
                if (bLooksLikeSection || bLooksLikeKey)
                {
                    commitPending();
                    // Fall through to process this line as the new
                    // section or key.
                }
                else
                {
                    if (sbValue.Length > 0) sbValue.Append("\r\n");
                    sbValue.Append(sRaw);
                    continue;
                }
            }

            // Section header.
            if (sTrim.StartsWith("[") && sTrim.EndsWith("]") && sTrim.Length >= 2)
            {
                string sInner = sTrim.Substring(1, sTrim.Length - 2);
                bSkipSection = sInner.StartsWith(";");
                if (bSkipSection)
                {
                    // We still create a section with an empty name
                    // and no pairs, but mark it skipped so we don't
                    // store anything? Simpler: just don't create a
                    // section at all -- the comment-out is total.
                    secCurrent = null;
                    continue;
                }
                if (sInner.Length == 0)
                {
                    // Anonymous section in a list -- assign a
                    // Record<N> name based on position. The N is
                    // (count of unnamed/recordNNN sections + 1)
                    // counting only data-bearing sections.
                    int iN = countAnonymousOrRecord(lsSections) + 1;
                    sInner = "Record" + iN;
                }
                secCurrent = new Section(sInner);
                lsSections.Add(secCurrent);
                continue;
            }

            // Skip comment and blank lines outside a value.
            if (sTrim.Length == 0) continue;
            if (sTrim.StartsWith(";") || sTrim.StartsWith("#")) continue;
            if (bSkipSection) continue;

            // Key=Value line.
            int iEq = sRaw.IndexOf('=');
            if (iEq <= 0) continue;  // not a valid key=value
            string sKey = sRaw.Substring(0, iEq).Trim();
            string sVal = sRaw.Substring(iEq + 1);

            // If no section opened yet, create an implicit
            // [Global] section.
            if (secCurrent == null)
            {
                secCurrent = new Section("Global");
                lsSections.Add(secCurrent);
            }

            string sValTrim = sVal.Trim();
            // Fenced multi-line starts with key=` or key=""" on a
            // line by itself (the equals is part of the line; the
            // fence token is the entire remainder).
            if (sValTrim == "`" || sValTrim == "\"\"\"")
            {
                sPendingKey = sKey;
                sbValue = new StringBuilder();
                sFenceToken = sValTrim;
                continue;
            }
            // Plain multi-line starts with key= (empty value on
            // this line). The accumulator gathers subsequent lines
            // until the next key= or section header.
            if (sValTrim.Length == 0)
            {
                sPendingKey = sKey;
                sbValue = new StringBuilder();
                continue;
            }
            // Single-line value.  Optional quoting, as in traditional .ini (and
            // .inix is a consistent superset of it):  if the value begins AND ends
            // with a double quote, those quotes are delimiters, not content.  Only
            // the outermost pair is removed, and everything between them is kept
            // exactly, including spaces.  So:
            //     Key=""          an empty string  (explicit, unambiguous)
            //     Key=" dog"      keeps the leading space
            //     Key="dog "      keeps the trailing space
            //     Key=""dog""     the literal text  "dog"  (with its quotes)
            // Quoting is optional: an unquoted value is trimmed, as before, so
            // existing files are unaffected.  Note this runs AFTER the two
            // multi-line tests above, so Key="" is an empty single-line value and
            // does NOT start a multi-line accumulation.
            if (sValTrim.Length >= 2 && sValTrim.StartsWith("\"") && sValTrim.EndsWith("\""))
                sValTrim = sValTrim.Substring(1, sValTrim.Length - 2);
            secCurrent.Pairs.Add(new Pair(sKey, sValTrim));
        }

        // EOF: commit any open multi-line value.
        commitPending();
        return lsSections;
    }

    // looksLikePartOfValue: a heuristic to keep the plain-multi-line
    // parser from prematurely closing a value when a continuation
    // line happens to contain an '=' character. We only treat
    // 'x=y' as a key if the part before '=' is a plausible
    // identifier (letters, digits, underscores). For values that
    // contain natural-language text with arbitrary '=', users
    // should use the fenced form.
    private static bool looksLikePartOfValue(string sLine)
    {
        int iEq = sLine.IndexOf('=');
        if (iEq <= 0) return true;
        string sBefore = sLine.Substring(0, iEq).Trim();
        if (sBefore.Length == 0) return true;
        foreach (char c in sBefore)
        {
            if (!char.IsLetterOrDigit(c) && c != '_' && c != ' ' && c != '-')
                return true;
        }
        return false;
    }

    private static int countAnonymousOrRecord(List<Section> l)
    {
        int n = 0;
        foreach (Section s in l)
            if (s.Name != null && s.Name.StartsWith("Record",
                StringComparison.OrdinalIgnoreCase)) n++;
        return n;
    }

    // ----------------------------------------------------------------
    // Write helpers. Output is UTF-8 with BOM and CRLF line endings,
    // matching the rest of DbDo's text-file conventions.
    // ----------------------------------------------------------------

    // Choose the right fence for a value: prefer plain multi-line
    // (no fence) when the value has no '=' or '[' that would
    // confuse the plain parser; prefer ` when '"""' is in the
    // value; prefer """ when "`" is in the value; fall back to
    // ` as a default for fenced values.
    private static string chooseFence(string sValue)
    {
        if (sValue == null) return null;
        bool bHasEq     = sValue.IndexOf('=') >= 0;
        bool bHasBkt    = sValue.IndexOf('[') >= 0;
        bool bMultiline = sValue.IndexOf('\n') >= 0 || sValue.IndexOf('\r') >= 0;
        if (!bMultiline && !bHasEq && !bHasBkt) return null;  // single-line literal OK
        if (!bMultiline) return "`";  // single-line but with = or [ -> fence it for safety
        // Multi-line. Choose a fence not present as a sole line.
        bool bBacktickFree = !containsSoleLine(sValue, "`");
        bool bTriquoteFree = !containsSoleLine(sValue, "\"\"\"");
        if (bBacktickFree) return "`";
        if (bTriquoteFree) return "\"\"\"";
        // Both candidate fences collide. Backtick is rare in
        // real-world text; prefer it and hope.
        return "`";
    }

    private static bool containsSoleLine(string sValue, string sToken)
    {
        // Split on \n; a "sole line" is one whose trim equals the token.
        int i = 0;
        while (i < sValue.Length)
        {
            int j = sValue.IndexOf('\n', i);
            int end = (j < 0) ? sValue.Length : j;
            string sLine = sValue.Substring(i, end - i).TrimEnd('\r').Trim();
            if (sLine == sToken) return true;
            i = (j < 0) ? sValue.Length : j + 1;
        }
        return false;
    }

    // writeAsConfig: serialize a list of named sections as an
    // .inix configuration file. Sections are written in the order
    // given; pairs within each section in the order given.
    public static void writeAsConfig(string sPath, List<Section> lsSections)
    {
        if (string.IsNullOrEmpty(sPath)) throw new ArgumentException("writeAsConfig requires a path.");
        if (lsSections == null) throw new ArgumentNullException("lsSections");
        using (StreamWriter w = new StreamWriter(sPath, false, new UTF8Encoding(true)))
        {
            w.NewLine = "\r\n";
            bool bFirst = true;
            foreach (Section sec in lsSections)
            {
                if (sec == null) continue;
                if (!bFirst) w.WriteLine();
                bFirst = false;
                if (!string.IsNullOrEmpty(sec.Name)
                    && !string.Equals(sec.Name, "Global", StringComparison.OrdinalIgnoreCase))
                    w.WriteLine("[" + sec.Name + "]");
                writePairs(w, sec.Pairs);
            }
        }
    }

    // writeAsTable: serialize a sequence of records as an .inix
    // list-of-records file. The leading-zero width on the section
    // name [RecordNNN] is chosen so ASCII sort matches numeric
    // order. Pairs within each record are written in the order
    // they appear in lsFields; values missing from a record are
    // simply omitted (no "key=" with empty value).
    public static void writeAsTable(string sPath, List<string> lsFields,
                                    List<Dictionary<string, string>> lsRows)
    {
        if (string.IsNullOrEmpty(sPath)) throw new ArgumentException("writeAsTable requires a path.");
        if (lsFields == null || lsRows == null) throw new ArgumentNullException();
        int n = lsRows.Count;
        int iWidth = (n == 0) ? 1 : (int)Math.Floor(Math.Log10(n)) + 1;
        string sFmt = "D" + iWidth;
        using (StreamWriter w = new StreamWriter(sPath, false, new UTF8Encoding(true)))
        {
            w.NewLine = "\r\n";
            for (int i = 0; i < n; i++)
            {
                if (i > 0) w.WriteLine();
                w.WriteLine("[Record" + (i + 1).ToString(sFmt) + "]");
                Dictionary<string, string> dRow = lsRows[i];
                if (dRow == null) continue;
                List<Pair> lsPairs = new List<Pair>();
                foreach (string sF in lsFields)
                {
                    string sV;
                    if (!dRow.TryGetValue(sF, out sV)) continue;
                    if (sV == null) continue;
                    lsPairs.Add(new Pair(sF, sV));
                }
                writePairs(w, lsPairs);
            }
        }
    }

    private static void writePairs(StreamWriter w, List<Pair> lsPairs)
    {
        foreach (Pair p in lsPairs)
        {
            string sKey = p.Key ?? "";
            string sVal = p.Value ?? "";
            string sFence = chooseFence(sVal);
            if (sFence == null)
            {
                // Quote the value when writing it bare would not read back exactly.
                // Three cases need it: an empty value (a bare "Key =" would re-read as
                // the START of a multi-line value), a value with a leading or trailing
                // space (which a bare value loses to trimming), and a value that itself
                // begins and ends with a double quote (whose own quotes would be taken
                // as delimiters on the way back in).  Wrapping adds one outer pair and
                // the reader strips exactly one, so the original text survives.
                bool bNeedsQuote = (sVal.Length == 0)
                    || (sVal != sVal.Trim())
                    || (sVal.Length >= 2 && sVal.StartsWith("\"") && sVal.EndsWith("\""));
                if (bNeedsQuote) w.WriteLine(sKey + " = \"" + sVal + "\"");
                else w.WriteLine(sKey + " = " + sVal);
            }
            else
            {
                // Fenced form, even if value happens to be single
                // line -- chooseFence returns non-null when the
                // single line contains '=' or '['.
                w.WriteLine(sKey + "=" + sFence);
                // Write value lines verbatim. Normalize line
                // endings to CRLF so the output is a Windows file.
                string sNormalized = sVal.Replace("\r\n", "\n").Replace("\r", "\n");
                foreach (string sLn in sNormalized.Split('\n'))
                    w.WriteLine(sLn);
                w.WriteLine(sFence);
            }
        }
    }

    // writeValue: surgically set, replace, or remove ONE key in an
    // .inix file, preserving every comment and every other line --
    // unlike writeAsConfig, which rewrites the whole file and
    // would drop comments. Used for the per-user settings file,
    // where the shipped template's documentation comments must
    // survive machine writes.
    //
    // Inix-aware in both directions: a value containing a newline
    // is written in the fenced form (key=` ... `), an existing
    // fenced value is replaced or removed as a whole block, and
    // the scan skips over fenced blocks so a key= line INSIDE a
    // fenced value is never mistaken for a real key. Pass an
    // empty value to remove the key. Returns false on I/O failure
    // so callers can log it (this class stays log-independent).
    public static bool writeValue(string sPath, string sSection, string sKey, string sValue)
    {
        string sHeader = "[" + sSection + "]";
        List<string> lsLines = new List<string>();
        if (File.Exists(sPath))
        {
            try { lsLines.AddRange(File.ReadAllLines(sPath, new UTF8Encoding(true))); } catch { return false; }
        }

        int iSectionEnd = -1, iSectionStart = -1;
        for (int i = 0; i < lsLines.Count; i++)
        {
            string sTrim = lsLines[i].Trim();
            if (sTrim.Equals(sHeader, StringComparison.OrdinalIgnoreCase))
            {
                iSectionStart = i;
                iSectionEnd = lsLines.Count;
                for (int j = i + 1; j < lsLines.Count; j++)
                {
                    string sJ = lsLines[j].Trim();
                    if (sJ.StartsWith("[")) { iSectionEnd = j; break; }
                }
                break;
            }
        }

        // Render the new value as one or more lines: fenced when
        // it spans lines or itself contains '=' or '[' (the cases
        // where the plain form is unreliable), plain otherwise.
        List<string> lsNewLines = new List<string>();
        if (!string.IsNullOrEmpty(sValue))
        {
            bool bFenced = sValue.IndexOf('\n') >= 0
                        || sValue.IndexOf('=') >= 0
                        || sValue.TrimStart().StartsWith("[");
            if (bFenced)
            {
                lsNewLines.Add(sKey + "=`");
                string sNormalized = sValue.Replace("\r\n", "\n").Replace("\r", "\n");
                foreach (string sLn in sNormalized.Split('\n')) lsNewLines.Add(sLn);
                lsNewLines.Add("`");
            }
            else lsNewLines.Add(sKey + " = " + sValue);
        }

        if (iSectionStart < 0)
        {
            if (lsLines.Count > 0 && lsLines[lsLines.Count - 1].Trim().Length > 0) lsLines.Add("");
            lsLines.Add(sHeader);
            lsLines.AddRange(lsNewLines);
        }
        else
        {
            int iFound = -1, iFoundEnd = -1;  // inclusive line range of the existing entry
            for (int i = iSectionStart + 1; i < iSectionEnd; i++)
            {
                string sT = lsLines[i].Trim();
                if (sT.Length == 0) continue;
                if (sT.StartsWith(";") || sT.StartsWith("#")) continue;
                int iEq = sT.IndexOf('=');
                if (iEq <= 0) continue;
                string sName = sT.Substring(0, iEq).Trim();
                string sRest = sT.Substring(iEq + 1).Trim();
                bool bFenceOpen = (sRest == "`" || sRest == "\"\"\"");
                int iEntryEnd = i;
                if (bFenceOpen)
                {
                    iEntryEnd = iSectionEnd - 1;  // unterminated fence: rest of section
                    for (int j = i + 1; j < iSectionEnd; j++)
                        if (lsLines[j].Trim() == sRest) { iEntryEnd = j; break; }
                }
                if (string.Equals(sName, sKey, StringComparison.OrdinalIgnoreCase))
                { iFound = i; iFoundEnd = iEntryEnd; break; }
                i = iEntryEnd;  // skip past a fenced block belonging to another key
            }
            if (iFound >= 0)
            {
                lsLines.RemoveRange(iFound, iFoundEnd - iFound + 1);
                if (lsNewLines.Count > 0) lsLines.InsertRange(iFound, lsNewLines);
            }
            else if (lsNewLines.Count > 0)
            {
                lsLines.InsertRange(iSectionStart + 1, lsNewLines);
            }
        }

        try {
            using (StreamWriter w = new StreamWriter(sPath, false, new UTF8Encoding(true))) {
                w.NewLine = "\r\n";
                foreach (string sLn in lsLines) w.WriteLine(sLn);
            }
            return true;
        }
        catch { return false; }
    }

    // fileTask: the declared purpose of an .inix file, read from the "FileTask" key
    // of its implicit [Global] section -- e.g. "report" for a report definition or
    // "transfer" for an import map.  Returns "" when the file declares no FileTask,
    // is unreadable, or is absent.  The value is trimmed and lower-cased so callers
    // can compare it directly.  This is what lets a command that offers a pick-list
    // of .inix files show only the files whose task matches the command, so a report
    // picker never lists an import map or a settings file, and vice versa.
    public static string fileTask(string sPath)
    {
        try { return fileTask(read(sPath)); }
        catch { return ""; }
    }

    // fileTask overload for callers that have already parsed the file, so the
    // FileTask can be read without a second pass over the disk.
    public static string fileTask(List<Section> lsSections)
    {
        if (lsSections == null) return "";
        foreach (Section section in lsSections)
            if (string.Equals(section.Name, "Global", StringComparison.OrdinalIgnoreCase))
            {
                string sValue = section.get("FileTask");
                return string.IsNullOrEmpty(sValue) ? "" : sValue.Trim().ToLowerInvariant();
            }
        return "";
    }

    // writeArrayValue: the write side of the .inix array convention.  Stores the
    // items under sKey, choosing the presentation automatically: inline and
    // comma-separated when the items are few (up to six) and short and none contains
    // a space, comma, or backtick; otherwise one item per line, which writeValue
    // renders as a fenced block.  Blank items are dropped; an empty list clears the
    // key.
    public static bool writeArrayValue(string sPath, string sSection, string sKey, List<string> lsItems)
    {
        List<string> lsClean = new List<string>();
        if (lsItems != null)
            foreach (string sItem in lsItems)
            {
                string sOne = (sItem == null ? "" : sItem).Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ").Trim();
                if (sOne.Length > 0) lsClean.Add(sOne);
            }
        if (lsClean.Count == 0) return writeValue(sPath, sSection, sKey, "");

        bool bInline = lsClean.Count <= 6;
        if (bInline)
        {
            int iLength = 0;
            foreach (string sOne in lsClean)
            {
                if (sOne.IndexOf(' ') >= 0 || sOne.IndexOf(',') >= 0 || sOne.IndexOf('`') >= 0) { bInline = false; break; }
                iLength += sOne.Length + 2;
            }
            if (iLength > 80) bInline = false;
        }
        string sValue = bInline ? string.Join(", ", lsClean.ToArray())
                                : string.Join("\n", lsClean.ToArray());
        return writeValue(sPath, sSection, sKey, sValue);
    }
}

} // namespace Homer
