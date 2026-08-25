// Inix.cs (Homer.InixCodec + Homer.InixTable) -- the shared Inix toolkit.
// Portable across EdSharp, DbDo, and every other Homer Tool: copy this file
// as-is and reference it with `using Homer;`.
//
// .inix is a superset of classic .ini: ';' or '#' comments, [Section]
// headers, name = value lines, PLUS verbatim multi-line values (backtick or
// triple-quote fences), inline or fenced arrays, an implicit [Global]
// section, and order-preserving round trips.
//
// As TABULAR DATA, .inix is a screen-reader-friendly way to review a
// table: instead of many field values crowded onto one line, each record
// is a [RecordNNN] section whose fields sit on their own field = value
// lines, with fenced multi-line values when a value needs them. One
// record, one screen of related lines -- no column counting.
//
// InixCodec is the reader-writer for the format. InixTable (added
// 24 August 2026) is the GENERIC table-conversion layer: it moves tabular
// data between .inix, .csv, .tsv, Markdown pipe tables, and .xlsx
// workbooks, in any direction, with .inix as the home format. The .xlsx
// side is pure OpenXML over System.IO.Compression -- no Office, no ACE
// provider, no COM -- so it runs anywhere .NET runs. Compiling this file
// therefore needs references to System.IO.Compression.dll and
// System.Xml.dll (BuildEdSharp adds them).

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

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

// =====================================================================
// InputHistory: shared persistence scheme for recent-input lists, the
// same layout DbDo uses: slot keys term1, term2, ... termN under one
// section per command, newest first, up to historyCount entries
// (default 10). Pure list logic; the caller supplies key read/write
// delegates bound to whatever settings store the app uses (an
// InixCodec path, a classic .ini, or FileDir's dual layer), so one
// implementation serves EdSharp, FileDir, and DbDo.

// =====================================================================
// InixTable: generic tabular conversions with .inix as the home format.
// A table is fields (column names, in order) plus rows (one ordered
// dictionary of field -> string value per record). Values are always
// strings: the purpose is faithful, screen-reader-friendly REVIEW of
// data, not calculation.
// =====================================================================
public static class InixTable
{
    public class TableData
    {
        public List<string> Fields = new List<string>();
        public List<Dictionary<string, string>> Rows = new List<Dictionary<string, string>>();
    }

    // ---- reading and writing the home format ----

    public static TableData readInix(string sPath)
    {
        TableData table = new TableData();
        List<InixCodec.Section> lsSections = InixCodec.read(sPath);
        foreach (InixCodec.Section section in lsSections)
        {
            if (section == null || section.Pairs.Count == 0) continue;
            // A Global section in an otherwise table-shaped file is a
            // document header, not a row (same tolerance DbDo uses).
            if (string.Equals(section.Name, "Global", StringComparison.OrdinalIgnoreCase) && lsSections.Count > 1) continue;
            Dictionary<string, string> dRow = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (InixCodec.Pair pair in section.Pairs)
            {
                if (!dRow.ContainsKey(pair.Key)) dRow[pair.Key] = pair.Value ?? "";
                if (!containsField(table.Fields, pair.Key)) table.Fields.Add(pair.Key);
            }
            table.Rows.Add(dRow);
        }
        return table;
    }

    public static void writeInix(string sPath, TableData table)
    {
        InixCodec.writeAsTable(sPath, table.Fields, table.Rows);
    }

    // ---- CSV and TSV (RFC 4180: quotes, embedded delimiters, embedded
    // line breaks; tolerant of both CRLF and LF) ----

    public static TableData readDelimited(string sPath, char cDelimiter)
    {
        string sText = File.ReadAllText(sPath, detectEncoding(sPath));
        List<List<string>> llRecords = parseDelimited(sText, cDelimiter);
        return rowsFromGrid(llRecords);
    }

    public static void writeDelimited(string sPath, TableData table, char cDelimiter)
    {
        using (StreamWriter writer = new StreamWriter(sPath, false, new UTF8Encoding(true)))
        {
            writer.NewLine = "\r\n";
            writer.WriteLine(joinDelimited(table.Fields, cDelimiter));
            foreach (Dictionary<string, string> dRow in table.Rows)
            {
                List<string> lsValues = new List<string>();
                foreach (string sField in table.Fields)
                {
                    string sValue;
                    lsValues.Add(dRow != null && dRow.TryGetValue(sField, out sValue) ? (sValue ?? "") : "");
                }
                writer.WriteLine(joinDelimited(lsValues, cDelimiter));
            }
        }
    }

    static List<List<string>> parseDelimited(string sText, char cDelimiter)
    {
        List<List<string>> llRecords = new List<List<string>>();
        List<string> lsRecord = new List<string>();
        StringBuilder sbField = new StringBuilder();
        bool bQuoted = false;
        bool bAny = false;
        int i = 0;
        while (i < sText.Length)
        {
            char c = sText[i];
            if (bQuoted)
            {
                if (c == '"')
                {
                    if (i + 1 < sText.Length && sText[i + 1] == '"') { sbField.Append('"'); i += 2; continue; }
                    bQuoted = false; i++; continue;
                }
                sbField.Append(c); i++; continue;
            }
            if (c == '"' && sbField.Length == 0) { bQuoted = true; bAny = true; i++; continue; }
            if (c == cDelimiter) { lsRecord.Add(sbField.ToString()); sbField.Length = 0; bAny = true; i++; continue; }
            if (c == '\r' || c == '\n')
            {
                if (c == '\r' && i + 1 < sText.Length && sText[i + 1] == '\n') i++;
                i++;
                if (bAny || sbField.Length > 0 || lsRecord.Count > 0)
                {
                    lsRecord.Add(sbField.ToString()); sbField.Length = 0;
                    llRecords.Add(lsRecord); lsRecord = new List<string>();
                    bAny = false;
                }
                continue;
            }
            sbField.Append(c); bAny = true; i++;
        }
        if (bAny || sbField.Length > 0 || lsRecord.Count > 0)
        {
            lsRecord.Add(sbField.ToString());
            llRecords.Add(lsRecord);
        }
        return llRecords;
    }

    static string joinDelimited(List<string> lsValues, char cDelimiter)
    {
        StringBuilder sbLine = new StringBuilder();
        for (int i = 0; i < lsValues.Count; i++)
        {
            if (i > 0) sbLine.Append(cDelimiter);
            string sValue = lsValues[i] ?? "";
            bool bNeedsQuote = sValue.IndexOf(cDelimiter) >= 0 || sValue.IndexOf('"') >= 0
                || sValue.IndexOf('\r') >= 0 || sValue.IndexOf('\n') >= 0
                || (sValue.Length > 0 && (sValue[0] == ' ' || sValue[sValue.Length - 1] == ' '));
            if (bNeedsQuote) sbLine.Append('"').Append(sValue.Replace("\"", "\"\"")).Append('"');
            else sbLine.Append(sValue);
        }
        return sbLine.ToString();
    }

    // ---- Markdown pipe tables. A cell cannot hold a real line break,
    // so line breaks become <br> on the way out and back again on the
    // way in; '|' is escaped as '\|'. ----

    public static TableData readMarkdown(string sPath)
    {
        List<List<string>> llRecords = new List<List<string>>();
        foreach (string sRaw in File.ReadAllLines(sPath, detectEncoding(sPath)))
        {
            string sLine = sRaw.Trim();
            if (sLine.Length < 2 || sLine[0] != '|') continue;
            if (isMarkdownSeparator(sLine)) continue;
            llRecords.Add(splitMarkdownRow(sLine));
        }
        return rowsFromGrid(llRecords);
    }

    public static void writeMarkdown(string sPath, TableData table)
    {
        using (StreamWriter writer = new StreamWriter(sPath, false, new UTF8Encoding(true)))
        {
            writer.NewLine = "\r\n";
            writer.WriteLine(markdownRow(table.Fields));
            StringBuilder sbRule = new StringBuilder("|");
            for (int i = 0; i < table.Fields.Count; i++) sbRule.Append(" --- |");
            writer.WriteLine(sbRule.ToString());
            foreach (Dictionary<string, string> dRow in table.Rows)
            {
                List<string> lsValues = new List<string>();
                foreach (string sField in table.Fields)
                {
                    string sValue;
                    lsValues.Add(dRow != null && dRow.TryGetValue(sField, out sValue) ? (sValue ?? "") : "");
                }
                writer.WriteLine(markdownRow(lsValues));
            }
        }
    }

    static bool isMarkdownSeparator(string sLine)
    {
        foreach (char c in sLine) if (c != '|' && c != '-' && c != ':' && c != ' ' && c != '\t') return false;
        return sLine.IndexOf('-') >= 0;
    }

    static List<string> splitMarkdownRow(string sLine)
    {
        List<string> lsCells = new List<string>();
        StringBuilder sbCell = new StringBuilder();
        // Interior of "| a | b |": strip one leading and one trailing bar.
        string sInner = sLine.Substring(1);
        if (sInner.EndsWith("|")) sInner = sInner.Substring(0, sInner.Length - 1);
        int i = 0;
        while (i < sInner.Length)
        {
            char c = sInner[i];
            if (c == '\\' && i + 1 < sInner.Length && sInner[i + 1] == '|') { sbCell.Append('|'); i += 2; continue; }
            if (c == '|') { lsCells.Add(cleanMarkdownCell(sbCell.ToString())); sbCell.Length = 0; i++; continue; }
            sbCell.Append(c); i++;
        }
        lsCells.Add(cleanMarkdownCell(sbCell.ToString()));
        return lsCells;
    }

    static string cleanMarkdownCell(string sCell)
    {
        return sCell.Trim().Replace("<br>", "\n").Replace("<br/>", "\n").Replace("<br />", "\n");
    }

    static string markdownRow(List<string> lsValues)
    {
        StringBuilder sbRow = new StringBuilder("|");
        foreach (string sRaw in lsValues)
        {
            string sValue = (sRaw ?? "").Replace("|", "\\|");
            sValue = sValue.Replace("\r\n", "<br>").Replace("\r", "<br>").Replace("\n", "<br>");
            sbRow.Append(' ').Append(sValue).Append(" |");
        }
        return sbRow.ToString();
    }

    // ---- XLSX, pure OpenXML: a workbook is a zip of XML parts. Writing
    // uses inline strings (no shared-string table needed); reading
    // handles both inline and shared strings and reads every cell as
    // the text it shows. First worksheet only -- the format's job here
    // is table review, not workbook management. ----

    public static TableData readXlsx(string sPath)
    {
        using (FileStream stream = new FileStream(sPath, FileMode.Open, FileAccess.Read))
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
        {
            List<string> lsShared = readSharedStrings(archive);
            ZipArchiveEntry entrySheet = findFirstSheet(archive);
            if (entrySheet == null) throw new InvalidDataException("No worksheet was found inside the workbook.");
            List<List<string>> llRecords = new List<List<string>>();
            using (Stream streamSheet = entrySheet.Open())
            using (XmlReader reader = XmlReader.Create(streamSheet))
            {
                List<string> lsRow = null;
                int iColumn = 0;
                string sCellType = "";
                bool bInValue = false, bInInline = false;
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (reader.LocalName == "row") { lsRow = new List<string>(); iColumn = 0; }
                        else if (reader.LocalName == "c" && lsRow != null)
                        {
                            string sRef = reader.GetAttribute("r");
                            int iAt = (sRef != null) ? columnIndex(sRef) : iColumn;
                            while (lsRow.Count < iAt) lsRow.Add("");
                            iColumn = iAt + 1;
                            sCellType = reader.GetAttribute("t") ?? "";
                            lsRow.Add("");
                        }
                        else if (reader.LocalName == "v") bInValue = true;
                        else if (reader.LocalName == "is") bInInline = true;
                        else if (reader.LocalName == "t" && bInInline && lsRow != null && lsRow.Count > 0 && !reader.IsEmptyElement)
                        {
                            lsRow[lsRow.Count - 1] = lsRow[lsRow.Count - 1] + reader.ReadElementContentAsString();
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.Text && bInValue && lsRow != null && lsRow.Count > 0)
                    {
                        string sValue = reader.Value;
                        if (sCellType == "s")
                        {
                            int iShared;
                            if (int.TryParse(sValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out iShared)
                                && iShared >= 0 && iShared < lsShared.Count) sValue = lsShared[iShared];
                        }
                        lsRow[lsRow.Count - 1] = sValue;
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        if (reader.LocalName == "v") bInValue = false;
                        else if (reader.LocalName == "is") bInInline = false;
                        else if (reader.LocalName == "row" && lsRow != null) { llRecords.Add(lsRow); lsRow = null; }
                    }
                }
            }
            return rowsFromGrid(llRecords);
        }
    }

    public static void writeXlsx(string sPath, TableData table)
    {
        using (FileStream stream = new FileStream(sPath, FileMode.Create, FileAccess.Write))
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            writeArchiveText(archive, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                + "</Types>");
            writeArchiveText(archive, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                + "</Relationships>");
            writeArchiveText(archive, "xl/workbook.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<sheets><sheet name=\"Table\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            writeArchiveText(archive, "xl/_rels/workbook.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + "</Relationships>");
            ZipArchiveEntry entrySheet = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using (Stream streamSheet = entrySheet.Open())
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Encoding = new UTF8Encoding(false);
                using (XmlWriter writer = XmlWriter.Create(streamSheet, settings))
                {
                    string sNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                    writer.WriteStartElement("worksheet", sNs);
                    writer.WriteStartElement("sheetData", sNs);
                    writeXlsxRow(writer, sNs, 1, table.Fields);
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        Dictionary<string, string> dRow = table.Rows[i];
                        List<string> lsValues = new List<string>();
                        foreach (string sField in table.Fields)
                        {
                            string sValue;
                            lsValues.Add(dRow != null && dRow.TryGetValue(sField, out sValue) ? (sValue ?? "") : "");
                        }
                        writeXlsxRow(writer, sNs, i + 2, lsValues);
                    }
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
            }
        }
    }

    static void writeXlsxRow(XmlWriter writer, string sNs, int iRow, List<string> lsValues)
    {
        writer.WriteStartElement("row", sNs);
        writer.WriteAttributeString("r", iRow.ToString(CultureInfo.InvariantCulture));
        for (int i = 0; i < lsValues.Count; i++)
        {
            writer.WriteStartElement("c", sNs);
            writer.WriteAttributeString("r", columnName(i) + iRow.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("t", "inlineStr");
            writer.WriteStartElement("is", sNs);
            writer.WriteStartElement("t", sNs);
            writer.WriteAttributeString("xml", "space", null, "preserve");
            writer.WriteString(lsValues[i] ?? "");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }

    static List<string> readSharedStrings(ZipArchive archive)
    {
        List<string> lsShared = new List<string>();
        ZipArchiveEntry entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry == null) return lsShared;
        using (Stream stream = entry.Open())
        using (XmlReader reader = XmlReader.Create(stream))
        {
            StringBuilder sbItem = null;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si") sbItem = new StringBuilder();
                else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && sbItem != null && !reader.IsEmptyElement)
                    sbItem.Append(reader.ReadElementContentAsString());
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "si" && sbItem != null)
                { lsShared.Add(sbItem.ToString()); sbItem = null; }
            }
        }
        return lsShared;
    }

    static ZipArchiveEntry findFirstSheet(ZipArchive archive)
    {
        ZipArchiveEntry entry = archive.GetEntry("xl/worksheets/sheet1.xml");
        if (entry != null) return entry;
        foreach (ZipArchiveEntry candidate in archive.Entries)
            if (candidate.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                && candidate.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return candidate;
        return null;
    }

    // ---- shared plumbing ----

    static TableData rowsFromGrid(List<List<string>> llRecords)
    {
        TableData table = new TableData();
        if (llRecords.Count == 0) return table;
        foreach (string sHeader in llRecords[0])
        {
            string sName = (sHeader ?? "").Trim();
            if (sName.Length == 0) sName = "Field" + (table.Fields.Count + 1).ToString(CultureInfo.InvariantCulture);
            while (containsField(table.Fields, sName)) sName += "2";
            table.Fields.Add(sName);
        }
        for (int i = 1; i < llRecords.Count; i++)
        {
            List<string> lsRecord = llRecords[i];
            bool bAll = true;
            foreach (string sValue in lsRecord) if ((sValue ?? "").Length > 0) { bAll = false; break; }
            if (bAll) continue;
            Dictionary<string, string> dRow = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int j = 0; j < table.Fields.Count; j++)
                dRow[table.Fields[j]] = (j < lsRecord.Count) ? (lsRecord[j] ?? "") : "";
            table.Rows.Add(dRow);
        }
        return table;
    }

    static bool containsField(List<string> lsFields, string sName)
    {
        foreach (string sField in lsFields)
            if (string.Equals(sField, sName, StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    static int columnIndex(string sCellRef)
    {
        int iIndex = 0;
        foreach (char c in sCellRef)
        {
            if (c >= 'A' && c <= 'Z') iIndex = iIndex * 26 + (c - 'A' + 1);
            else if (c >= 'a' && c <= 'z') iIndex = iIndex * 26 + (c - 'a' + 1);
            else break;
        }
        return (iIndex > 0) ? iIndex - 1 : 0;
    }

    static string columnName(int iIndex)
    {
        string sName = "";
        iIndex++;
        while (iIndex > 0)
        {
            int iRemainder = (iIndex - 1) % 26;
            sName = ((char)('A' + iRemainder)) + sName;
            iIndex = (iIndex - 1) / 26;
        }
        return sName;
    }

    static Encoding detectEncoding(string sPath)
    {
        // UTF-8 with or without a byte-order mark covers every Homer
        // file; the StreamReader BOM sniff handles UTF-16 arrivals.
        using (StreamReader reader = new StreamReader(sPath, new UTF8Encoding(false), true))
        { reader.Peek(); return reader.CurrentEncoding; }
    }

    static void writeArchiveText(ZipArchive archive, string sName, string sContent)
    {
        ZipArchiveEntry entry = archive.CreateEntry(sName);
        using (Stream stream = entry.Open())
        {
            byte[] aBytes = new UTF8Encoding(false).GetBytes(sContent);
            stream.Write(aBytes, 0, aBytes.Length);
        }
    }


    // ---- Embedded inix tables in Markdown (.mdx, or .md that opts in).
    // A fenced code block whose info string is "inix" holds records in
    // the tabular inix form; expansion replaces the block with a real
    // Markdown table, which Pandoc then renders as a real table in
    // docx, HTML, and every other output -- no filter, no extension.
    // A [Global] section at the top of a block may supply options, but
    // none is required:
    //     caption = A caption printed under the table
    //     fields  = the columns to show, in order (default: every
    //               field, in first-seen order)
    // Pipe tables carry single-line cells; if any cell is multi-line,
    // a GRID table is written instead, because grid tables are the
    // Markdown form that holds multi-line cells. Text outside the
    // fences passes through untouched, so a file with no inix blocks
    // expands to itself. ----

    public static string expandMarkdownText(string sText)
    {
        string[] aLines = sText.Replace("\r\n", "\n").Split('\n');
        StringBuilder sbOut = new StringBuilder();
        List<string> lsBlock = null;
        string sFenceClose = null;
        foreach (string sLine in aLines)
        {
            if (lsBlock == null)
            {
                string sTrim = sLine.TrimStart();
                if ((sTrim.StartsWith("```") || sTrim.StartsWith("~~~"))
                    && sTrim.Substring(3).Trim().ToLowerInvariant() == "inix")
                {
                    lsBlock = new List<string>();
                    sFenceClose = sTrim.Substring(0, 3);
                    continue;
                }
                sbOut.Append(sLine).Append("\r\n");
                continue;
            }
            if (sLine.TrimStart().StartsWith(sFenceClose))
            {
                sbOut.Append(renderBlock(lsBlock));
                lsBlock = null;
                continue;
            }
            lsBlock.Add(sLine);
        }
        // An unclosed block is a mistake worth surfacing gently: it is
        // rendered as far as it goes rather than swallowed.
        if (lsBlock != null) sbOut.Append(renderBlock(lsBlock));
        return sbOut.ToString();
    }

    public static void expandMarkdownFile(string sSourcePath, string sDestPath)
    {
        string sText = File.ReadAllText(sSourcePath, detectEncoding(sSourcePath));
        File.WriteAllText(sDestPath, expandMarkdownText(sText), new UTF8Encoding(true));
    }

    static string renderBlock(List<string> lsBlockLines)
    {
        List<InixCodec.Section> lsSections = InixCodec.parseLines(lsBlockLines.ToArray());
        string sCaption = "";
        List<string> lsWanted = null;
        TableData table = new TableData();
        foreach (InixCodec.Section section in lsSections)
        {
            if (section == null || section.Pairs.Count == 0) continue;
            if (string.Equals(section.Name, "Global", StringComparison.OrdinalIgnoreCase) && lsSections.Count > 1)
            {
                string sValue = section.get("caption");
                if (sValue != null) sCaption = sValue;
                List<string> lsFields = section.getArray("fields");
                if (lsFields != null && lsFields.Count > 0) lsWanted = lsFields;
                continue;
            }
            Dictionary<string, string> dRow = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (InixCodec.Pair pair in section.Pairs)
            {
                if (!dRow.ContainsKey(pair.Key)) dRow[pair.Key] = pair.Value ?? "";
                if (!containsField(table.Fields, pair.Key)) table.Fields.Add(pair.Key);
            }
            table.Rows.Add(dRow);
        }
        if (lsWanted != null)
        {
            List<string> lsOrdered = new List<string>();
            foreach (string sField in lsWanted)
                if (containsField(table.Fields, sField.Trim())) lsOrdered.Add(sField.Trim());
            if (lsOrdered.Count > 0) table.Fields = lsOrdered;
        }
        if (table.Fields.Count == 0) return "\r\n";
        bool bMultiline = false;
        foreach (Dictionary<string, string> dRow in table.Rows)
            foreach (string sValue in dRow.Values)
                if (sValue != null && (sValue.IndexOf('\n') >= 0 || sValue.IndexOf('\r') >= 0)) { bMultiline = true; break; }
        StringBuilder sbTable = new StringBuilder();
        sbTable.Append("\r\n");
        if (bMultiline) appendGridTable(sbTable, table);
        else
        {
            sbTable.Append(markdownRow(table.Fields)).Append("\r\n");
            StringBuilder sbRule = new StringBuilder("|");
            for (int i = 0; i < table.Fields.Count; i++) sbRule.Append(" --- |");
            sbTable.Append(sbRule.ToString()).Append("\r\n");
            foreach (Dictionary<string, string> dRow in table.Rows)
                sbTable.Append(markdownRow(valuesFor(table, dRow))).Append("\r\n");
        }
        if (sCaption.Length > 0) sbTable.Append("\r\n: ").Append(sCaption.Replace("\r", " ").Replace("\n", " ")).Append("\r\n");
        sbTable.Append("\r\n");
        return sbTable.ToString();
    }

    static List<string> valuesFor(TableData table, Dictionary<string, string> dRow)
    {
        List<string> lsValues = new List<string>();
        foreach (string sField in table.Fields)
        {
            string sValue;
            lsValues.Add(dRow != null && dRow.TryGetValue(sField, out sValue) ? (sValue ?? "") : "");
        }
        return lsValues;
    }

    // A Pandoc GRID table: dashed borders, '=' under the header, cells
    // that may span several lines. Every cell is padded to the column
    // width, which is sized to the longest line in that column.
    static void appendGridTable(StringBuilder sbOut, TableData table)
    {
        int iColumns = table.Fields.Count;
        List<List<List<string>>> lllRows = new List<List<List<string>>>();
        List<List<string>> llHeader = new List<List<string>>();
        foreach (string sField in table.Fields) llHeader.Add(cellLines(sField));
        lllRows.Add(llHeader);
        foreach (Dictionary<string, string> dRow in table.Rows)
        {
            List<List<string>> llRow = new List<List<string>>();
            foreach (string sValue in valuesFor(table, dRow)) llRow.Add(cellLines(sValue));
            lllRows.Add(llRow);
        }
        int[] aWidths = new int[iColumns];
        foreach (List<List<string>> llRow in lllRows)
            for (int i = 0; i < iColumns; i++)
                foreach (string sLine in llRow[i])
                    if (sLine.Length > aWidths[i]) aWidths[i] = sLine.Length;
        for (int i = 0; i < iColumns; i++) if (aWidths[i] < 3) aWidths[i] = 3;
        string sDashes = gridRule(aWidths, '-');
        string sEquals = gridRule(aWidths, '=');
        sbOut.Append(sDashes).Append("\r\n");
        for (int iRow = 0; iRow < lllRows.Count; iRow++)
        {
            List<List<string>> llRow = lllRows[iRow];
            int iHeight = 1;
            foreach (List<string> lsCell in llRow) if (lsCell.Count > iHeight) iHeight = lsCell.Count;
            for (int iLine = 0; iLine < iHeight; iLine++)
            {
                sbOut.Append('|');
                for (int i = 0; i < iColumns; i++)
                {
                    string sCell = (iLine < llRow[i].Count) ? llRow[i][iLine] : "";
                    sbOut.Append(' ').Append(sCell.PadRight(aWidths[i])).Append(" |");
                }
                sbOut.Append("\r\n");
            }
            sbOut.Append(iRow == 0 ? sEquals : sDashes).Append("\r\n");
        }
    }

    static List<string> cellLines(string sValue)
    {
        List<string> lsLines = new List<string>();
        foreach (string sLine in (sValue ?? "").Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
            lsLines.Add(sLine);
        return lsLines;
    }

    static string gridRule(int[] aWidths, char cFill)
    {
        StringBuilder sbRule = new StringBuilder("+");
        foreach (int iWidth in aWidths) sbRule.Append(new string(cFill, iWidth + 2)).Append('+');
        return sbRule.ToString();
    }

    // ---- in-memory helpers for hosts (EdSharp code blocks, DbDo) ----

    // Parse delimited TEXT already in memory (e.g. captured command
    // output) with the same RFC 4180 state machine used for files.
    public static TableData tableFromDelimitedText(string sText, char cDelimiter)
    {
        return rowsFromGrid(parseDelimited(sText ?? "", cDelimiter));
    }

    // Render a table as Markdown text: a pipe table normally, a grid
    // table when any cell is multi-line (the same decision the .mdx
    // expansion makes), so the result converts to real tables in docx
    // and HTML.
    public static string tableToMarkdown(TableData table)
    {
        if (table == null || table.Fields.Count == 0) return "";
        bool bMultiline = false;
        foreach (Dictionary<string, string> dRow in table.Rows)
            foreach (string sValue in dRow.Values)
                if (sValue != null && (sValue.IndexOf('\n') >= 0 || sValue.IndexOf('\r') >= 0)) { bMultiline = true; break; }
        StringBuilder sbTable = new StringBuilder();
        if (bMultiline) appendGridTable(sbTable, table);
        else
        {
            sbTable.Append(markdownRow(table.Fields)).Append("\r\n");
            StringBuilder sbRule = new StringBuilder("|");
            for (int i = 0; i < table.Fields.Count; i++) sbRule.Append(" --- |");
            sbTable.Append(sbRule.ToString()).Append("\r\n");
            foreach (Dictionary<string, string> dRow in table.Rows)
                sbTable.Append(markdownRow(valuesFor(table, dRow))).Append("\r\n");
        }
        return sbTable.ToString();
    }

    // ---- the one-call converter ----

    // Reads by the source extension, writes by the destination
    // extension. Known on both sides: inix, csv, tsv (or tab), md (or
    // markdown), xlsx. Throws with a plain message for anything else.
    public static void convertFile(string sSourcePath, string sDestPath)
    {
        TableData table = readAny(sSourcePath);
        writeAny(sDestPath, table);
    }

    public static TableData readAny(string sPath)
    {
        switch (extensionOf(sPath))
        {
            case "inix": return readInix(sPath);
            case "csv": return readDelimited(sPath, ',');
            case "tsv": case "tab": return readDelimited(sPath, '\t');
            case "md": case "markdown": return readMarkdown(sPath);
            case "xlsx": return readXlsx(sPath);
            default: throw new ArgumentException("Reading ." + extensionOf(sPath) + " is not supported. Supported: .inix, .csv, .tsv, .md, .xlsx");
        }
    }

    public static void writeAny(string sPath, TableData table)
    {
        switch (extensionOf(sPath))
        {
            case "inix": writeInix(sPath, table); break;
            case "csv": writeDelimited(sPath, table, ','); break;
            case "tsv": case "tab": writeDelimited(sPath, table, '\t'); break;
            case "md": case "markdown": writeMarkdown(sPath, table); break;
            case "xlsx": writeXlsx(sPath, table); break;
            default: throw new ArgumentException("Writing ." + extensionOf(sPath) + " is not supported. Supported: .inix, .csv, .tsv, .md, .xlsx");
        }
    }

    static string extensionOf(string sPath)
    {
        return Path.GetExtension(sPath ?? "").TrimStart('.').ToLowerInvariant();
    }
}

public static class InputHistory
{
    public const int CountCeiling = 100;
    public const int DefaultCount = 10;

    // clampCount: parse a configured history count, falling back to
    // the default and clamping to the ceiling.
    public static int clampCount(string sConfigured)
    {
        int iValue;
        if (!int.TryParse((sConfigured ?? "").Trim(), out iValue)) return DefaultCount;
        if (iValue < 1) return DefaultCount;
        if (iValue > CountCeiling) return CountCeiling;
        return iValue;
    }

    // load: read slots newest-first until the first empty one.
    public static List<string> load(Func<string, string> fnReadKey, int iMax)
    {
        List<string> lsItems = new List<string>();
        if (fnReadKey == null) return lsItems;
        if (iMax < 1) iMax = DefaultCount;
        for (int i = 1; i <= iMax; i++)
        {
            string sTerm = null;
            try { sTerm = fnReadKey("term" + i); } catch { }
            if (string.IsNullOrEmpty(sTerm)) break;
            lsItems.Add(sTerm);
        }
        return lsItems;
    }

    // push: move sNew to the front, dropping any case-insensitive
    // duplicate, and truncate to iMax.
    public static List<string> push(List<string> lsItems, string sNew, int iMax)
    {
        if (lsItems == null) lsItems = new List<string>();
        if (string.IsNullOrEmpty(sNew)) return lsItems;
        if (iMax < 1) iMax = DefaultCount;
        lsItems.RemoveAll(delegate(string sOne) { return string.Equals(sOne, sNew, StringComparison.OrdinalIgnoreCase); });
        lsItems.Insert(0, sNew);
        if (lsItems.Count > iMax) lsItems.RemoveRange(iMax, lsItems.Count - iMax);
        return lsItems;
    }

    // store: write every slot 1..iMax, blank-padding beyond the list
    // so stale entries from a longer history are cleared.
    public static void store(List<string> lsItems, Action<string, string> fnWriteKey, int iMax)
    {
        if (fnWriteKey == null) return;
        if (iMax < 1) iMax = DefaultCount;
        for (int i = 1; i <= iMax; i++)
        {
            string sTerm = (lsItems != null && i <= lsItems.Count) ? (lsItems[i - 1] ?? "") : "";
            try { fnWriteKey("term" + i, sTerm); } catch { }
        }
    }
}

} // namespace Homer
