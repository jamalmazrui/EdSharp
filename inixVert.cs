// inixVert.cs -- the Inix table converter for the Convert pipeline.
//
// USAGE:
//   inixVert <source-file> <dest-file> [/quiet]
//
// The source and destination formats come from the file extensions.
// Supported on either side: .inix, .csv, .tsv, .md, .xlsx -- any
// direction, with .inix as the screen-reader-friendly home format
// (one [RecordNNN] section per row, one field = value line per field).
// All the real work is Homer.InixTable in the shared Inix.cs; this file
// is the thin command-line wrapper EdSharp's Import and Export tables
// call.
//
// SECOND MODE: when the source is .mdx or .md and the destination is
// .md or .txt, the embedded-inix EXPANSION runs instead: fenced
// "inix" code blocks become real Markdown tables (grid tables when a
// cell is multi-line) and everything else passes through unchanged.
// The mdx2htm and mdx2docx batch files chain this with pandoc.
//
// EXIT CODES:  0 success   3 the conversion failed   4 bad arguments
//
// A detailed log is written to inixVert.log beside this executable:
// environment, arguments, and any error in full. A failure never ends
// with only a console message and an empty log.
//
// Build: BuildEdSharp compiles this with Inix.cs into
// Convert\inixVert.exe (references System.IO.Compression.dll and
// System.Xml.dll).

using System;
using System.IO;
using System.Reflection;
using System.Text;

public static class InixVert
{
    static StreamWriter fileLog;
    static bool bQuiet;

    static void say(string sMessage)
    {
        if (!bQuiet) Console.WriteLine(sMessage);
        if (fileLog != null) { fileLog.WriteLine(sMessage); fileLog.Flush(); }
    }

    public static int Main(string[] aArguments)
    {
        string sExeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        string sLogPath = Path.Combine(sExeDir, "inixVert.log");
        try { fileLog = new StreamWriter(sLogPath, false, new UTF8Encoding(true)); }
        catch (Exception) { fileLog = null; }
        try
        {
            foreach (string sArgument in aArguments)
                if (string.Equals(sArgument, "/quiet", StringComparison.OrdinalIgnoreCase)) bQuiet = true;
            say("inixVert  " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            say("  executable:        " + Assembly.GetExecutingAssembly().Location);
            say("  working directory: " + Environment.CurrentDirectory);
            say("  command line:      " + Environment.CommandLine);
            say("");
            int iPlain = 0;
            string sSource = null, sDest = null;
            foreach (string sArgument in aArguments)
            {
                if (sArgument.StartsWith("/")) continue;
                if (iPlain == 0) sSource = sArgument;
                else if (iPlain == 1) sDest = sArgument;
                iPlain++;
            }
            if (sSource == null || sDest == null || iPlain != 2)
            {
                say("USAGE: inixVert <source-file> <dest-file> [/quiet]");
                say("Formats come from the extensions: .inix, .csv, .tsv, .md, .xlsx.");
                return 4;
            }
            if (!File.Exists(sSource))
            {
                say("FAILED: the source file does not exist: " + sSource);
                return 3;
            }
            string sSourceExt = Path.GetExtension(sSource).TrimStart('.').ToLowerInvariant();
            string sDestExt = Path.GetExtension(sDest).TrimStart('.').ToLowerInvariant();
            bool bSourceMarkdown = (sSourceExt == "mdx" || sSourceExt == "md" || sSourceExt == "markdown");
            bool bDestMarkdown = (sDestExt == "md" || sDestExt == "markdown" || sDestExt == "txt");
            if (bSourceMarkdown && bDestMarkdown)
            {
                say("Expanding embedded inix tables:  " + sSource);
                Homer.InixTable.expandMarkdownFile(sSource, sDest);
                say("Written  " + sDest);
                say("Done.");
                return 0;
            }
            say("Reading  " + sSource);
            Homer.InixTable.TableData table = Homer.InixTable.readAny(sSource);
            string sFields = (table.Fields.Count == 1) ? "field" : "fields";
            string sRows = (table.Rows.Count == 1) ? "row" : "rows";
            say("  " + table.Fields.Count + " " + sFields + ", " + table.Rows.Count + " " + sRows);
            say("Writing  " + sDest);
            Homer.InixTable.writeAny(sDest, table);
            say("Done.");
            return 0;
        }
        catch (Exception exception)
        {
            say("FAILED: " + exception.Message);
            say(exception.ToString());
            return 3;
        }
        finally
        {
            if (fileLog != null) fileLog.Close();
        }
    }
}
