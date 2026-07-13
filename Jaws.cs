// Jaws.cs -- part of the shared Homer toolkit (namespace Homer).
//
// Installs an application's JAWS script family into every installed version of
// JAWS, and compiles it there.  This is the technique DbDo has used: the app owns
// its scripts (in its Scripts folder) and installs them itself, rather than
// shipping a separate script-installer executable.  The installer's Finish page
// simply runs "<App>.exe --install-jaws-settings", so the same code can be re-run
// later from the app's Help menu.
//
// install() takes the folder holding the scripts, so it is app-agnostic: EdSharp,
// FileDir, and DbDo all call the same code.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Homer {


// =====================================================================
// JawsSettingsInstaller: copy DbDo.jkm and DbDo.jss into every
// installed JAWS user-settings folder and run scompile.exe to
// produce DbDo.jsb there. The Pascal-Script equivalent that
// shipped with v1.0.39 worked but lived inside DbDo_setup.iss;
// moving it to C# lets the user re-trigger it later without
// re-running the full installer, and consolidates the JAWS-
// version-discovery logic in one place.
//
// Invoked two ways:
//   - From the installer's [Run] section as
//     `DbDo.exe --install-jaws-settings`, which runs the install
//     and exits without launching the GUI.
//   - From the Help menu's "Install JAWS settings" command, which
//     re-runs the install (for users who upgraded JAWS to a new
//     year-version after installing DbDo).
//
// Returns a multi-line report of what was done. Caller chooses
// whether to show it (the menu version pops a dialog; the CLI
// version prints it).
// =====================================================================
public static class JawsSettingsInstaller
{
    // Find scompile.exe for a given JAWS year-version. Tries the
    // registry value HKLM\Software\Freedom Scientific\JAWS\<ver>\
    // Target first, then falls back to {pf}\Freedom Scientific\
    // JAWS\<ver>\scompile.exe.
    private static string findScompilePath(string sVersion)
    {
        try
        {
            using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine
                .OpenSubKey(@"Software\Freedom Scientific\JAWS\" + sVersion))
            {
                if (key != null)
                {
                    string sTarget = key.GetValue("Target") as string;
                    if (!string.IsNullOrEmpty(sTarget))
                    {
                        string sCompile = System.IO.Path.Combine(sTarget, "scompile.exe");
                        if (System.IO.File.Exists(sCompile)) return sCompile;
                    }
                }
            }
        }
        catch { /* registry access can fail under low privilege; fall through */ }

        string sPf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        string sFallback = System.IO.Path.Combine(sPf,
            @"Freedom Scientific\JAWS\" + sVersion + @"\scompile.exe");
        if (System.IO.File.Exists(sFallback)) return sFallback;
        return null;
    }

    // Run the install. Returns a human-readable report and, via
    // the iCopied / iCompiled out-parameters, totals the caller
    // can use for status text. Records every path placed in a
    // log under %APPDATA%\DbDo\jawsSettings.log so the matching
    // uninstall path can remove exactly those files.
    public static string install(string sAppFolder, out int iCopied, out int iCompiled)
    {
        iCopied = 0;
        iCompiled = 0;
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        System.Collections.Generic.List<string> lLog = new System.Collections.Generic.List<string>();

        string sJawsRoot = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            @"Freedom Scientific\JAWS");
        if (!System.IO.Directory.Exists(sJawsRoot))
        {
            sb.AppendLine("JAWS does not appear to be installed for the current user.");
            sb.AppendLine("(No folder at " + sJawsRoot + ")");
            return sb.ToString();
        }

        string sJkmSource = System.IO.Path.Combine(sAppFolder, "DbDo.jkm");
        string sJssSource = System.IO.Path.Combine(sAppFolder, "DbDo.jss");
        if (!System.IO.File.Exists(sJkmSource))
        {
            sb.AppendLine("DbDo.jkm not found in " + sAppFolder + ".");
            return sb.ToString();
        }
        if (!System.IO.File.Exists(sJssSource))
        {
            sb.AppendLine("DbDo.jss not found in " + sAppFolder + ".");
            return sb.ToString();
        }

        foreach (string sVersionPath in System.IO.Directory.GetDirectories(sJawsRoot))
        {
            string sVersion = System.IO.Path.GetFileName(sVersionPath);
            string sSettingsPath = System.IO.Path.Combine(sVersionPath, "Settings");
            if (!System.IO.Directory.Exists(sSettingsPath)) continue;
            string sScompile = findScompilePath(sVersion);

            foreach (string sLangPath in System.IO.Directory.GetDirectories(sSettingsPath))
            {
                string sLang = System.IO.Path.GetFileName(sLangPath);
                string sJkmTarget = System.IO.Path.Combine(sLangPath, "DbDo.jkm");
                string sJssTarget = System.IO.Path.Combine(sLangPath, "DbDo.jss");
                string sJsbTarget = System.IO.Path.Combine(sLangPath, "DbDo.jsb");

                bool bJkmOk = false, bJssOk = false, bJsbOk = false;
                try { System.IO.File.Copy(sJkmSource, sJkmTarget, true); bJkmOk = true; iCopied++; lLog.Add(sJkmTarget); }
                catch (Exception ex) { sb.AppendLine("FAIL: copy jkm to " + sJkmTarget + ": " + ex.Message); }
                try { System.IO.File.Copy(sJssSource, sJssTarget, true); bJssOk = true; iCopied++; lLog.Add(sJssTarget); }
                catch (Exception ex) { sb.AppendLine("FAIL: copy jss to " + sJssTarget + ": " + ex.Message); }

                if (bJssOk && !string.IsNullOrEmpty(sScompile))
                {
                    try
                    {
                        System.Diagnostics.ProcessStartInfo psi =
                            new System.Diagnostics.ProcessStartInfo(sScompile, "DbDo.jss");
                        psi.WorkingDirectory = sLangPath;
                        psi.UseShellExecute = false;
                        psi.CreateNoWindow = true;
                        psi.RedirectStandardOutput = true;
                        psi.RedirectStandardError = true;
                        using (System.Diagnostics.Process proc = System.Diagnostics.Process.Start(psi))
                        {
                            proc.WaitForExit(10000);
                            if (proc.HasExited && proc.ExitCode == 0 && System.IO.File.Exists(sJsbTarget))
                            { bJsbOk = true; iCompiled++; lLog.Add(sJsbTarget); }
                            else
                            {
                                string sErr = proc.HasExited ? proc.StandardError.ReadToEnd() : "(timed out)";
                                sb.AppendLine("FAIL: compile " + sJsbTarget + " - " + sErr.Trim());
                            }
                        }
                    }
                    catch (Exception ex) { sb.AppendLine("FAIL: compile " + sJsbTarget + ": " + ex.Message); }
                }
                else if (bJssOk && string.IsNullOrEmpty(sScompile))
                {
                    sb.AppendLine("WARN: scompile.exe not found for JAWS " + sVersion
                        + "; placed jss but did not compile (run scompile manually).");
                }

                sb.AppendLine("JAWS " + sVersion + " / " + sLang + ": "
                    + (bJkmOk ? "jkm " : "no-jkm ")
                    + (bJssOk ? "jss " : "no-jss ")
                    + (bJsbOk ? "jsb" : "no-jsb"));
            }
        }

        // Persist the log so the matching uninstall step can
        // remove exactly the files we placed.
        try
        {
            string sLogPath = getLogPath();
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(sLogPath));
            System.IO.File.WriteAllLines(sLogPath, lLog);
        }
        catch (Exception ex) { sb.AppendLine("WARN: could not write log: " + ex.Message); }

        return sb.ToString();
    }

    // Uninstall: read the log, delete each path listed, then
    // delete the log itself. Mirrors what the Pascal Script
    // CurUninstallStepChanged did in v1.0.39.
    public static string uninstall(out int iDeleted)
    {
        iDeleted = 0;
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        string sLogPath = getLogPath();
        if (!System.IO.File.Exists(sLogPath))
        {
            sb.AppendLine("No jawsSettings.log found; nothing to remove.");
            return sb.ToString();
        }
        try
        {
            string[] aPaths = System.IO.File.ReadAllLines(sLogPath);
            foreach (string sPath in aPaths)
            {
                if (string.IsNullOrWhiteSpace(sPath)) continue;
                try
                {
                    if (System.IO.File.Exists(sPath))
                    {
                        System.IO.File.Delete(sPath);
                        iDeleted++;
                        sb.AppendLine("removed " + sPath);
                    }
                }
                catch (Exception ex) { sb.AppendLine("FAIL: " + sPath + ": " + ex.Message); }
            }
            try { System.IO.File.Delete(sLogPath); } catch { /* tolerate */ }
            try { System.IO.Directory.Delete(System.IO.Path.GetDirectoryName(sLogPath), false); }
            catch { /* only succeeds if empty; ok either way */ }
        }
        catch (Exception ex) { sb.AppendLine("FAIL: read log: " + ex.Message); }
        return sb.ToString();
    }

    private static string getLogPath()
    {
        return System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            @"DbDo\jawsSettings.log");
    }
}

} // namespace Homer (Jaws.cs)
