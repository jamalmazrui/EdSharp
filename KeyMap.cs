// KeyMap.cs -- portable Homer class. Use after `using Homer;`.
//
// The one authoritative table associating a UI context, a command name, a
// one-line summary, an optional longer description, and a key chord. Menu
// building, the Alternate Menu (Alt+F10), the Key Describer (Control+F1), and
// the menu status bar all read from here, so a command is described once and
// every surface agrees. A command with no context is treated as "Global".
//
// Portable by design: depends only on .NET plus the WinForms Keys enum and
// ToolStripMenuItem. To reuse in DbDuo or another project, copy this file and
// keep the `namespace Homer` line. No COM registration, no strong name, no GAC
// -- the deliberate inverse of the old HomerJax packaging.
//
// Camel Type throughout: Hungarian-prefixed fields, lowerCamelCase methods,
// c_-prefixed constants, functions rather than subprocedures.

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Homer {

public static class KeyMap {

    public const string sGlobalContext = "Global";

    // command name -> key chord (Keys.None when unbound).
    public static Dictionary<string, Keys> dCommandToKey =
        new Dictionary<string, Keys>(StringComparer.OrdinalIgnoreCase);
    // key chord -> the menu item that owns it (first binding wins).
    public static Dictionary<Keys, ToolStripMenuItem> dKeyToMenu =
        new Dictionary<Keys, ToolStripMenuItem>();
    // menu item -> command name.
    public static Dictionary<ToolStripMenuItem, string> dMenuToCommand =
        new Dictionary<ToolStripMenuItem, string>();
    // command name -> one-line summary, the form spoken in the status bar,
    // shown inline in the Alternate Menu, and announced by the Key Describer.
    public static Dictionary<string, string> dCommandToSummary =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    // command name -> optional multi-line description for the Alternate Menu's
    // detail pane. The Key Describer uses only the one-line summary.
    public static Dictionary<string, string> dCommandToDescription =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    // command name -> UI context in which its chord is live ("Global" default).
    public static Dictionary<string, string> dCommandToContext =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    // commands whose chord collided with an earlier binding, for diagnostics.
    public static List<string> lsConflicts = new List<string>();

    // Key Describer mode. When on, a hotkey announces its command, chord, and
    // summary instead of running. Toggled by Control+F1 (Help > Key Describer).
    public static bool bKeyDescriber = false;

    // KeyBinding: one typed row of the table -- a command, the context its
    // chord is live in, its summary and description, and the chord itself
    // (Keys.None when unbound). The structured form the Alternate Menu and any
    // future export (TSV, a generated [Hotkeys] block) can read.
    public class KeyBinding {
        public Keys keyData;
        public string sCommand, sContext, sSummary, sDescription;
        public KeyBinding(string sContextIn, string sCommandIn, string sSummaryIn, string sDescriptionIn, Keys keyDataIn) {
            keyData = keyDataIn;
            sCommand = sCommandIn;
            sContext = string.IsNullOrEmpty(sContextIn) ? sGlobalContext : sContextIn;
            sSummary = sSummaryIn == null ? "" : sSummaryIn;
            sDescription = sDescriptionIn == null ? "" : sDescriptionIn;
        }
        public bool bUnbound { get { return keyData == Keys.None; } }
    }

    // lsExtraBindings: bindings that do not come from a menu item -- chords live
    // only inside a dialog or text control. Registered via addBinding and
    // merged in by bindings(). Empty until such contexts are catalogued.
    public static List<KeyBinding> lsExtraBindings = new List<KeyBinding>();

    // register: called as each menu item is built. Additive and idempotent;
    // the first binding of a chord wins, later collisions are recorded.
    public static void register(string sCommand, Keys keyData, ToolStripMenuItem menuItem, string sContext) {
        if (string.IsNullOrEmpty(sCommand)) return;
        dCommandToKey[sCommand] = keyData;
        if (menuItem != null) dMenuToCommand[menuItem] = sCommand;
        if (menuItem != null && keyData != Keys.None) {
            if (dKeyToMenu.ContainsKey(keyData)) { if (!lsConflicts.Contains(sCommand)) lsConflicts.Add(sCommand); }
            else dKeyToMenu[keyData] = menuItem;
        }
        dCommandToContext[sCommand] = string.IsNullOrEmpty(sContext) ? sGlobalContext : sContext;
    }

    public static void addBinding(KeyBinding binding) { if (binding != null) lsExtraBindings.Add(binding); }

    public static void setSummary(string sCommand, string sSummary) {
        if (string.IsNullOrEmpty(sCommand)) return;
        dCommandToSummary[sCommand] = sSummary == null ? "" : sSummary;
    }

    public static string getSummary(string sCommand) {
        string sValue;
        if (sCommand != null && dCommandToSummary.TryGetValue(sCommand, out sValue)) return sValue;
        return "";
    }

    public static void setDescription(string sCommand, string sDescription) {
        if (string.IsNullOrEmpty(sCommand)) return;
        dCommandToDescription[sCommand] = sDescription == null ? "" : sDescription;
    }

    public static string getDescription(string sCommand) {
        string sValue;
        if (sCommand != null && dCommandToDescription.TryGetValue(sCommand, out sValue)) return sValue;
        return "";
    }

    public static string getContext(string sCommand) {
        string sValue;
        if (sCommand != null && dCommandToContext.TryGetValue(sCommand, out sValue)) return sValue;
        return sGlobalContext;
    }

    public static Keys getKey(string sCommand) {
        Keys keyData;
        if (sCommand != null && dCommandToKey.TryGetValue(sCommand, out keyData)) return keyData;
        return Keys.None;
    }

    public static string commandForKey(Keys keyData) {
        ToolStripMenuItem menuItem;
        if (dKeyToMenu.TryGetValue(keyData, out menuItem)) {
            string sCommand;
            if (dMenuToCommand.TryGetValue(menuItem, out sCommand)) return sCommand;
        }
        return "";
    }

    // bindings: every command projected to a typed KeyBinding row, plus the
    // extra (non-menu) bindings. The Alternate Menu reads this.
    public static List<KeyBinding> bindings() {
        List<KeyBinding> lsRows = new List<KeyBinding>();
        foreach (KeyValuePair<string, Keys> pair in dCommandToKey)
            lsRows.Add(new KeyBinding(getContext(pair.Key), pair.Key, getSummary(pair.Key), getDescription(pair.Key), pair.Value));
        foreach (KeyBinding binding in lsExtraBindings) lsRows.Add(binding);
        return lsRows;
    }

    public static List<string> commands() { return new List<string>(dCommandToKey.Keys); }
}

}
