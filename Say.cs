// Say.cs (speech subsystem: Say + UIA notification support) -- portable, reusable across EdSharp, DbDuo, and other C#
// projects. Say.sayForced dispatches JAWS COM -> NVDA controller client -> native UIA notification. attach(Form) once at startup. No app dependencies.
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

public static class Say
{
    private static Label lbl;
    // Form reference saved at attach time so the pure-UIA path can
    // raise notifications from the form's AccessibleObject rather
    // than the 1x1 Label's. Narrator is particular about
    // notification source: it tends to honor notifications from
    // a top-level window's AccessibleObject and ignore them from
    // marginal hidden controls. JAWS and NVDA are more permissive
    // and pick up either source.
    private static Form frmOwner;

    // Reflection cache for AccessibleObject.RaiseAutomationNotification
    // was removed in v1.0.87 when the UIA path switched to native
    // P/Invoke against UIAutomationCore.dll. See dispatchNativeUia
    // Notification and the NotificationHostControl / AnnouncerProvider
    // classes below.

    // Attach the live region to a form. Called once during form
    // construction. The Label is a real on-screen control with
    // LiveSetting=Polite; WinForms automatically raises the UIA
    // LiveRegionChanged event when the Label's Text changes,
    // which JAWS, NVDA, and Narrator all listen for.
    //
    // Placement: 1x1 at the form origin (0, 0). The Label sits
    // tucked under the MenuStrip so it is visually unnoticeable
    // to sighted users, but stays inside the form's client area,
    // which guarantees it participates in the UIA tree. (A
    // far-off-screen Location can sometimes be pruned from the
    // tree by WinForms.)
    public static void attach(Form frm)
    {
        if (lbl != null) return;
        frmOwner = frm;
        lbl = new Label();
        lbl.AutoSize = false;
        lbl.Size = new Size(1, 1);
        lbl.Location = new Point(0, 0);
        lbl.TabStop = false;
        lbl.Visible = true;
        lbl.Text = "";
        lbl.AccessibleName = "";
        lbl.AccessibleRole = AccessibleRole.StaticText;
        // Live region turned OFF.  The announcement is made solely by the UIA
        // Notification event (see sayViaUia).  Leaving this Assertive meant any
        // change to the label's Text raised LiveRegionChanged as well, so one
        // message produced two UIA events and a UIA-listening reader spoke it
        // twice.  The label is kept only as the element the notification is
        // raised against.
        lbl.LiveSetting = System.Windows.Forms.Automation.AutomationLiveSetting.Off;
        frm.Controls.Add(lbl);
        // Force the handle to be created so the AccessibleObject
        // is wired up before the first say() call. Without this,
        // the very first announcement can race ahead of UIA tree
        // initialization and be missed.
        IntPtr hForce = lbl.Handle;
    }

    // Push a string to the live region. The Label's Text change
    // auto-raises the UIA LiveRegionChanged event (because
    // LiveSetting was set to Polite at attach time), and we also
    // fire the UIA Notification event by reflection as a belt-
    // and-suspenders fallback. If the live region was never
    // attached (e.g., CLI-only mode), the call is silently
    // ignored.
    //
    // The say() pipeline. Determines which screen reader is
    // currently running and speaks through that reader's own
    // API so the speech uses the user's configured voice and
    // verbosity. No SAPI fallback -- if no screen reader is
    // running, no speech happens. Per-reader paths:
    //
    //   1. JAWS:     FreedomSci.JawsApi COM object's SayString.
    //                Detection: FindWindow("JFWUI2"), which is
    //                JAWS's top-level UI window class. Avoids the
    //                cost of creating the COM object just to test.
    //
    //   2. NVDA:     nvdaControllerClient.dll P/Invoke. The DLL
    //                must be shipped alongside DbDo.exe. NVDA
    //                does NOT expose a COM API; the controller
    //                client is the documented IPC channel.
    //                Detection: nvdaController_testIfRunning()
    //                returns 0 when NVDA is running.
    //
    //   3. Narrator: UIA Notification event via
    //                AccessibleObject.RaiseAutomationNotification.
    //                Narrator does not have a per-app API; it
    //                listens to UIA events. RaiseAutomationNotification
    //                is the documented "announce this text now"
    //                event. Also reaches NVDA and JAWS in their
    //                UIA-enabled modes, but we already covered
    //                those by direct API, so this is effectively
    //                the Narrator path.
    //                Detection: SystemParametersInfo SPI_GETSCREENREADER
    //                returns true when any screen reader is
    //                running. We use this only as a "does anything
    //                care?" hint -- the Notification event is
    //                cheap to fire whether or not Narrator is on.
    //
    // The Label-based live-region path that previous versions
    // used (Label.LiveSetting=Assertive) is preserved as part of
    // the Narrator path because some configurations of NVDA and
    // JAWS in UIA-only mode listen for LiveRegionChanged in
    // addition to / instead of their direct APIs.
    //
    // Priority order: if JAWS is detected, JAWS speaks and we
    // stop (we don't want JAWS to also receive a UIA Notification
    // and speak twice). Same for NVDA. Otherwise UIA Notification
    // fires for Narrator (and any other UIA-listening reader).
    public static void say(string sText)
    {
        // Extra-Speech gate: when off, DbDo's direct speech is
        // suppressed but the screen reader's natural focus and
        // selection announcements still occur. The flag is toggled
        // by Toggle-Extra-Speech (Alt+Shift+S) and persisted to
        // [General] extraSpeech in DbDo.inix. The toggle command
        // itself uses sayForced so the user always hears their
        // own action confirmed regardless of the flag's state.
        if (!bExtraSpeechEnabled) return;
        sayForced(sText);
    }

    // sayForced: bypass the Extra-Speech gate. Used for two
    // narrow purposes: the Toggle-Extra-Speech command's own
    // confirmation (so the user hears "extra speech off" when
    // turning it off), and the Test-Reader command (which must
    // produce speech to be diagnostic).
    public static void sayForced(string sText)
    {
        string sNew = sText ?? "";
        if (isJawsRunning() && jawsSay(sNew)) { sLastPath = "JAWS COM"; return; }
        if (isNvdaRunning() && nvdaSay(sNew)) { sLastPath = "NVDA controller client"; return; }
        // Fall through to the UIA path for Narrator and any
        // unrecognized UIA-aware reader. Also useful as a debug
        // signal: if you have a screen reader running that's
        // detectable via SPI_GETSCREENREADER but not JAWS or
        // NVDA, this is the path that reaches it.
        sayViaUia(sNew);
        sLastPath = "UIA Notification + LiveRegionChanged (Narrator and others)";
        // Optional SAPI backup (off by default; see bUseSapiAsBackup). Speaks
        // only when no screen reader is detected, so it never doubles a live
        // reader and gives audible output where a bare UIA Notification would
        // be silent. This is SayIt's UseSAPIAsBackup behavior, modernized.
        if (bUseSapiAsBackup && !isJawsRunning() && !isNvdaRunning() && !isAnyScreenReaderActive())
            if (sapiSay(sNew)) sLastPath = "SAPI backup (no screen reader detected)";
    }

    // sayParts: speak each unit as its OWN direct-speech message
    // rather than concatenating into one run-on string. Both JAWS
    // (SayString flush=false) and the NVDA controller client queue
    // successive messages, so the units come out in order with a
    // natural pause between them, which is markedly more
    // intelligible. Each unit passes through the same Extra-Speech
    // gate as say(); empty units are skipped so a missing piece
    // doesn't leave a dead pause. This is the preferred way to
    // voice any announcement built of logical units -- a cell's
    // header / row / value, a list of columns, a column's values.
    public static void sayParts(IEnumerable<string> parts)
    {
        if (parts == null) return;
        foreach (string sPart in parts)
            if (!string.IsNullOrEmpty(sPart)) say(sPart);
    }

    public static void sayParts(params string[] aParts)
    {
        sayParts((IEnumerable<string>)aParts);
    }

    // Extra-Speech enabled flag. Public mutable so the toggle
    // command can flip it. Loaded from DbDo.inix at startup;
    // default true. Persisted on change.
    public static bool bExtraSpeechEnabled = true;

    // Optional SAPI text-to-speech backup, modeled on SayIt's UseSAPIAsBackup
    // property. Off by default, so the JAWS -> NVDA -> UIA cascade is exactly
    // as before. When true, sayForced falls back to SAPI.SPVoice only when no
    // screen reader is detected at all -- a guaranteed-audio last resort.
    public static bool bUseSapiAsBackup = false;

    // Records which path the most recent say() call used. The
    // Test-Reader command displays this so the user can confirm
    // whether they are hearing speech through the direct JAWS
    // COM call, the direct NVDA controller client, or the
    // generic UIA Notification fallback.
    private static string sLastPath = "(none yet)";
    public static string lastSpeechPath()
    {
        return sLastPath;
    }

    // Diagnostic snapshot of the speech pipeline. Returns a
    // multi-line string covering which readers are detected,
    // which DLL is loadable, and which path the most recent
    // say() used. Test-Reader presents this in a MessageBox.
    public static string speechDiagnostic()
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("Speech pipeline diagnostic");
        sb.AppendLine();
        sb.AppendLine("JAWS running (window class JFWUI2 found): " + (isJawsRunning() ? "yes" : "no"));
        sb.AppendLine("JAWS COM ProgID FreedomSci.JawsApi reachable: "
            + ((oJawsApi != null) ? "yes (cached)" : "(not yet probed or unavailable)"));
        sb.AppendLine();
        bool bNvdaDll = false;
        try { nvdaController_testIfRunning(); bNvdaDll = true; }
        catch (DllNotFoundException) { bNvdaDll = false; }
        catch { bNvdaDll = true; }
        sb.AppendLine("nvdaControllerClient.dll loadable: " + (bNvdaDll ? "yes" : "no (drop the DLL next to EdSharp.exe to enable NVDA support)"));
        sb.AppendLine("NVDA running (controller client says so): " + (isNvdaRunning() ? "yes" : "no"));
        sb.AppendLine();
        // SystemParametersInfo SPI_GETSCREENREADER: a generic
        // "some screen reader is on" probe. Independent of JAWS
        // or NVDA detection.
        sb.AppendLine("Windows reports any screen reader active: " + (isAnyScreenReaderActive() ? "yes" : "no"));
        sb.AppendLine();
        sb.AppendLine("Most recent say() used path: " + sLastPath);
        return sb.ToString();
    }

    // SystemParametersInfo SPI_GETSCREENREADER (action 70).
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern bool SystemParametersInfo(int iAction, int iParam, ref bool bResult, int iUpdate);
    private const int SPI_GETSCREENREADER = 70;
    private static bool isAnyScreenReaderActive()
    {
        try
        {
            bool bActive = false;
            if (SystemParametersInfo(SPI_GETSCREENREADER, 0, ref bActive, 0))
                return bActive;
        }
        catch { }
        return false;
    }

    // ---- JAWS: detection + speak ----
    // JAWS exposes a top-level UI window of class "JFWUI2". This
    // is the cheap detection path -- FindWindow does not require
    // COM startup. (Creating FreedomSci.JawsApi when JAWS is not
    // running is slow and produces a misleading "succeeded but
    // nobody listening" object.)
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr FindWindow(string sClass, string sTitle);
    private static bool isJawsRunning()
    {
        try { return FindWindow("JFWUI2", null) != IntPtr.Zero; }
        catch { return false; }
    }

    // Speak through JAWS. The COM object is cached after first
    // successful creation. If COM creation succeeds, we re-use
    // the object across calls; if SayString throws (JAWS
    // crashed or was closed mid-session), reset and let the
    // next call try again.
    private static object oJawsApi;
    private static object oSapiVoice;
    private static bool jawsSay(string sText)
    {
        if (string.IsNullOrEmpty(sText)) return false;
        try
        {
            if (oJawsApi == null)
            {
                Type comType = Type.GetTypeFromProgID("FreedomSci.JawsApi");
                if (comType == null) return false;
                oJawsApi = Activator.CreateInstance(comType);
                if (oJawsApi == null) return false;
            }
            dynamic oJaws = oJawsApi;
            // SayString(text, flush). flush=false queues politely
            // rather than interrupting whatever JAWS is reading.
            object result = oJaws.SayString(sText, false);
            return result != null && (bool)result;
        }
        catch
        {
            oJawsApi = null;
            return false;
        }
    }

    // ---- SAPI: guaranteed text-to-speech backup (optional) ----
    // Mirrors SayIt's SAPISay: late-bound SAPI.SPVoice, spoken
    // asynchronously (SVSFlagsAsync = 1) so it never blocks the UI
    // thread. Reached only when bUseSapiAsBackup is on and no screen
    // reader is present.
    private static bool sapiSay(string sText)
    {
        if (string.IsNullOrEmpty(sText)) return false;
        try
        {
            if (oSapiVoice == null)
            {
                Type comType = Type.GetTypeFromProgID("SAPI.SPVoice");
                if (comType == null) return false;
                oSapiVoice = Activator.CreateInstance(comType);
                if (oSapiVoice == null) return false;
            }
            dynamic oVoice = oSapiVoice;
            oVoice.Speak(sText, 1);
            return true;
        }
        catch
        {
            oSapiVoice = null;
            return false;
        }
    }

    // ---- NVDA: detection + speak ----
    // NVDA ships a Controller Client DLL with C-style exports.
    // Starting in NVDA 2026.1 the DLL is named simply
    // nvdaControllerClient.dll (no architecture suffix); earlier
    // releases shipped it as nvdaControllerClient64.dll for x64
    // hosts and nvdaControllerClient32.dll for x86. NVDA's own
    // current C# example uses the unsuffixed name and that is
    // what the DbDo build script downloads and bundles. Place
    // the DLL next to DbDo.exe and the DllImport finds it via
    // the standard Windows DLL search order. If the DLL is
    // missing, DllNotFoundException is caught silently and the
    // NVDA path is unavailable.
    //
    // testIfRunning returns 0 when NVDA is running, non-zero
    // otherwise (it's a Windows error code). speakText takes
    // a wide-char string and returns 0 on success.
    [DllImport("nvdaControllerClient.dll", CharSet = CharSet.Unicode, EntryPoint = "nvdaController_testIfRunning")]
    private static extern int nvdaController_testIfRunning();
    [DllImport("nvdaControllerClient.dll", CharSet = CharSet.Unicode, EntryPoint = "nvdaController_speakText")]
    private static extern int nvdaController_speakText([MarshalAs(UnmanagedType.LPWStr)] string sText);
    [DllImport("nvdaControllerClient.dll", CharSet = CharSet.Unicode, EntryPoint = "nvdaController_cancelSpeech")]
    private static extern int nvdaController_cancelSpeech();

    private static bool bNvdaProbed;
    private static bool bNvdaDllPresent;
    private static bool isNvdaRunning()
    {
        // First call: probe whether the DLL is loadable at all.
        // Subsequent calls skip the probe and just call
        // testIfRunning, which is cheap and reliable.
        if (!bNvdaProbed)
        {
            bNvdaProbed = true;
            try
            {
                int iResult = nvdaController_testIfRunning();
                bNvdaDllPresent = true;
                return iResult == 0;
            }
            catch (DllNotFoundException)
            {
                bNvdaDllPresent = false;
                return false;
            }
            catch { bNvdaDllPresent = false; return false; }
        }
        if (!bNvdaDllPresent) return false;
        try { return nvdaController_testIfRunning() == 0; }
        catch { return false; }
    }

    private static bool nvdaSay(string sText)
    {
        if (string.IsNullOrEmpty(sText)) return false;
        if (!bNvdaDllPresent) return false;
        try { return nvdaController_speakText(sText) == 0; }
        catch { return false; }
    }

    // ---- Narrator / generic UIA path ----
    // Fire the UIA Notification event, and ONLY that event.  This is the path
    // used when neither JAWS nor NVDA is running, so it must not double-speak:
    // exactly one announcement is raised per message.
    private static void sayViaUia(string sText)
    {
        if (lbl == null) return;
        if (lbl.IsDisposed) return;
        try
        {
            if (lbl.InvokeRequired)
            {
                lbl.Invoke(new Action<string>(sayViaUia), new object[] { sText });
                return;
            }
            // Raise the UIA Notification event ONLY.  Writing the label's Text used
            // to be the announcement mechanism, but the label is an ASSERTIVE live
            // region, so that write also raises LiveRegionChanged: two UIA events
            // for one message, which a UIA-listening reader speaks TWICE.  The
            // native UIA Notification reaches Narrator (and any other UIA reader) on
            // its own, so the live-region write is gone.  The label itself stays --
            // it is the element the notification is raised against -- but its text is
            // never changed now, so it raises no event of its own.
            raiseUiaNotification(sText);
        }
        catch { }
    }

    // raiseUiaNotification: legacy entry point used by sayViaUia
    // (which is invoked from say() as a Narrator-only fallback when
    // neither JAWS COM nor NVDA controller-client is reachable).
    // The historical reflection-on-Label path was a documented
    // dead end on .NET Framework 4.8: only Label / LinkLabel /
    // GroupBox / ProgressBar honor AccessibleObject.RaiseAutomation
    // Notification on that runtime, and even then the dispatch did
    // not reach Narrator reliably on Windows 11.
    //
    // v1.0.87 replaces the body with a native-P/Invoke call to
    // UiaRaiseNotificationEvent against a fresh per-call custom
    // IRawElementProviderSimple. That technique reaches all three
    // readers (verified by ChatGPT's reference sample, which the
    // user confirmed worked end to end). The reflection path is
    // gone; the method signature stays so sayViaUia's call site
    // doesn't have to change.
    private static void raiseUiaNotification(string sText)
    {
        dispatchNativeUiaNotification(sText, /*All*/ 2);
    }

    // sayUiaString: pure-UIA speech path. Fires the UIA
    // RaiseAutomationNotification event directly, without going
    // through JAWS COM, NVDA controller client, or the Label/
    // LiveRegionChanged intermediary. Use for testing the UIA
    // path in isolation or when the caller wants a single,
    // simple announcement that doesn't depend on JAWS or NVDA
    // being detected.
    //
    // Reaches JAWS, NVDA, and Narrator in their UIA-listening
    // modes (which is their default). Does NOT go through
    // JAWS's FreedomSci.JawsApi.SayString or NVDA's
    // nvdaControllerClient.dll, so this path is independent of
    // any per-screen-reader API.
    //
    // Inspired by the WPF approach in Kelly Ford's UIANotifications
    // demo (https://github.com/kellylford/TheWorkBench/tree/main/UiaNotifyDemo).
    // The WinForms equivalent uses AccessibleObject's
    // RaiseAutomationNotification instead of WPF's
    // UIElementAutomationPeer.RaiseNotificationEvent, but the
    // underlying UIA event is the same.
    //
    // Source element: the form's own AccessibleObject (not the
    // hidden 1x1 Label that the legacy say() path uses). Narrator
    // honors notifications from top-level windows but tends to
    // ignore them from marginal hidden controls. JAWS and NVDA
    // accept either. Processing mode is "All" (value 2), matching
    // the behavior of the legacy sayViaUia path that DbDo's
    // Narrator fallback was already using successfully -- "All"
    // means deliver to every listener without queue restrictions.
    // sayUiaString: pure-UIA speech path. Fires the UIA
    // notification event directly via native UIAutomationCore.dll,
    // bypassing JAWS COM, NVDA controller client, and the legacy
    // reflection-on-AccessibleObject technique that previous
    // versions (v1.0.76 - v1.0.86) used. The new path is adapted
    // from ChatGPT's verified reference sample which reaches
    // JAWS, NVDA, and Narrator on the same machine.
    //
    // Why it works where the managed RaiseAutomationNotification
    // failed: on .NET Framework 4.8, AccessibleObject.RaiseAutomation
    // Notification is honored only by Label / LinkLabel / GroupBox /
    // ProgressBar, and even those reach Narrator inconsistently.
    // The native UiaRaiseNotificationEvent (in UIAutomationCore.dll)
    // has no such restriction; given a fully-implemented
    // IRawElementProviderSimple anchored to a real window through
    // WM_GETOBJECT, the screen reader's UIA listener picks the
    // event up directly.
    //
    // Reaches JAWS, NVDA, and Narrator in their UIA-listening
    // modes. Does NOT go through JAWS's FreedomSci.JawsApi.SayString
    // or NVDA's nvdaControllerClient.dll.
    public static void sayUiaString(string sText)
    {
        dispatchNativeUiaNotification(sText, /*All*/ 2);
    }

    // sayUiaStringForced: like sayUiaString but uses the
    // ImportantMostRecent processing mode, which screen readers
    // treat as "interrupt current speech and say this now." Use
    // for time-critical announcements (errors, completion of a
    // long operation the user is waiting on).
    public static void sayUiaStringForced(string sText)
    {
        dispatchNativeUiaNotification(sText, /*ImportantMostRecent*/ 1);
    }

    // dispatchNativeUiaNotification: the actual native UIA path.
    // Adapted from ChatGPT-generated reference sample (which the
    // user confirmed works for JAWS, NVDA, and Narrator on Windows
    // 11 with .NET Framework 4.8). Bypasses the managed
    // AccessibleObject.RaiseAutomationNotification API entirely.
    //
    //   1. Create a fresh hidden NotificationHostControl (custom
    //      Control subclass that overrides WndProc to respond to
    //      WM_GETOBJECT by returning a custom IRawElementProvider
    //      Simple). The control is added to the form's Controls
    //      so it has a real window handle; UIA's source-validity
    //      check passes because WM_GETOBJECT returns a provider
    //      genuinely associated with the hwnd.
    //   2. Call UiaRaiseNotificationEvent (P/Invoke into
    //      UIAutomationCore.dll) with the provider as the source.
    //   3. Retain the host on a small ring so it survives long
    //      enough for the screen reader's async UIA listener to
    //      finish reading the Name property. We keep the last 5
    //      hosts before disposing the oldest. The unique sequence
    //      number on each host also prevents UIA from deduping
    //      "same source + same text" pairs.
    //
    // Threading: must run on the UI thread (creating Controls and
    // adding to frmOwner.Controls is a UI-thread operation). If
    // called from a worker thread, marshal via frmOwner.BeginInvoke.
    //
    // Why a fresh host per call (and not a reused pool): ChatGPT's
    // empirical investigation (uia_notify_winforms_repeat_fix
    // sample, "Most Important Behavioral Finding" in its findings
    // doc) determined that reusing a provider for multiple
    // announcements causes screen readers / UIA to silently drop
    // later events, even when the text differs. The dedupe
    // heuristic appears to operate per-source-HWND, so any reuse
    // -- whether a single host or a small pool cycled
    // round-robin -- will eventually drop announcements when two
    // consecutive calls land on the same host. A fresh control
    // per call is the only pattern proven to work for JAWS, NVDA,
    // and Narrator simultaneously.
    //
    // The cost is small in absolute terms: creating a 1x1 Control
    // with an HWND is on the order of tenths of a millisecond and
    // a few KB of managed memory. The retention ring caps the live
    // host count at 5; older hosts are disposed and their HWNDs
    // released. For a database manager where announcements fire
    // on user actions (cell moves, refreshes, edits), this is
    // invisible in normal usage. Do not "optimize" this to a
    // shared host -- it has been tried and it breaks Narrator and
    // NVDA after the first announcement.
    private static System.Collections.Generic.List<NotificationHostControl> lsOldHosts
        = new System.Collections.Generic.List<NotificationHostControl>();
    private static int iNotificationSequence;

    private static void dispatchNativeUiaNotification(string sText, int iProcessing)
    {
        if (string.IsNullOrEmpty(sText)) return;
        if (frmOwner == null || frmOwner.IsDisposed) return;
        if (frmOwner.InvokeRequired)
        {
            try
            {
                frmOwner.BeginInvoke(new Action(delegate()
                {
                    dispatchNativeUiaNotification(sText, iProcessing);
                }));
            }
            catch { /* form may be tearing down; non-fatal */ }
            return;
        }
        try
        {
            iNotificationSequence++;
            NotificationHostControl host = new NotificationHostControl(
                iNotificationSequence, sText);
            host.Left = 0;
            host.Top = 0;
            frmOwner.Controls.Add(host);
            host.CreateControl();  // force hwnd allocation
            lsOldHosts.Add(host);

            // Retain the last 5 hosts. Disposing immediately can
            // cause UIA / the screen reader to lose the source
            // before reading the Name property asynchronously.
            while (lsOldHosts.Count > 5)
            {
                NotificationHostControl old = lsOldHosts[0];
                lsOldHosts.RemoveAt(0);
                try
                {
                    if (frmOwner != null && !frmOwner.IsDisposed)
                        frmOwner.Controls.Remove(old);
                }
                catch { }
                try { old.Dispose(); } catch { }
            }

            // Fully qualify the enum types: both
            // System.Windows.Forms.Automation and
            // System.Windows.Automation define same-named enums
            // (AutomationNotificationKind, AutomationNotification
            // Processing). The UIAutomationCore.dll P/Invoke
            // expects the System.Windows.Automation versions.
            System.Windows.Automation.AutomationNotificationProcessing proc =
                (iProcessing == 1)
                    ? System.Windows.Automation.AutomationNotificationProcessing.ImportantMostRecent
                    : System.Windows.Automation.AutomationNotificationProcessing.All;
            string sActivityId = "EdSharp." + iNotificationSequence.ToString() + "." + Guid.NewGuid().ToString("N");
            UiaNative.UiaRaiseNotificationEvent(
                host.ProviderAnnouncer,
                System.Windows.Automation.AutomationNotificationKind.Other,
                proc,
                sText,
                sActivityId);
        }
        catch { /* native UIA dispatch failures are non-fatal */ }
    }
}

// =====================================================================
// Native UIA dispatch helpers (v1.0.87). Adapted from ChatGPT's
// reference sample that the user confirmed reaches JAWS, NVDA, and
// Narrator simultaneously on Windows 11 with .NET Framework 4.8.
//
// NotificationHostControl is a 1x1 invisible Control that owns a
// genuine HWND. By overriding WndProc to respond to WM_GETOBJECT,
// it returns a custom IRawElementProviderSimple (AnnouncerProvider)
// when UIA queries the window. UIA / the screen reader sees a
// legitimate UIA element in the tree, so when we call
// UiaRaiseNotificationEvent against that provider, the event
// reaches all three readers.
//
// AnnouncerProvider is the IRawElementProviderSimple implementation.
// It exposes a minimal but complete set of UIA properties: Name
// (the message text), AutomationId, ControlType=Text, FrameworkId,
// IsControlElement=true, IsContentElement=true. HostRawElementProvider
// anchors the provider to the host window via
// AutomationInteropProvider.HostProviderFromHandle.
//
// UiaNative.UiaRaiseNotificationEvent is a P/Invoke binding to the
// native function in UIAutomationCore.dll. The managed wrapper
// AccessibleObject.RaiseAutomationNotification on .NET Framework
// 4.8 only works for four specific control types (Label, LinkLabel,
// GroupBox, ProgressBar) and even then doesn't reach Narrator
// reliably; the native function has no such restriction.
// =====================================================================
public class NotificationHostControl : Control
{
    private const int iWmGetObject = 0x003D;
    private AnnouncerProvider providerAnnouncer;
    public int iSequence;
    public string sMessageText;

    public NotificationHostControl(int iSeq, string sMsg)
    {
        this.TabStop = false;
        this.Width = 1;
        this.Height = 1;
        this.Text = "EdSharp notification host " + iSeq.ToString();
        this.AccessibleName = sMsg;
        this.iSequence = iSeq;
        this.sMessageText = sMsg;
        this.providerAnnouncer = null;
    }

    public AnnouncerProvider ProviderAnnouncer
    {
        get
        {
            ensureProvider();
            return providerAnnouncer;
        }
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        ensureProvider();
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == iWmGetObject)
        {
            ensureProvider();
            m.Result = AutomationInteropProvider.ReturnRawElementProvider(
                this.Handle, m.WParam, m.LParam, providerAnnouncer);
            return;
        }
        base.WndProc(ref m);
    }

    private void ensureProvider()
    {
        if (providerAnnouncer == null && this.IsHandleCreated)
        {
            providerAnnouncer = new AnnouncerProvider(this.Handle, iSequence, sMessageText);
        }
    }
}

public class AnnouncerProvider : IRawElementProviderSimple
{
    private const int iAutomationIdProperty   = 30011;
    private const int iControlTypeProperty    = 30003;
    private const int iFrameworkIdProperty    = 30024;
    private const int iIsContentElementProperty = 30017;
    private const int iIsControlElementProperty = 30016;
    private const int iNameProperty           = 30005;
    private const int iTextControlType        = 50020;

    private IntPtr hWindow;
    private int iSequence;
    private string sName;

    public AnnouncerProvider(IntPtr hWnd, int iSeq, string sNameIn)
    {
        this.hWindow = hWnd;
        this.iSequence = iSeq;
        this.sName = sNameIn;
    }

    public ProviderOptions ProviderOptions
    {
        get { return ProviderOptions.ServerSideProvider; }
    }

    public IRawElementProviderSimple HostRawElementProvider
    {
        get { return AutomationInteropProvider.HostProviderFromHandle(hWindow); }
    }

    public object GetPatternProvider(int iPatternId)
    {
        return null;
    }

    public object GetPropertyValue(int iPropertyId)
    {
        if (iPropertyId == iNameProperty)              return sName;
        if (iPropertyId == iAutomationIdProperty)      return "EdSharpNotification" + iSequence.ToString();
        if (iPropertyId == iControlTypeProperty)       return iTextControlType;
        if (iPropertyId == iFrameworkIdProperty)       return "WinForms";
        if (iPropertyId == iIsControlElementProperty)  return true;
        if (iPropertyId == iIsContentElementProperty)  return true;
        return null;
    }
}

public static class UiaNative
{
    [DllImport("UIAutomationCore.dll", PreserveSig = true)]
    public static extern int UiaRaiseNotificationEvent(
        IRawElementProviderSimple provider,
        System.Windows.Automation.AutomationNotificationKind notificationKind,
        System.Windows.Automation.AutomationNotificationProcessing notificationProcessing,
        [MarshalAs(UnmanagedType.BStr)] string displayString,
        [MarshalAs(UnmanagedType.BStr)] string activityId);
}

} // namespace Homer
