using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace RecallTape.OneNote
{
    /// <summary>
    /// RecallTape spike add-in.
    ///
    /// PURPOSE (v0.0.1): answer the three Phase 0 questions in docs/PLAN.md with real data --
    /// what does OneNote page XML look like, is the selection identifiable in it, and what does
    /// selected INK look like. It dumps; it does not yet write anything back to a page.
    ///
    /// WHY an add-in and not a script: on Office 16.0.20228 OneNote no longer serves external
    /// automation clients. CoCreateInstance("OneNote.Application") returns an object whose every
    /// method fails E_FAIL, and OneNote never registers in the Running Object Table. The add-in
    /// path is intact because OneNote calls IN to us at OnConnection holding a live Application
    /// pointer -- we never activate anything. See docs/analysis/ for the full evidence.
    /// </summary>
    // NOTE: no [ClassInterface] attribute, deliberately -- the default (AutoDispatch) is what the
    // known-good reference implementation uses. Reflecting over River.OneMoreAddIn.AddIn shows
    // exactly three attributes: ComVisible, Guid, ProgId. Specifying ClassInterfaceType.None here
    // made OneNote demote LoadBehavior 3 -> 2 without ever calling OnConnection.
    [ComVisible(true)]
    [Guid("AA568A3C-2A53-479B-B188-2367D2E27CE4")]
    [ProgId("RecallTape.AddIn")]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        // The Application pointer OneNote hands us. Late-bound on purpose: the typed PIA would
        // mean redistributing interop assemblies, and this spike needs four methods.
        private dynamic onenote;

        private static readonly string LogDir =
            Path.Combine(Path.GetTempPath(), "RecallTape");

        /// <summary>
        /// Logs purely so that "OneNote never instantiated us" and "OneNote instantiated us but
        /// never called OnConnection" are distinguishable. Without this the two failure modes look
        /// identical from outside: LoadBehavior 3 -> 2 and an empty log directory.
        /// </summary>
        public AddIn()
        {
            Log("ctor: instantiated, pid=" + System.Diagnostics.Process.GetCurrentProcess().Id
                + " host=" + System.Diagnostics.Process.GetCurrentProcess().ProcessName);
        }

        // ---- IDTExtensibility2 -------------------------------------------------------------

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                Directory.CreateDirectory(LogDir);
                onenote = Application;
                Log("OnConnection mode=" + ConnectMode + " application=" + (Application == null ? "NULL" : "ok"));
            }
            catch (Exception ex) { Log("OnConnection FAILED: " + ex); }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("OnDisconnection " + RemoveMode);
            onenote = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { Log("OnStartupComplete"); }
        public void OnBeginShutdown(ref Array custom) { Log("OnBeginShutdown"); }

        // ---- IRibbonExtensibility ----------------------------------------------------------

        public string GetCustomUI(string RibbonID)
        {
            Log("GetCustomUI ribbonID=" + RibbonID);
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='rtTab' label='RecallTape'>
        <group id='rtSpike' label='Spike'>
          <button id='rtDump' label='Dump Page XML' size='large'
                  imageMso='FileSaveAsCurrentFileFormat' onAction='DumpPage'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        // ---- The one command --------------------------------------------------------------

        /// <summary>
        /// Dump the current page, with the selection marked, to %TEMP%\RecallTape\.
        /// Deliberately read-only: this build cannot modify a page, which makes it safe to point
        /// at any notebook including somebody's real notes.
        /// </summary>
        public void DumpPage(IRibbonControl control)
        {
            try
            {
                string pageId = onenote.Windows.CurrentWindow.CurrentPageId;
                Log("DumpPage pageId=" + pageId);

                string xml;
                onenote.GetPageContent(pageId, out xml, PageInfo.Selection);

                string stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                string path = Path.Combine(LogDir, "page-" + stamp + ".xml");
                File.WriteAllText(path, xml, Encoding.UTF8);
                Log("wrote " + path + " (" + xml.Length + " chars)");

                string hier;
                onenote.GetHierarchy("", HierarchyScope.Pages, out hier);
                File.WriteAllText(Path.Combine(LogDir, "hierarchy-" + stamp + ".xml"), hier, Encoding.UTF8);
                Log("wrote hierarchy (" + hier.Length + " chars)");
            }
            catch (Exception ex) { Log("DumpPage FAILED: " + ex); }
        }

        // ---- Logging ------------------------------------------------------------------------
        //
        // A file log is not optional here: an add-in runs inside a host we do not control, and a
        // silent failure to load is indistinguishable from a failure to work. Unknown state must
        // be visibly unknown.
        private static void Log(string message)
        {
            try
            {
                Directory.CreateDirectory(LogDir);
                File.AppendAllText(Path.Combine(LogDir, "recalltape.log"),
                    DateTime.Now.ToString("HH:mm:ss.fff") + "| " + message + Environment.NewLine);
            }
            catch { /* logging must never take the host down */ }
        }
    }
}
