using Extensibility;
using ON = Microsoft.Office.Interop.OneNote;
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
    /// pointer. See docs/analysis/ and docs/RAG_Sessions/ for the evidence.
    /// </summary>
    [ComVisible(true)]
    [Guid("AA568A3C-2A53-479B-B188-2367D2E27CE4")]
    [ProgId("RecallTape.AddIn")]
    public partial class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        // No cached Application, deliberately. Three approaches were tried and measured:
        //
        //   dynamic                      E_FAIL in GetITypeInfoFromIDispatch -- the C# dynamic COM
        //                                binder needs IDispatch::GetTypeInfo, which OneNote refuses.
        //   cast OnConnection's object   E_NOINTERFACE on IApplication {A47D3223-...}. The cast
        //                                appears to succeed because QI is lazy, then throws on first
        //                                member access. The object registers as "Application2 Class".
        //   new ON.Application()         WORKS. Constructs through the PIA coclass, which wires the
        //                                interfaces correctly.
        //
        // This is also what OneMore does exclusively, and for a second reason worth honouring:
        // holding the OnConnection reference across the process boundary stops OneNote shutting
        // down. Acquire late, release early, never cache.

        private static readonly string LogDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EWC3 Labs", "RecallTape");

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
                // Deliberately NOT keeping this reference -- see the note above.
                Log("OnConnection mode=" + ConnectMode
                    + " application=" + (Application == null ? "NULL" : "provided (not retained)"));
            }
            catch (Exception ex) { Log("OnConnection FAILED: " + ex); }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("OnDisconnection " + RemoveMode);
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete");
            // Start listening for clicks on tape strips, couriered in by the protocol handler.
            // Background thread, so it can never keep OneNote alive.
            try { StartCommandService(); }
            catch (Exception ex) { Log("command service failed to start: " + ex.Message); }
        }
        public void OnBeginShutdown(ref Array custom) { Log("OnBeginShutdown"); }

        // ---- IRibbonExtensibility ----------------------------------------------------------

        public string GetCustomUI(string RibbonID)
        {
            Log("GetCustomUI ribbonID=" + RibbonID);
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='rtTab' label='RecallTape'>
        <group id='rtTape' label='Tape'>
          <button id='rtTapeSel' label='Tape' size='large'
                  imageMso='HighlightColorPicker' onAction='TapeSelection'/>
          <button id='rtUntape' label='Remove Tape' size='large'
                  imageMso='ClearFormatting' onAction='RemoveTape'
                  screentip='Remove the selected tape'
                  supertip='Click a tape strip, or click inside taped text, then press this.'/>
          <button id='rtUntapeAll' label='Remove All' size='large'
                  imageMso='DeleteAllComments' onAction='RemoveAllTape'
                  screentip='Remove every tape on this page'
                  supertip='Asks first. Clears the whole page, which can undo a lot of work.'/>
        </group>
        <group id='rtSpike' label='Spike'>
          <button id='rtDump' label='Dump Page XML' size='large'
                  imageMso='FileSaveAsCurrentFileFormat' onAction='DumpPage'/>
          <button id='rtSurvey' label='Survey Notebooks' size='large'
                  imageMso='FindDialog' onAction='SurveyNotebooks'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        // ---- The one command --------------------------------------------------------------

        /// <summary>
        /// Dump the current page, with the selection marked, to %LOCALAPPDATA%\EWC3 Labs\RecallTape\.
        /// Deliberately read-only: this build has no code path that can modify a page, which makes
        /// it safe to point at any notebook including somebody's real notes.
        /// </summary>
        public void DumpPage(IRibbonControl control)
        {
            ON.Application app = null;
            try
            {
                app = new ON.Application();

                string pageId = app.Windows.CurrentWindow.CurrentPageId;
                string stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                Log("DumpPage pageId=" + pageId);

                // piSelection marks the current selection inside the returned XML -- the foothold
                // the whole product stands on. Schema pinned explicitly; never inherit a default
                // that Microsoft can move.
                string xml;
                app.GetPageContent(pageId, out xml, ON.PageInfo.piSelection, ON.XMLSchema.xs2013);
                Write("page-" + stamp + ".xml", xml);
                Log("wrote page-" + stamp + ".xml (" + xml.Length + " chars)");

                // Same page WITH binary data, so we can see how images/ink carry their payload.
                string full;
                app.GetPageContent(pageId, out full, ON.PageInfo.piBinaryDataSelection, ON.XMLSchema.xs2013);
                Write("page-" + stamp + "-binary.xml", full);
                Log("wrote page-" + stamp + "-binary.xml (" + full.Length + " chars)");

                string hier;
                app.GetHierarchy("", ON.HierarchyScope.hsPages, out hier, ON.XMLSchema.xs2013);
                Write("hierarchy-" + stamp + ".xml", hier);
                Log("wrote hierarchy-" + stamp + ".xml (" + hier.Length + " chars)");
            }
            catch (Exception ex) { Log("DumpPage FAILED: " + ex); }
            finally
            {
                // Release promptly. A lingering proxy across the surrogate boundary is what keeps
                // OneNote from shutting down cleanly.
                if (app != null && Marshal.IsComObject(app)) Marshal.ReleaseComObject(app);
            }
        }

        /// <summary>
        /// Walk every page in every open notebook and report what content types actually occur.
        ///
        /// Reports counts, structural skeletons, and recognizedText samples, plus a full dump of the
        /// first page containing InkWord so the real structure can be read end to end. Summary over
        /// bulk: 300 pages of raw XML is noise, not evidence. Everything stays under %LOCALAPPDATA%.
        /// </summary>
        private const int MaxPages = 300;
        private const int RecognizedTextSamples = 40;

        public void SurveyNotebooks(IRibbonControl control)
        {
            ON.Application app = null;
            var sb = new StringBuilder();
            try
            {
                app = new ON.Application();
                string hier;
                app.GetHierarchy("", ON.HierarchyScope.hsPages, out hier, ON.XMLSchema.xs2013);

                var doc = new System.Xml.XmlDocument();
                doc.LoadXml(hier);
                var ns = new System.Xml.XmlNamespaceManager(doc.NameTable);
                ns.AddNamespace("one", doc.DocumentElement.NamespaceURI);

                var pages = doc.SelectNodes("//one:Page", ns);
                sb.AppendLine("RecallTape notebook survey  " + DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
                sb.AppendLine("pages found: " + pages.Count
                    + (pages.Count > MaxPages ? "  (SURVEYING FIRST " + MaxPages + " ONLY)" : ""));
                sb.AppendLine();

                var totals = new System.Collections.Generic.SortedDictionary<string, int>();
                var samples = new System.Collections.Generic.List<string>();
                int scanned = 0, failed = 0, inkWordWithText = 0, imagesWithOcr = 0;
                string inkWordSkeleton = null, imageSkeleton = null, inkPageDump = null;

                foreach (System.Xml.XmlNode page in pages)
                {
                    if (scanned >= MaxPages) break;
                    var idAttr = page.Attributes["ID"];
                    if (idAttr == null) continue;
                    try
                    {
                        string px;
                        app.GetPageContent(idAttr.Value, out px, ON.PageInfo.piBasic, ON.XMLSchema.xs2013);
                        var pd = new System.Xml.XmlDocument();
                        pd.LoadXml(px);
                        foreach (System.Xml.XmlNode n in pd.SelectNodes("//*"))
                        {
                            string name = n.LocalName;
                            totals[name] = (totals.ContainsKey(name) ? totals[name] : 0) + 1;
                            if (name == "InkWord")
                            {
                                if (inkPageDump == null)
                                {
                                    inkPageDump = idAttr.Value;
                                    Write("ink-page-sample.xml", px);
                                }
                                var rt = n.Attributes["recognizedText"];
                                if (rt != null && rt.Value.Length > 0)
                                {
                                    inkWordWithText++;
                                    if (samples.Count < RecognizedTextSamples) samples.Add(rt.Value);
                                }
                                if (inkWordSkeleton == null) inkWordSkeleton = Skeleton(n);
                            }
                            else if (name == "Image")
                            {
                                if (n.SelectSingleNode("*[local-name()='OCRData']") != null) imagesWithOcr++;
                                if (imageSkeleton == null) imageSkeleton = Skeleton(n);
                            }
                        }
                        scanned++;
                    }
                    catch { failed++; }
                }

                sb.AppendLine("pages scanned: " + scanned + "   failed: " + failed);
                sb.AppendLine();
                sb.AppendLine("--- element totals across all scanned pages ---");
                foreach (var kv in totals) sb.AppendLine(string.Format("  {0,-22} {1}", kv.Key, kv.Value));
                sb.AppendLine();
                sb.AppendLine("--- the questions this was run to answer ---");
                sb.AppendLine("  InkWord elements ........... " + Count(totals, "InkWord"));
                sb.AppendLine("  ...with recognizedText ..... " + inkWordWithText);
                sb.AppendLine("  InkDrawing elements ........ " + Count(totals, "InkDrawing"));
                sb.AppendLine("  InkParagraph elements ...... " + Count(totals, "InkParagraph"));
                sb.AppendLine("  Image elements ............. " + Count(totals, "Image"));
                sb.AppendLine("  ...with OCRData ............ " + imagesWithOcr);
                sb.AppendLine("  Table elements ............. " + Count(totals, "Table"));
                sb.AppendLine();
                sb.AppendLine("full XML of first InkWord page: " + (inkPageDump != null ? "ink-page-sample.xml" : "(no InkWord found)"));
                sb.AppendLine();
                sb.AppendLine("--- structural skeletons (attribute names only, content elided) ---");
                sb.AppendLine("  InkWord: " + (inkWordSkeleton ?? "(none found)"));
                sb.AppendLine("  Image:   " + (imageSkeleton ?? "(none found)"));
                sb.AppendLine();
                sb.AppendLine("--- recognizedText samples (capped at " + RecognizedTextSamples + ", deliberately) ---");
                if (samples.Count == 0) sb.AppendLine("  (none)");
                foreach (var t in samples) sb.AppendLine("  [" + t + "]");

                string stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
                Write("survey-" + stamp + ".txt", sb.ToString());
                Log("survey complete: " + scanned + " pages -> survey-" + stamp + ".txt");
            }
            catch (Exception ex) { Log("SurveyNotebooks FAILED: " + ex); }
            finally
            {
                if (app != null && Marshal.IsComObject(app)) Marshal.ReleaseComObject(app);
            }
        }

        private static int Count(System.Collections.Generic.SortedDictionary<string, int> d, string k)
        {
            return d.ContainsKey(k) ? d[k] : 0;
        }

        /// <summary>Element name plus attribute names; content-bearing values are elided.</summary>
        private static string Skeleton(System.Xml.XmlNode n)
        {
            var sb = new StringBuilder("<" + n.LocalName);
            foreach (System.Xml.XmlAttribute a in n.Attributes)
            {
                bool contentish = a.Name == "recognizedText" || a.Name == "alt";
                sb.Append(" ").Append(a.Name).Append("=")
                  .Append(contentish ? "{elided}" : (a.Value.Length > 28 ? "{...}" : a.Value));
            }
            sb.Append(">");
            foreach (System.Xml.XmlNode c in n.ChildNodes) sb.Append(" <").Append(c.LocalName).Append(">");
            return sb.ToString();
        }

        private static void Write(string name, string content)
        {
            Directory.CreateDirectory(LogDir);
            File.WriteAllText(Path.Combine(LogDir, name), content, Encoding.UTF8);
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
