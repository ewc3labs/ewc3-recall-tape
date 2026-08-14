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
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'
           onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='rtTab' label='RecallTape'>

        <!-- Layout is deliberate. The RecallTape tab sits at the FAR RIGHT of the tab strip, so
             the pointer arrives at the right-hand end of the ribbon. Putting Tape - the button
             pressed constantly - leftmost meant crossing the whole ribbon every single time.
             So the order runs quietest-to-loudest, left to right, and Tape is both rightmost and
             the only large button. Remove All is small and sits furthest from Tape, because it is
             the one that can undo an evening. -->

        <group id='rtSetup' label='Setup'>
          <menu id='rtSettings' label='Settings' size='large' imageMso='AddInsMenu'
                screentip='RecallTape settings'>
            <checkBox id='rtDev' label='Show developer tools'
                      getPressed='GetDevPressed' onAction='ToggleDev'
                      screentip='Dump Page XML and Survey Notebooks'
                      supertip='Diagnostics for reporting bugs. Off by default - they are no use while studying.'/>
            <menuSeparator id='rtSepLog' title='Logging'/>
            <checkBox id='rtLogOn' label='Write a log file'
                      getPressed='GetLogPressed' onAction='ToggleLog'
                      supertip='Records what each button did. No page content is ever written to it.'/>
            <menu id='rtLogSize' label='Log size limit' imageMso='FileSave'>
              <toggleButton id='rtLog1'  label='1 MB'   getPressed='GetLog1'  onAction='SetLog1'/>
              <toggleButton id='rtLog5'  label='5 MB'   getPressed='GetLog5'  onAction='SetLog5'/>
              <toggleButton id='rtLog20' label='20 MB'  getPressed='GetLog20' onAction='SetLog20'/>
            </menu>
            <button id='rtOpenLog' label='Open log folder' imageMso='Folder' onAction='OpenLogFolder'/>
            <menuSeparator id='rtSepAbout'/>
            <button id='rtAbout' label='About RecallTape' imageMso='Info' onAction='About'/>
          </menu>
        </group>

        <group id='rtRemoveGrp' label='Remove'>
          <button id='rtUntapeAll' label='Remove All'
                  imageMso='InkDeleteAllInk' onAction='RemoveAllTape'
                  screentip='Remove every tape on this page'
                  supertip='Asks first. Clears the whole page, which can undo a lot of work.'/>
          <button id='rtUntape' label='Remove Tape'
                  imageMso='ClearFormatting' onAction='RemoveTape'
                  screentip='Remove the tape you are pointing at'
                  supertip='Put the cursor in taped text, or click a tape strip, then press this.'/>
        </group>

        <group id='rtTapeGrp' label='Tape'>
          <button id='rtTapeSel' label='Tape' size='large'
                  imageMso='ShapeFillColorPicker' onAction='TapeSelection'
                  screentip='Cover the selection'
                  supertip='Select text, ink or an image. Select nothing and you get a box you can drag anywhere.'/>
        </group>

        <group id='rtSpike' label='Developer' getVisible='GetDevVisible'>
          <button id='rtDump' label='Dump Page XML' size='large'
                  imageMso='FileSaveAsCurrentFileFormat' onAction='DumpPage'/>
          <button id='rtSurvey' label='Survey Notebooks' size='large'
                  imageMso='FindDialog' onAction='SurveyNotebooks'/>
          <gallery id='rtIcons' label='Icon Browser' size='large' imageMso='PictureStylesGallery'
                   columns='8' itemWidth='32' itemHeight='32'
                   getItemCount='GetIconCount' getItemImageMso='GetIconImage'
                   getItemLabel='GetIconLabel' getItemScreentip='GetIconLabel'
                   onAction='PickIcon'
                   screentip='Browse Office icons, rendered by OneNote itself'
                   supertip='An icon that shows up blank does not exist in OneNote. Click one to see its id.'/>
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

        // ---- Ribbon state -------------------------------------------------------------------
        //
        // The ribbon is built once, at load. Anything that changes afterwards needs the IRibbonUI
        // handed to us here, so we can ask Office to re-query a control.

        // Typed as OBJECT, and called by name. This is not laziness.
        //
        // Declared as a hand-rolled IRibbonUI, onLoad NEVER FIRED - no error, no log line, nothing.
        // Office binds ribbon callbacks through IDispatch and silently drops any whose signature it
        // cannot marshal, so a callback that is merely unbindable is indistinguishable from one that
        // was never wired up. That is the same failure shape that cost an hour on IDTExtensibility2,
        // and the same lesson: our declaration of an Office interface is a guess, and Office does not
        // tell you when the guess is wrong.
        //
        // object + InvokeMember goes through IDispatch by NAME, which is how the ribbon talks anyway.
        // Nothing to get wrong, and it works with whatever Office actually hands us.
        private object ribbon;

        public void OnRibbonLoad(object ribbonUI)
        {
            ribbon = ribbonUI;
            Log("ribbon loaded (" + (ribbonUI == null ? "null" : ribbonUI.GetType().Name)
                + "); developer tools " + (Settings.DeveloperTools ? "ON" : "off"));
        }

        private void Refresh(string controlId)
        {
            if (ribbon == null) { Log("ribbon refresh skipped: onLoad never gave us a ribbon"); return; }
            try
            {
                ribbon.GetType().InvokeMember("InvalidateControl",
                    System.Reflection.BindingFlags.InvokeMethod, null, ribbon,
                    new object[] { controlId });
            }
            catch (Exception ex) { Log("ribbon refresh failed: " + ex.Message); }
        }

        public bool GetDevVisible(IRibbonControl control) { return Settings.DeveloperTools; }
        public bool GetDevPressed(IRibbonControl control) { return Settings.DeveloperTools; }

        public void ToggleDev(IRibbonControl control, bool pressed)
        {
            Settings.DeveloperTools = pressed;
            Log("developer tools " + (pressed ? "ON" : "off"));
            Refresh("rtSpike");
        }

        public bool GetLogPressed(IRibbonControl control) { return Settings.Logging; }

        public void ToggleLog(IRibbonControl control, bool pressed)
        {
            // Log the change BEFORE honouring it when switching off, so the file says why it stops.
            if (!pressed) Log("logging turned OFF by user");
            Settings.Logging = pressed;
            if (pressed) Log("logging turned on");
        }

        public bool GetLog1(IRibbonControl control) { return Settings.LogMaxKB == 1024; }
        public bool GetLog5(IRibbonControl control) { return Settings.LogMaxKB == 5120; }
        public bool GetLog20(IRibbonControl control) { return Settings.LogMaxKB == 20480; }

        public void SetLog1(IRibbonControl control, bool pressed) { SetLogSize(1024); }
        public void SetLog5(IRibbonControl control, bool pressed) { SetLogSize(5120); }
        public void SetLog20(IRibbonControl control, bool pressed) { SetLogSize(20480); }

        private void SetLogSize(int kb)
        {
            Settings.LogMaxKB = kb;
            Log("log size limit set to " + kb + " KB");
            Refresh("rtLog1"); Refresh("rtLog5"); Refresh("rtLog20");
        }

        public void OpenLogFolder(IRibbonControl control)
        {
            try
            {
                Directory.CreateDirectory(LogDir);
                System.Diagnostics.Process.Start("explorer.exe", "\"" + LogDir + "\"");
            }
            catch (Exception ex) { Log("OpenLogFolder failed: " + ex.Message); }
        }

        public void About(IRibbonControl control)
        {
            var v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            MessageBoxW(IntPtr.Zero, "RecallTape " + v + "\n\n"
                + "Tape over your own notes and quiz yourself.\n\n"
                + "github.com/ewc3labs/ewc3-recall-tape", "About RecallTape", MB_ICONINFORMATION);
        }

        // ---- Icon browser -------------------------------------------------------------------
        //
        // Which imageMso ids actually exist is per-application, not per-Office: HighlightColorPicker
        // and DeleteAllComments are real ids that render nothing in OneNote, which is why two of our
        // buttons shipped blank. No published list says which survive here.
        //
        // So let the host answer. OneNote draws these itself, and anything it cannot draw comes back
        // blank - which is the answer. Developer tools only.
        private static readonly string[] IconIds = new string[]
        {
            // covering / filling - candidates for Tape
            "ShapeFillColorPicker", "ObjectPictureFill", "FillDown", "FillRight", "RecolorColorPicker",
            "BlackAndWhiteInverseGrayscale", "WatermarkGallery", "PictureCrop", "PictureReset",
            "GroupBorder", "ShapeRightArrow", "GroupShapes", "SnapToGrid", "PictureCompress",
            "HighlightColorPicker", "TextHighlightColorPicker", "FontColorPicker", "Strikethrough",
            // removing / clearing - candidates for Remove
            "InkDeleteAllInk", "Clear", "FormFieldClear", "PivotTableClearMenu", "HyperlinkRemove",
            "PageNumbersRemove", "RemoveDuplicates", "RemoveCitation", "SheetDelete",
            "RecordsDeleteRecord", "DataFormDeleteRecord", "AdpOutputOperationsTableRemove",
            "HeaderFooterRemoveHeaderWord", "DeleteAllComments", "CancelTaskAssignment",
            // settings / tools
            "AddInsMenu", "ReadingViewToolsMenu", "AutoCorrect", "Organizer", "ControlTitle",
            "Info", "Help", "SearchUI", "FileFind", "Folder", "FileSave", "FileOpen",
            "StylesDialogClassic", "StylesModifyStyle", "StylesStyleInspector", "ApplyStylesPane",
            "SourceControlRefreshStatus", "FileCompactAndRepairDatabase", "SizeToGridOutlook",
            // general - things that might read as tape, study or notes
            "A", "Bullets", "Numbering", "ChangeCase", "Copy", "PasteAlternative", "TextBoxInsert",
            "NewNote", "NewTask", "AcceptTask", "StarUnratedFull", "ReminderGallery", "StartTimer",
            "CollapseAll", "OutlineExpandAll", "OutlineShowDetail", "OutlineDemote", "IndentDecrease",
            "ReviewShowOrHideMarkup", "ReviewShowMarkupMenu", "CompareAndCombine", "WordCount",
            "Spelling", "XmlSource", "CodeEdit", "ImagerScan", "BarcodeInsert", "CreateDiagram",
            "PanAndZoomWindow", "PageNext", "PagePrevious", "UpArrow2", "DownArrow2", "SortUp",
            "TableInsert", "ConvertTextToTable", "HorizontalLineInsert", "BookmarkInsert",
            "PictureInsertFromFile", "PictureStylesGallery", "GroupPictureStyles", "ShapeSmileyFace"
        };

        public int GetIconCount(IRibbonControl control) { return IconIds.Length; }
        public string GetIconImage(IRibbonControl control, int index) { return IconIds[index]; }
        public string GetIconLabel(IRibbonControl control, int index) { return IconIds[index]; }

        public void PickIcon(IRibbonControl control, string selectedId, int selectedIndex)
        {
            string name = IconIds[selectedIndex];
            Log("icon picked: " + name);
            // No clipboard: that would pull System.Windows.Forms into the surrogate for one line,
            // and we kept it out on purpose. The id is short enough to read and type.
            MessageBoxW(IntPtr.Zero, "imageMso id:\n\n    " + name, "Icon Browser", MB_ICONINFORMATION);
        }

        // ---- Logging ------------------------------------------------------------------------
        //
        // A file log is not optional here: an add-in runs inside a host we do not control, and a
        // silent failure to load is indistinguishable from a failure to work. Unknown state must
        // be visibly unknown.
        private static void Log(string message)
        {
            if (!Settings.Logging) return;
            try
            {
                Directory.CreateDirectory(LogDir);
                string path = Path.Combine(LogDir, "recalltape.log");
                Roll(path);
                File.AppendAllText(path,
                    DateTime.Now.ToString("HH:mm:ss.fff") + "| " + message + Environment.NewLine);
            }
            catch { /* logging must never take the host down */ }
        }

        /// <summary>
        /// Keep the log bounded.
        ///
        /// We log every operation, and this add-in is meant to be used daily for a whole semester on
        /// a machine whose owner never asked for our diagnostics. An unbounded log in someone's
        /// profile is rude - more so now that it lives in LocalAppData rather than a temp folder
        /// Windows would eventually empty for us.
        ///
        /// One previous file is kept, so the worst case is twice the limit and a user can still see
        /// what happened before the roll. Deleting the older one silently is the point: nothing here
        /// is worth keeping forever.
        /// </summary>
        private static void Roll(string path)
        {
            try
            {
                var fi = new FileInfo(path);
                if (!fi.Exists || fi.Length < (long)Settings.LogMaxKB * 1024) return;

                string previous = path + ".1";
                if (File.Exists(previous)) File.Delete(previous);
                File.Move(path, previous);
                File.AppendAllText(path, DateTime.Now.ToString("HH:mm:ss.fff")
                    + "| log rolled at " + fi.Length + " bytes; previous kept as recalltape.log.1"
                    + Environment.NewLine);
            }
            catch { }
        }
    }
}
