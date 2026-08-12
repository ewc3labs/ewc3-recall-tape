using ON = Microsoft.Office.Interop.OneNote;
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;

namespace RecallTape.OneNote
{
    /// <summary>
    /// v0.0.1 tape: cover the current selection, and take it off again.
    ///
    /// This is the Phase 0 exit criterion from docs/PLAN.md -- "a hotkey in real OneNote that visibly
    /// covers a real selection and uncovers it again, without eating the content."
    ///
    /// TWO ANCHOR WORLDS, TWO MECHANISMS (docs/analysis/onenote-page-xml-shapes.md):
    ///
    ///   typed text   restyle the run itself. A one:T has no coordinates of its own, so an overlay
    ///                cannot follow it -- but restyled text IS the text, so it reflows, re-wraps and
    ///                moves with the paragraph for free. Anchoring solved by not needing any.
    ///
    ///   ink, images  overlay a positioned one:Image at z = highest + 1. There is no one:Shape in the
    ///                schema, so an image carrying a solid-colour PNG is the only way to cover a
    ///                region. These are absolutely positioned, so tape does NOT yet follow the content
    ///                if the user drags it -- see RT-07.
    ///
    /// NON-DESTRUCTIVE BY CONSTRUCTION: overlays only ADD an element; text taping only WRAPS the
    /// existing CDATA in a span. Removal deletes the element we added, or unwraps the span, restoring
    /// the original byte for byte. Ink stroke data is never touched -- it is not even inline, ink
    /// elements carry only a CallbackID. Worst case if everything failed half-way is a stray black box.
    /// </summary>
    public partial class AddIn
    {
        /// <summary>Marks an overlay Image as ours. `alt` is invisible in normal use and survives round-trips.</summary>
        private const string TapeMarker = "RecallTape:tape";

        /// <summary>8x8 solid PNG, 72 bytes. A solid colour scales to any size cleanly.</summary>
        private const string TapePng =
            "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAD0lEQVR42mOQxwEYhpYEABZPF0Gm6+EsAAAAAElFTkSuQmCC";

        /// <summary>A little margin so the tape covers descenders and antialiasing, like real tape.</summary>
        private const double TapePadding = 2.0;

        /// <summary>
        /// The text-tape colour. Deliberately #1F1F1E rather than #1F1F1F or `black`: visually identical,
        /// and effectively never typed by a human, so the same value that draws the tape also identifies
        /// it on removal. One value doing both jobs beats a style plus a separate bookkeeping record
        /// that can drift out of sync with it.
        /// </summary>
        private const string TapeInk = "#1F1F1E";

        /// <summary>
        /// All THREE properties matter. A hand-made version using only background+mso-highlight looks
        /// right solely when the text is already black -- red text stays perfectly readable on the black
        /// bar. `color` is what actually hides the glyphs.
        /// </summary>
        private const string TapeSpanStyle =
            "color:" + TapeInk + ";background:" + TapeInk + ";mso-highlight:" + TapeInk;

        // ---- Tape ---------------------------------------------------------------------------

        public void TapeSelection(IRibbonControl control)
        {
            ON.Application app = null;
            try
            {
                app = new ON.Application();
                string pageId = app.Windows.CurrentWindow.CurrentPageId;

                string xml;
                app.GetPageContent(pageId, out xml, ON.PageInfo.piSelection, ON.XMLSchema.xs2013);

                var doc = new XmlDocument();
                doc.LoadXml(xml);
                var ns = doc.DocumentElement.NamespaceURI;

                int runs = TapeTextRuns(doc);
                int boxes = TapePositioned(doc, ns);

                if (runs == 0 && boxes == 0)
                {
                    Log("TapeSelection: nothing tapeable selected");
                    return;
                }

                UpdatePage(app, doc);
                Log("TapeSelection: taped " + runs + " text run(s), " + boxes + " overlay box(es)");
            }
            catch (Exception ex) { Log("TapeSelection FAILED: " + ex.Message); }
            finally { Release(app); }
        }

        /// <summary>
        /// Wrap each selected text run in the tape span.
        ///
        /// OneNote splits a one:T at the selection boundary when asked for piSelection, so the run
        /// marked selected="all" IS exactly the characters the user picked. Character-level taping with
        /// no offset arithmetic on our side -- OneNote does the splitting.
        /// </summary>
        private static int TapeTextRuns(XmlDocument doc)
        {
            int n = 0;
            foreach (XmlNode t in doc.SelectNodes("//*[local-name()='T'][@selected='all']"))
            {
                string content = t.InnerText;
                if (string.IsNullOrEmpty(content)) continue;                                     // empty run
                if (content.IndexOf(TapeInk, StringComparison.OrdinalIgnoreCase) >= 0) continue; // already taped

                SetCData(doc, t, "<span style='" + TapeSpanStyle + "'>" + content + "</span>");
                n++;
            }
            return n;
        }

        /// <summary>Overlay one box covering the selected positioned elements (ink, images).</summary>
        private static int TapePositioned(XmlDocument doc, string ns)
        {
            double x, y, w, h;
            if (!UnionOf(doc, "all", out x, out y, out w, out h)) return 0;

            int z = HighestZ(doc) + 1;

            var img = doc.CreateElement("one", "Image", ns);
            img.SetAttribute("alt", TapeMarker);

            var pos = doc.CreateElement("one", "Position", ns);
            pos.SetAttribute("x", Num(x - TapePadding));
            pos.SetAttribute("y", Num(y - TapePadding));
            pos.SetAttribute("z", z.ToString(CultureInfo.InvariantCulture));
            img.AppendChild(pos);

            var size = doc.CreateElement("one", "Size", ns);
            size.SetAttribute("width", Num(w + TapePadding * 2));
            size.SetAttribute("height", Num(h + TapePadding * 2));
            img.AppendChild(size);

            var data = doc.CreateElement("one", "Data", ns);
            data.InnerText = TapePng;
            img.AppendChild(data);

            doc.DocumentElement.AppendChild(img);
            Log(string.Format("  overlay box x={0:F1} y={1:F1} w={2:F1} h={3:F1} z={4}", x, y, w, h, z));
            return 1;
        }

        // ---- Remove tape --------------------------------------------------------------------

        public void RemoveTape(IRibbonControl control)
        {
            ON.Application app = null;
            try
            {
                app = new ON.Application();
                string pageId = app.Windows.CurrentWindow.CurrentPageId;

                string xml;
                app.GetPageContent(pageId, out xml, ON.PageInfo.piBasic, ON.XMLSchema.xs2013);

                var doc = new XmlDocument();
                doc.LoadXml(xml);
                DateTime expected = LastModified(doc);

                // 1. Unwrap taped text runs, then write the page back once.
                int runs = 0;
                foreach (XmlNode t in doc.SelectNodes("//*[local-name()='T']"))
                {
                    string content = t.InnerText;
                    if (string.IsNullOrEmpty(content)) continue;
                    if (content.IndexOf(TapeInk, StringComparison.OrdinalIgnoreCase) < 0) continue;

                    string restored = Regex.Replace(
                        content,
                        "<span[^>]*" + Regex.Escape(TapeInk) + "[^>]*>(.*?)</span>",
                        "$1",
                        RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    SetCData(doc, t, restored);
                    runs++;
                }
                if (runs > 0) UpdatePage(app, doc);

                // 2. Delete overlay images by objectID -- explicit removal of the element we own,
                //    rather than omitting it from an update payload and hoping.
                int boxes = 0;
                foreach (XmlNode n in doc.SelectNodes("//*[local-name()='Image']"))
                {
                    var alt = n.Attributes["alt"];
                    var oid = n.Attributes["objectID"];
                    if (alt == null || alt.Value != TapeMarker || oid == null) continue;
                    app.DeletePageContent(pageId, oid.Value, expected, false);
                    boxes++;
                }

                Log("RemoveTape: restored " + runs + " text run(s), removed " + boxes + " overlay box(es)");
            }
            catch (Exception ex) { Log("RemoveTape FAILED: " + ex.Message); }
            finally { Release(app); }
        }

        // ---- Helpers ------------------------------------------------------------------------

        private static bool UnionOf(XmlDocument doc, string state, out double x, out double y, out double w, out double h)
        {
            double x0 = double.MaxValue, y0 = double.MaxValue, x1 = double.MinValue, y1 = double.MinValue;
            bool any = false;

            foreach (XmlNode n in doc.SelectNodes("//*[@selected='" + state + "']"))
            {
                // Never include our own tape in the box, or taping twice grows it each time.
                var alt = n.Attributes == null ? null : n.Attributes["alt"];
                if (alt != null && alt.Value == TapeMarker) continue;

                double bx, by, bw, bh;
                if (!TryGetBox(n, out bx, out by, out bw, out bh)) continue;

                if (bx < x0) x0 = bx;
                if (by < y0) y0 = by;
                if (bx + bw > x1) x1 = bx + bw;
                if (by + bh > y1) y1 = by + bh;
                any = true;
            }

            x = x0; y = y0; w = x1 - x0; h = y1 - y0;
            return any;
        }

        private static bool TryGetBox(XmlNode n, out double x, out double y, out double w, out double h)
        {
            x = y = w = h = 0;
            var pos = n.SelectSingleNode("*[local-name()='Position']");
            var size = n.SelectSingleNode("*[local-name()='Size']");
            if (pos == null || size == null) return false;
            return Dbl(pos, "x", out x) && Dbl(pos, "y", out y)
                && Dbl(size, "width", out w) && Dbl(size, "height", out h);
        }

        private static bool Dbl(XmlNode n, string attr, out double v)
        {
            v = 0;
            var a = n.Attributes[attr];
            return a != null && double.TryParse(a.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out v);
        }

        private static int HighestZ(XmlDocument doc)
        {
            int max = 0;
            foreach (XmlNode p in doc.SelectNodes("//*[local-name()='Position']"))
            {
                var a = p.Attributes["z"];
                int z;
                if (a != null && int.TryParse(a.Value, out z) && z > max) max = z;
            }
            return max;
        }

        private static void SetCData(XmlDocument doc, XmlNode node, string content)
        {
            while (node.FirstChild != null) node.RemoveChild(node.FirstChild);
            node.AppendChild(doc.CreateCDataSection(content));
        }

        /// <summary>
        /// Write the page back with OneNote's optimistic-concurrency check ENABLED.
        ///
        /// OneMore passes DateTime.MinValue + force:true, which disables the check and always wins.
        /// Reasonable for "the user pressed a button just now"; wrong for us. RecallTape writes on a
        /// study-loop cadence against pages syncing from a Surface and a phone, and "never damage
        /// someone's notes" is non-negotiable #1 in PLAN.md. Last-write-wins is exactly the bug shape
        /// that violates it, so we pass the real expected timestamp and let OneNote refuse a stale write.
        /// </summary>
        private static void UpdatePage(ON.Application app, XmlDocument doc)
        {
            app.UpdatePageContent(doc.OuterXml, LastModified(doc), ON.XMLSchema.xs2013, false);
        }

        private static DateTime LastModified(XmlDocument doc)
        {
            var a = doc.DocumentElement.Attributes["lastModifiedTime"];
            DateTime dt;
            if (a != null && DateTime.TryParse(a.Value, CultureInfo.InvariantCulture,
                    DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out dt))
            {
                return dt;
            }
            return DateTime.MinValue;
        }

        private static string Num(double v)
        {
            return v.ToString("F2", CultureInfo.InvariantCulture);
        }

        private static void Release(ON.Application app)
        {
            if (app != null && Marshal.IsComObject(app)) Marshal.ReleaseComObject(app);
        }
    }
}
