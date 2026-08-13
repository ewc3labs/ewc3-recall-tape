using ON = Microsoft.Office.Interop.OneNote;
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;

namespace RecallTape.OneNote
{
    /// <summary>
    /// Tape: cover the current selection, peek under it, and take it off again.
    ///
    /// TWO ANCHOR WORLDS, TWO MECHANISMS (docs/analysis/onenote-page-xml-shapes.md):
    ///
    ///   typed text   restyle the run itself. A one:T has no coordinates of its own, so an overlay
    ///                cannot follow it -- but restyled text IS the text, so it reflows, re-wraps and
    ///                moves with the paragraph for free. Anchoring solved by not needing any.
    ///
    ///   ink, images  overlay a positioned one:Image at z = highest + 1. There is no one:Shape in the
    ///                schema, so an image carrying a solid-colour PNG is the only way to cover a region.
    ///
    /// NON-DESTRUCTIVE BY CONSTRUCTION: overlays only ADD an element; text taping only WRAPS the
    /// existing CDATA in a span. Removal deletes what we added or unwraps the span, restoring the
    /// original byte for byte. Ink stroke data is never touched -- it is not even inline.
    /// </summary>
    public partial class AddIn
    {
        /// <summary>`alt` marker: "RecallTape:{state}:{guid}". State is what it is; guid is which one.</summary>
        private const string TapePrefix = "RecallTape:";
        private const string StateTaped = "tape";
        private const string StatePeeked = "peek";

        /// <summary>8x8 solid PNG, 72 bytes. A solid colour scales to any size cleanly.</summary>
        private const string OpaquePng =
            "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAD0lEQVR42mOQxwEYhpYEABZPF0Gm6+EsAAAAAElFTkSuQmCC";

        /// <summary>
        /// 8x8 fully transparent RGBA PNG, 70 bytes. Peeking SWAPS the payload rather than deleting the
        /// element: the Image keeps its position, size, id and hyperlink, so it stays clickable and the
        /// tape can be put back. The state lives in the element itself, which means there is no separate
        /// bookkeeping record to drift out of sync with the page.
        /// </summary>
        private const string ClearPng =
            "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAYAAADED76LAAAADUlEQVR42mNgGAUgAAABCAABXbcZDQAAAABJRU5ErkJggg==";

        /// <summary>A little margin so the tape covers descenders and antialiasing, like real tape.</summary>
        private const double TapePadding = 2.0;

        /// <summary>
        /// The text-tape colour. Deliberately #1F1F1E rather than #1F1F1F or `black`: visually identical,
        /// and effectively never typed by a human, so the same value that draws the tape also identifies
        /// it on removal.
        /// </summary>
        private const string TapeInk = "#1F1F1E";

        /// <summary>
        /// All THREE properties matter. A version using only background+mso-highlight looks right solely
        /// when the text is already black -- coloured text stays perfectly readable on the bar.
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

                // Validate what we changed rather than assuming it stuck. OneNote is free to drop
                // attributes it does not like on write, and a silently-dropped hyperlink is
                // indistinguishable from "clicking does not work" when you are staring at the page.
                if (boxes > 0) VerifyTape(app, pageId);
            }
            catch (Exception ex) { Log("TapeSelection FAILED: " + ex.Message); }
            finally { Release(app); }
        }

        /// <summary>
        /// Wrap each selected text run in the tape span.
        ///
        /// OneNote splits a one:T at the selection boundary when asked for piSelection, so the run
        /// marked selected="all" IS exactly the characters the user picked -- character-level taping
        /// with no offset arithmetic on our side.
        /// </summary>
        private static int TapeTextRuns(XmlDocument doc)
        {
            int n = 0;
            foreach (XmlNode t in doc.SelectNodes("//*[local-name()='T'][@selected='all']"))
            {
                string content = t.InnerText;
                if (string.IsNullOrEmpty(content)) continue;
                if (content.IndexOf(TapeInk, StringComparison.OrdinalIgnoreCase) >= 0) continue;

                SetCData(doc, t, "<span style='" + TapeSpanStyle + "'>" + content + "</span>");
                n++;
            }
            return n;
        }

        /// <summary>Overlay one clickable box covering the selected positioned elements (ink, images).</summary>
        private static int TapePositioned(XmlDocument doc, string ns)
        {
            double x, y, w, h;
            if (!UnionOf(doc, "all", out x, out y, out w, out h)) return 0;

            // Our own id, not OneNote's objectID. An objectID does not exist until after the page is
            // written, which would force a second write just to attach the hyperlink. A GUID we mint
            // lets tape and link be created in one pass, and it survives round-trips inside `alt`.
            string id = Guid.NewGuid().ToString();
            int z = HighestZ(doc) + 1;

            var img = doc.CreateElement("one", "Image", ns);
            img.SetAttribute("alt", TapePrefix + StateTaped + ":" + id);
            img.SetAttribute("hyperlink", "recalltape://toggle/" + id);

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
            data.InnerText = OpaquePng;
            img.AppendChild(data);

            doc.DocumentElement.AppendChild(img);
            Log(string.Format("  overlay box x={0:F1} y={1:F1} w={2:F1} h={3:F1} z={4} id={5}", x, y, w, h, z, id));
            return 1;
        }

        /// <summary>
        /// Re-read the page and report what actually persisted. Cheap, and it turns "the click does
        /// nothing" into either "OneNote dropped the attribute" or "OneNote kept it and still does
        /// nothing" -- two very different problems.
        /// </summary>
        private static void VerifyTape(ON.Application app, string pageId)
        {
            try
            {
                string xml;
                app.GetPageContent(pageId, out xml, ON.PageInfo.piBasic, ON.XMLSchema.xs2013);
                var doc = new XmlDocument();
                doc.LoadXml(xml);

                foreach (XmlNode n in doc.SelectNodes("//*[local-name()='Image']"))
                {
                    var alt = n.Attributes["alt"];
                    if (alt == null || !alt.Value.StartsWith(TapePrefix, StringComparison.OrdinalIgnoreCase)) continue;

                    var link = n.Attributes["hyperlink"];
                    var oid = n.Attributes["objectID"];
                    Log("  verify: alt=" + alt.Value
                        + " hyperlink=" + (link == null ? "DROPPED BY ONENOTE" : link.Value)
                        + " objectID=" + (oid == null ? "(none)" : "ok"));
                }
            }
            catch (Exception ex) { Log("  verify failed: " + ex.Message); }
        }

        // ---- Peek ---------------------------------------------------------------------------

        /// <summary>
        /// Toggle one overlay between covered and revealed, in place.
        ///
        /// Swaps the PNG payload rather than deleting and recreating the element, so position, size,
        /// hyperlink and identity all survive -- the tape can be put back, and it stays clickable while
        /// revealed. Called from the command service when the user clicks a tape strip.
        /// </summary>
        private void TogglePeek(string id)
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

                XmlNode target = null;
                foreach (XmlNode n in doc.SelectNodes("//*[local-name()='Image']"))
                {
                    var alt = n.Attributes["alt"];
                    if (alt != null && alt.Value.EndsWith(":" + id, StringComparison.OrdinalIgnoreCase))
                    {
                        target = n;
                        break;
                    }
                }

                if (target == null)
                {
                    Log("TogglePeek: no tape with id " + id + " on the current page");
                    return;
                }

                bool covered = target.Attributes["alt"].Value.IndexOf(":" + StateTaped + ":",
                                   StringComparison.OrdinalIgnoreCase) >= 0;

                target.Attributes["alt"].Value =
                    TapePrefix + (covered ? StatePeeked : StateTaped) + ":" + id;

                // Replace whatever carries the image payload.
                //
                // WHY not just edit Data: a piBasic read returns a <one:CallbackID> REFERENCE rather
                // than inline <one:Data> -- that is the 12 KB vs 50 KB difference documented in
                // docs/analysis/onenote-page-xml-shapes.md. Reading piBinaryData just to find bytes we
                // are about to throw away would pull the whole payload for every click. We only ever
                // WRITE bytes here, so drop whichever reference is present and hand OneNote fresh Data.
                var ns2 = doc.DocumentElement.NamespaceURI;
                var payload = target.SelectSingleNode("*[local-name()='Data']");
                if (payload != null) target.RemoveChild(payload);
                var callback = target.SelectSingleNode("*[local-name()='CallbackID']");
                if (callback != null) target.RemoveChild(callback);

                var fresh = doc.CreateElement("one", "Data", ns2);
                fresh.InnerText = covered ? ClearPng : OpaquePng;
                target.AppendChild(fresh);

                UpdatePage(app, doc);
                Log("TogglePeek: " + id + " -> " + (covered ? "peeked" : "covered"));
            }
            catch (Exception ex) { Log("TogglePeek FAILED: " + ex.Message); }
            finally { Release(app); }
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

                // 2. Delete overlays by objectID -- both covered and peeked ones.
                int boxes = 0;
                foreach (XmlNode n in doc.SelectNodes("//*[local-name()='Image']"))
                {
                    var alt = n.Attributes["alt"];
                    var oid = n.Attributes["objectID"];
                    if (alt == null || oid == null) continue;
                    if (!alt.Value.StartsWith(TapePrefix, StringComparison.OrdinalIgnoreCase)) continue;

                    app.DeletePageContent(pageId, oid.Value, expected, false);
                    boxes++;
                }

                // Never let "I did nothing" look like "there was nothing to do". Hand-made
                // black-on-black is somebody's own formatting and we must not strip it -- but staying
                // silent about it reads as a bug.
                int foreign = 0;
                foreach (XmlNode t in doc.SelectNodes("//*[local-name()='T']"))
                {
                    string c = t.InnerText;
                    if (string.IsNullOrEmpty(c)) continue;
                    if (c.IndexOf(TapeInk, StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    if (Regex.IsMatch(c, @"mso-highlight|background\s*:", RegexOptions.IgnoreCase)) foreign++;
                }

                Log("RemoveTape: restored " + runs + " text run(s), removed " + boxes + " overlay box(es)"
                    + (foreign > 0
                        ? "; left " + foreign + " run(s) alone - highlighted, but not RecallTape's"
                        : ""));
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
                if (alt != null && alt.Value.StartsWith(TapePrefix, StringComparison.OrdinalIgnoreCase)) continue;

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
        /// someone's notes" is non-negotiable #1 in PLAN.md.
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
