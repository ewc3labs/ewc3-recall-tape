using ON = Microsoft.Office.Interop.OneNote;
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
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

        /// <summary>Outline colour for a peeked tape. Grey reads against both light pages and dark images.</summary>
        private static readonly Color PeekBorder = Color.FromArgb(180, 128, 128, 128);
        private const int PeekBorderPx = 2;

        /// <summary>Sanity cap. A page-sized tape should not produce a multi-megapixel PNG.</summary>
        private const int MaxPeekPx = 2000;

        /// <summary>A little margin so the tape covers descenders and antialiasing, like real tape.</summary>
        private const double TapePadding = 2.0;

        /// <summary>
        /// The text-tape colour. Deliberately #1F1F1E rather than #1F1F1F or `black`: visually identical,
        /// and effectively never typed by a human, so the same value that draws the tape also identifies
        /// it on removal.
        /// </summary>
        // OneNote REWRITES inline CSS on round-trip. Measured, from a live page:
        //
        //   we write:   color:#1F1F1E;background:#1F1F1E;mso-highlight:#1F1F1E
        //   we get back: background:black;mso-highlight:black
        //
        // It drops color: outright and snaps #1F1F1E to the named colour "black". So identifying our
        // tape by that exact string worked on tape we had just applied and FAILED on tape that had
        // been saved - the user selected a black bar and was told there was no tape there.
        //
        // Worse, #1F1F1E was never distinctive: it is OneNote's own default body text colour, and it
        // appears on ordinary untaped paragraphs in the same page. It was never a marker.
        //
        // So identity now comes from two places, in order of trust:
        //   1. a one:Meta on the paragraph, which the schema guarantees and OneNote must round-trip
        //   2. the SHAPE of the style - background and mso-highlight both set to the same dark colour
        // Rule 2 is what recognises tape applied by earlier builds, and it stays.
        private const string TapeInk = "black";
        private const string TapeMetaName = "RecallTape";

        /// <summary>
        /// All THREE properties matter. A version using only background+mso-highlight looks right solely
        /// when the text is already black -- coloured text stays perfectly readable on the bar.
        /// </summary>
        private const string TapeSpanStyle =
            "color:" + TapeInk + ";background:" + TapeInk + ";mso-highlight:" + TapeInk;

        /// <summary>A span is ours if background and mso-highlight are both a dark colour.</summary>
        private static readonly Regex TapeSpanRx = new Regex(
            @"<span([^>]*)>(.*?)</span>",
            RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex DarkFillRx = new Regex(
            @"(?:background|mso-highlight)\s*:\s*(?:black|#0{3,6}|#1F1F1E|rgb\(\s*(?:[0-9]|[1-2][0-9]|3[01])\s*,)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static bool IsTapeStyle(string spanAttrs)
        {
            if (string.IsNullOrEmpty(spanAttrs)) return false;
            // Both, not either: a user highlighting text is one property. Tape is always both.
            return DarkFillRx.IsMatch(spanAttrs)
                && spanAttrs.IndexOf("mso-highlight", StringComparison.OrdinalIgnoreCase) >= 0
                && spanAttrs.IndexOf("background", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>Does this run carry our tape?</summary>
        private static bool HasTape(string cdata)
        {
            if (string.IsNullOrEmpty(cdata)) return false;
            foreach (Match m in TapeSpanRx.Matches(cdata))
                if (IsTapeStyle(m.Groups[1].Value)) return true;
            return false;
        }

        /// <summary>Unwrap our spans, leaving anyone else's alone.</summary>
        private static string StripTape(string cdata)
        {
            return TapeSpanRx.Replace(cdata, delegate(Match m)
            {
                return IsTapeStyle(m.Groups[1].Value) ? m.Groups[2].Value : m.Value;
            });
        }

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

                string freeId = null;
                if (runs == 0 && boxes == 0)
                {
                    // Nothing selectable was selected. That is the NORMAL case for a page background
                    // image, a printout slide, or anything else OneNote refuses to hand us
                    // (backgroundImage="true" is never marked selected, no matter where the user
                    // clicks). Rather than fail, drop a free-floating tape box on the page and let
                    // them drag and resize it onto whatever they wanted covered.
                    //
                    // This works because our tape IS a one:Image, so OneNote already gives it move
                    // and resize handles for free. We are not building a canvas editor; we are
                    // placing an object and getting out of the way.
                    freeId = AddFreeTape(doc, ns);
                    boxes = 1;
                }

                UpdatePage(app, doc);
                Log("TapeSelection: taped " + runs + " text run(s), " + boxes + " overlay box(es)");

                // Validate what we changed rather than assuming it stuck. OneNote is free to drop
                // attributes it does not like on write, and a silently-dropped hyperlink is
                // indistinguishable from "clicking does not work" when you are staring at the page.
                if (boxes > 0) VerifyTape(app, pageId);

                // A printout page can be 24,000 points tall and the API exposes no viewport or scroll
                // position, so a box we place is quite possibly nowhere near what the user is looking
                // at. NavigateTo is the only way to reunite them with it.
                if (freeId != null) NavigateToTape(app, pageId, freeId);
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
                if (HasTape(content)) continue;

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

        /// <summary>
        /// Place a free-floating tape box for the user to drag onto whatever they want covered.
        ///
        /// Placement is a compromise, not a guess: the OneNote API exposes no viewport or scroll
        /// position (Window gives IDs, FullPageView, Active, DockedLocation - nothing about where the
        /// user is looking), so there is no such thing as "in the middle of the screen". We cascade
        /// from the last tape on the page so repeated presses stack neatly, fall back to the top-left
        /// of the page's content, and then navigate the user to it.
        /// </summary>
        private static string AddFreeTape(XmlDocument doc, string ns)
        {
            const double W = 240, H = 80, CASCADE = 24;

            double x = double.MaxValue, y = double.MaxValue;
            double lastX = double.MinValue, lastY = double.MinValue;

            foreach (XmlNode n in doc.SelectNodes("//*[local-name()='Image' or local-name()='InkDrawing' or local-name()='Outline']"))
            {
                double bx, by, bw, bh;
                if (!TryGetBox(n, out bx, out by, out bw, out bh)) continue;

                var alt = n.Attributes == null ? null : n.Attributes["alt"];
                bool ours = alt != null && alt.Value.StartsWith(TapePrefix, StringComparison.OrdinalIgnoreCase);

                if (ours) { if (by > lastY) { lastX = bx; lastY = by; } }
                else { if (bx < x) x = bx; if (by < y) y = by; }
            }

            double px, py;
            if (lastY > double.MinValue) { px = lastX + CASCADE; py = lastY + CASCADE; }   // stack on the last one
            else if (x < double.MaxValue) { px = x + 40; py = y + 40; }                    // just inside the content
            else { px = 100; py = 100; }                                                   // empty page

            string id = Guid.NewGuid().ToString();
            int z = HighestZ(doc) + 1;

            var img = doc.CreateElement("one", "Image", ns);
            img.SetAttribute("alt", TapePrefix + StateTaped + ":" + id);
            img.SetAttribute("hyperlink", "recalltape://toggle/" + id);

            var pos = doc.CreateElement("one", "Position", ns);
            pos.SetAttribute("x", Num(px));
            pos.SetAttribute("y", Num(py));
            pos.SetAttribute("z", z.ToString(CultureInfo.InvariantCulture));
            img.AppendChild(pos);

            var size = doc.CreateElement("one", "Size", ns);
            size.SetAttribute("width", Num(W));
            size.SetAttribute("height", Num(H));
            // isSetByUser tells OneNote the size is deliberate, so it shows resize handles rather
            // than treating the dimensions as an intrinsic property of the bitmap.
            size.SetAttribute("isSetByUser", "true");
            img.AppendChild(size);

            var data = doc.CreateElement("one", "Data", ns);
            data.InnerText = OpaquePng;
            img.AppendChild(data);

            doc.DocumentElement.AppendChild(img);
            Log(string.Format("  free tape box x={0:F1} y={1:F1} {2}x{3} z={4} id={5}", px, py, W, H, z, id));
            return id;
        }

        /// <summary>Scroll the user to a tape we just created, by looking up the objectID OneNote assigned it.</summary>
        private static void NavigateToTape(ON.Application app, string pageId, string id)
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
                    var oid = n.Attributes["objectID"];
                    if (alt == null || oid == null) continue;
                    if (!alt.Value.EndsWith(":" + id, StringComparison.OrdinalIgnoreCase)) continue;

                    app.NavigateTo(pageId, oid.Value, false);
                    Log("  navigated to the new tape box");
                    return;
                }
            }
            catch (Exception ex) { Log("  navigate failed: " + ex.Message); }
        }

        /// <summary>
        /// A transparent PNG with a visible outline, generated at the box's own size.
        ///
        /// WHY GENERATED RATHER THAN A CONSTANT: the covered state is a solid colour, so an 8x8 PNG
        /// scales to any size perfectly. A BORDER does not - stretch an 8x8 bordered image across a
        /// 240x80 box and the 1px edge becomes 30px wide and 10px tall. The outline has to be drawn at
        /// the real pixel dimensions, which means regenerating it whenever the box is peeked.
        ///
        /// WHY IT MATTERS: a fully transparent peeked tape is invisible AND unfindable. With twenty
        /// tapes on a page, revealing one and scrolling away loses it - you are reduced to hovering
        /// over the page hunting for the hyperlink cursor. The outline says "tape lives here, revealed".
        /// </summary>
        private static string BorderPng(double widthPt, double heightPt)
        {
            try
            {
                int w = Math.Max(PeekBorderPx * 2 + 1, Math.Min(MaxPeekPx, (int)Math.Round(widthPt)));
                int h = Math.Max(PeekBorderPx * 2 + 1, Math.Min(MaxPeekPx, (int)Math.Round(heightPt)));

                using (var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                {
                    using (var g = Graphics.FromImage(bmp))
                    using (var pen = new Pen(PeekBorder, PeekBorderPx))
                    {
                        g.Clear(Color.Transparent);
                        // Inset by half the pen width so the stroke lands inside the bitmap rather
                        // than being clipped in half by the edges.
                        float o = PeekBorderPx / 2f;
                        g.DrawRectangle(pen, o, o, w - PeekBorderPx, h - PeekBorderPx);
                    }
                    using (var ms = new MemoryStream())
                    {
                        bmp.Save(ms, ImageFormat.Png);
                        return Convert.ToBase64String(ms.ToArray());
                    }
                }
            }
            catch (Exception ex)
            {
                // Never fail a peek over cosmetics - fall back to the fully transparent constant.
                Log("  border generation failed, using plain transparent: " + ex.Message);
                return ClearPng;
            }
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
                if (covered)
                {
                    // Read the size at toggle time, not at tape time: the user may well have dragged
                    // the box to a different shape since, and the outline has to match what is there
                    // now rather than what was there then.
                    double bw = 240, bh = 80;
                    var sz = target.SelectSingleNode("*[local-name()='Size']");
                    if (sz != null) { Dbl(sz, "width", out bw); Dbl(sz, "height", out bh); }
                    fresh.InnerText = BorderPng(bw, bh);
                }
                else
                {
                    fresh.InnerText = OpaquePng;
                }
                target.AppendChild(fresh);

                UpdatePage(app, doc);
                Log("TogglePeek: " + id + " -> " + (covered ? "peeked" : "covered"));
            }
            catch (Exception ex) { Log("TogglePeek FAILED: " + ex.Message); }
            finally { Release(app); }
        }

        // ---- Remove tape --------------------------------------------------------------------

        /// <summary>
        /// Remove only the tape the user picked.
        ///
        /// Two different gestures have to work, and they produce DIFFERENT XML:
        ///
        ///   Clicking a tape strip selects an object. Our overlay IS a one:Image, so OneNote hands
        ///   it back selected="all" and there is nothing to work out.
        ///
        ///   Putting the CARET in taped text selects nothing. Per the schema, "all" means the object
        ///   is selected and "partial" means it merely CONTAINS a selection - so a zero-length caret
        ///   can never produce an "all" anywhere on the page. Requiring one (which we did) meant
        ///   clicking inside black text and pressing Remove got "select some tape first", which is
        ///   both wrong and impossible to act on: you cannot see what you are selecting.
        /// </summary>
        public void RemoveTape(IRibbonControl control)
        {
            RemoveTapeScoped(selectionOnly: true);
        }

        /// <summary>
        /// Remove every tape on the page. This is the dangerous one: it can undo a study session's
        /// worth of work in a click, and there is no undo that reaches back through a sync. So it
        /// counts first, says exactly how many, and defaults the dialog to No.
        /// </summary>
        public void RemoveAllTape(IRibbonControl control)
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

                int n = CountTape(doc);
                if (n == 0)
                {
                    Log("RemoveAllTape: nothing to remove");
                    Tell(app, "There is no RecallTape on this page.", "RecallTape");
                    return;
                }

                string question = n == 1
                    ? "Remove the 1 tape strip on this page?"
                    : "Remove all " + n + " tape strips on this page?";

                if (!Confirm(app, question + "\n\nWhat is underneath is not affected.", "Remove all tape"))
                {
                    Log("RemoveAllTape: cancelled by user (" + n + " would have been removed)");
                    return;
                }

                RemoveTapeScoped(selectionOnly: false);
            }
            catch (Exception ex) { Log("RemoveAllTape FAILED: " + ex.Message); }
            finally { Release(app); }
        }

        /// <summary>Count tape on the page: overlay images plus taped text runs.</summary>
        private static int CountTape(XmlDocument doc)
        {
            int n = 0;
            foreach (XmlNode img in doc.SelectNodes("//*[local-name()='Image']"))
            {
                var alt = img.Attributes["alt"];
                if (alt != null && alt.Value.StartsWith(TapePrefix, StringComparison.OrdinalIgnoreCase)) n++;
            }
            foreach (XmlNode t in doc.SelectNodes("//*[local-name()='T']"))
            {
                string c = t.InnerText;
                if (HasTape(c)) n++;
            }
            return n;
        }

        private void RemoveTapeScoped(bool selectionOnly)
        {
            ON.Application app = null;
            try
            {
                app = new ON.Application();
                string pageId = app.Windows.CurrentWindow.CurrentPageId;

                // piSelection when we care which one they picked; piBasic is enough otherwise.
                string xml;
                app.GetPageContent(pageId, out xml,
                    selectionOnly ? ON.PageInfo.piSelection : ON.PageInfo.piBasic, ON.XMLSchema.xs2013);

                var doc = new XmlDocument();
                doc.LoadXml(xml);

                // What counts as "the user picked this"?
                //
                // A caret is NOT "nothing selected", which is what we assumed twice and were wrong
                // about twice. Measured from a live page with the cursor inside a taped word, OneNote
                // SPLITS the run at the insertion point and marks a ZERO-LENGTH one:T selected="all":
                //
                //   <one:T selected="none"><![CDATA[<span ...>WHOOMP!</span>]]></one:T>
                //   <one:T selected="all"><![CDATA[]]></one:T>          <-- the caret, empty
                //   <one:T selected="none"><![CDATA[<span ...> DERE IT IS!</span>]]></one:T>
                //
                // So there IS an "all" on the page, it just holds no text and neither half of the tape
                // is marked at all. Testing "is anything selected" found that empty run, took the
                // selection path, and matched nothing - the caret branch never even ran.
                //
                // Note the split: one tape came back as TWO taped runs. That is why the caret scope is
                // the whole one:OE. Scoping to either run would strip half a tape.
                Func<XmlNode, bool> inScope;
                if (!selectionOnly)
                {
                    inScope = delegate { return true; };
                }
                else
                {
                    XmlNode realSelection = null;
                    XmlNode caretMark = null;
                    foreach (XmlNode n in doc.SelectNodes("//*[@selected='all']"))
                    {
                        if (n.LocalName == "T" && string.IsNullOrEmpty(n.InnerText))
                        {
                            if (caretMark == null) caretMark = n;
                        }
                        else if (realSelection == null) realSelection = n;
                    }

                    if (realSelection != null)
                    {
                        inScope = IsSelected;
                    }
                    else
                    {
                        XmlNode scope = caretMark != null ? NearestOE(caretMark) : CaretScope(doc);
                        Log("RemoveTape: caret, no selection; scope = "
                            + (scope == null ? "NONE - OneNote marked nothing" : scope.LocalName));
                        if (scope == null) inScope = delegate { return false; };
                        else inScope = node => IsWithin(node, scope);
                    }
                }

                // 1. Text runs: unwrap our span, restoring the original CDATA byte for byte.
                int runs = 0;
                foreach (XmlNode t in doc.SelectNodes("//*[local-name()='T']"))
                {
                    string content = t.InnerText;
                    if (string.IsNullOrEmpty(content)) continue;
                    if (!HasTape(content)) continue;
                    if (!inScope(t)) continue;

                    SetCData(doc, t, StripTape(content));
                    runs++;
                }
                if (runs > 0) UpdatePage(app, doc);

                // 2. Overlays: delete by objectID, so we remove the element we own and nothing else.
                int boxes = 0;
                int failed = 0;
                foreach (XmlNode n in doc.SelectNodes("//*[local-name()='Image']"))
                {
                    var alt = n.Attributes["alt"];
                    var oid = n.Attributes["objectID"];
                    if (alt == null || oid == null) continue;
                    if (!alt.Value.StartsWith(TapePrefix, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!inScope(n)) continue;

                    // Fresh timestamp per delete: the previous delete moved it.
                    try
                    {
                        app.DeletePageContent(pageId, oid.Value, CurrentTimestamp(app, pageId), false);
                        boxes++;
                    }
                    catch (Exception ex)
                    {
                        // Report what we managed rather than abandoning the rest of the page. One
                        // stubborn box should not strand the other nineteen.
                        failed++;
                        Log("  could not remove overlay " + oid.Value + ": " + ex.Message);
                    }
                }

                string what = selectionOnly ? "RemoveTape" : "RemoveAllTape";
                if (selectionOnly && runs == 0 && boxes == 0)
                {
                    Log(what + ": nothing selected that is taped");
                    Tell(app, "No tape here.\n\n"
                            + "Put the cursor in taped text, or click a tape strip, then press "
                            + "Remove Tape.\n\n"
                            + "To clear the whole page, use Remove All.", "RecallTape");
                    return;
                }

                int foreign = 0;
                if (!selectionOnly)
                {
                    foreach (XmlNode t in doc.SelectNodes("//*[local-name()='T']"))
                    {
                        string c = t.InnerText;
                        if (string.IsNullOrEmpty(c)) continue;
                        if (HasTape(c)) continue;
                        if (Regex.IsMatch(c, @"mso-highlight|background\s*:", RegexOptions.IgnoreCase)) foreign++;
                    }
                }

                Log(what + ": restored " + runs + " text run(s), removed " + boxes + " overlay box(es)"
                    + (failed > 0 ? "; " + failed + " overlay(s) FAILED" : "")
                    + (foreign > 0 ? "; left " + foreign + " run(s) alone - highlighted, but not RecallTape's" : ""));
            }
            catch (Exception ex) { Log("RemoveTape FAILED: " + ex.Message); }
            finally { Release(app); }
        }

        /// <summary>
        /// Where is the caret? The deepest element marked "partial", climbed to its paragraph.
        ///
        /// The climb matters. OneNote splits a one:T at the insertion point, so a single taped run
        /// with a caret in the middle of it comes back as TWO taped runs. Scoping to the one that
        /// happens to be marked would strip half a tape and leave the other half. The one:OE is the
        /// smallest unit guaranteed to hold a whole tape.
        ///
        /// Cost of that choice: several separately taped words in ONE paragraph come off together.
        /// Accepted for now - a half-removed tape is a visible defect, removing a neighbour is not.
        /// </summary>
        private static XmlNode CaretScope(XmlDocument doc)
        {
            XmlNode deepest = null;
            int best = -1;
            foreach (XmlNode n in doc.SelectNodes("//*[@selected='partial']"))
            {
                int d = 0;
                for (XmlNode p = n; p != null; p = p.ParentNode) d++;
                if (d > best) { best = d; deepest = n; }
            }
            if (deepest == null) return null;

            return NearestOE(deepest) ?? deepest;
        }

        /// <summary>The paragraph containing this node - the smallest unit that holds a whole tape.</summary>
        private static XmlNode NearestOE(XmlNode n)
        {
            for (XmlNode p = n; p != null; p = p.ParentNode)
                if (p.LocalName == "OE") return p;
            return null;
        }

        /// <summary>Is this node the scope, or inside it?</summary>
        private static bool IsWithin(XmlNode n, XmlNode scope)
        {
            for (XmlNode p = n; p != null; p = p.ParentNode)
                if (ReferenceEquals(p, scope)) return true;
            return false;
        }

        /// <summary>selected="all" on the node, or on anything inside it.</summary>
        private static bool IsSelected(XmlNode n)
        {
            var a = n.Attributes == null ? null : n.Attributes["selected"];
            if (a != null && a.Value == "all") return true;
            return n.SelectSingleNode(".//*[@selected='all']") != null;
        }

        // ---- Talking to the user ------------------------------------------------------------
        //
        // P/Invoke rather than WinForms: this is two dialogs, and pulling System.Windows.Forms into
        // the surrogate would mean owning its DPI and message-loop behaviour for no benefit.

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = false)]
        private static extern int MessageBoxW(IntPtr hWnd, string text, string caption, uint type);

        private const uint MB_YESNO = 0x00000004;
        private const uint MB_ICONWARNING = 0x00000030;
        private const uint MB_ICONINFORMATION = 0x00000040;
        private const uint MB_DEFBUTTON2 = 0x00000100;   // default to No on a destructive question
        private const int IDYES = 6;

        private static IntPtr OneNoteWindow(ON.Application app)
        {
            // Parent to OneNote so the dialog cannot appear behind it - a modal hidden behind the
            // window it belongs to looks exactly like a hang.
            try { return new IntPtr((long)app.Windows.CurrentWindow.WindowHandle); }
            catch { return IntPtr.Zero; }
        }

        private static bool Confirm(ON.Application app, string text, string caption)
        {
            return MessageBoxW(OneNoteWindow(app), text, caption,
                               MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2) == IDYES;
        }

        private static void Tell(ON.Application app, string text, string caption)
        {
            MessageBoxW(OneNoteWindow(app), text, caption, MB_ICONINFORMATION);
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
        /// <summary>
        /// The page's CURRENT timestamp, re-read from OneNote.
        ///
        /// Every mutating call bumps lastModifiedTime - UpdatePageContent and each individual
        /// DeletePageContent. We hand that timestamp back as a concurrency check with force:false, so
        /// the SECOND write in any batch was always going to fail 0x80042010 against a value captured
        /// up front.
        ///
        /// Remove All on a page holding both text tape and an overlay box did exactly that: the text
        /// update landed, the first overlay delete threw, the catch aborted everything after it, and
        /// the user had to press the button a second time to finish the job. It looked like the two
        /// kinds of tape were fighting. They were not - the first write invalidated the second.
        /// The count in the dialog was right because counting happens before any write.
        ///
        /// Re-reading before each write is the honest fix. Passing DateTime.MinValue with force:true
        /// would also make the symptom go away, by disabling the check that stops us overwriting an
        /// edit we never saw.
        /// </summary>
        private static DateTime CurrentTimestamp(ON.Application app, string pageId)
        {
            string xml;
            app.GetPageContent(pageId, out xml, ON.PageInfo.piBasic, ON.XMLSchema.xs2013);
            var fresh = new XmlDocument();
            fresh.LoadXml(xml);
            return LastModified(fresh);
        }

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
