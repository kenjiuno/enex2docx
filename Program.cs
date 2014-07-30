using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace enex2docx {
    class Program {
        class ResUt {
            class HUt {
                public static string GetStr(byte[] bin) {
                    String s = "";
                    foreach (byte b in bin) s += b.ToString("x2");
                    return s;
                }
            }

            public static Res1[] Parse(XElement elnote) {
                List<Res1> al = new List<Res1>();
                foreach (var elres in elnote.Elements("resource")) {
                    foreach (var eldata in elres.Elements("data")) {
                        if (eldata.Attribute("encoding").Value == "base64") {
                            byte[] bin = Convert.FromBase64String(eldata.Value);
                            al.Add(new Res1 { si = new MemoryStream(bin, false), elres = elres, md5 = HUt.GetStr(MD5.Create().ComputeHash(bin)) });
                        }
                        break;
                    }
                }
                return al.ToArray();
            }
        }
        class Res1 {
            public MemoryStream si;
            public String md5;

            public XElement elres;

            public String mime { get { return elres.Element("mime").Value; } }
            public String width { get { return elres.Element("width").Value; } }
            public String height { get { return elres.Element("height").Value; } }
        }

        static void Main(string[] args) {
            if (args.Length < 2) {
                helpYa();
                Environment.ExitCode = 1;
                return;
            }
            String dir = args[1];
            XDocument enex = XDocument.Load(args[0]);
            var notes = enex.Elements("en-export").Elements("note");
            Fnno fnno = new Fnno();
            foreach (var note in notes) {
                var title = note.Element("title").Value;
                var content = note.Element("content").Value;

                var created = note.Element("created").Value;
                DateTime dt;
                if (!DateTime.TryParse(created, out dt)) {
                    Match M = Regex.Match(created, "^(?<y>\\d{4})(?<M>\\d{2})(?<d>\\d{2})T(?<H>\\d{2})(?<m>\\d{2})(?<s>\\d{2})Z");
                    if (M.Success) {
                        dt = new DateTime(
                            int.Parse(M.Groups["y"].Value),
                            int.Parse(M.Groups["M"].Value),
                            int.Parse(M.Groups["d"].Value),
                            int.Parse(M.Groups["H"].Value),
                            int.Parse(M.Groups["m"].Value),
                            int.Parse(M.Groups["s"].Value)
                            );
                    }
                    else dt = DateTime.Now;
                }

                content = content.Replace("&nbsp;", "&#xa0;");

                XDocument ennote = XDocument.Parse(content);
                var elroot = ennote.Element("en-note");

                String fpdocx = Path.Combine(dir, fnno.Next(title) + ".docx");
                Res1[] reses = ResUt.Parse(note);
                using (DocX doc = DocX.Create(fpdocx)) {
                    WDoc wd = new WDoc { doc = doc };
                    //new XUt { }.Walk(elroot.Nodes());
                    new WUt { wd = wd, reses = reses }.Walk(elroot.Nodes(), 0);
                    doc.Save();
                }

                File.SetLastWriteTimeUtc(fpdocx, dt);
                //break;
            }
        }

        class WDoc {
            public DocX doc;

            Paragraph p;

            List<Paragraph> al = new List<Paragraph>();

            internal void Str(string t) {
                var p = Para().Append(t);
                foreach (var kv in dict) kv.Value.Add(p);
            }

            private Paragraph LPara() {
                if (p == null)
                    p = doc.InsertParagraph();
                return p;
            }

            private Paragraph Para() {
                if (al.Count != 0)
                    return al[al.Count - 1];
                return LPara();
            }

            internal void Newl() {
                Para().AppendLine();
            }

            internal void AddPic(MemoryStream si, int cx, int cy) {
                var image = doc.AddImage(si);
                float maxx = 600;
                if (cx > maxx) {
                    float s = maxx / cx;
                    cx = (int)(cx * s);
                    cy = (int)(cy * s);
                }
                var pic = image.CreatePicture(
                    cx,
                    cy
                    );
                Para().AppendPicture(pic);
            }

            internal void Newp() {
                p = doc.InsertParagraph();
            }

            int i = 0;

            internal string Capture() {
                i++;
                string k = "#" + i;
                dict.Add(k, new List<Paragraph>());
                return k;
            }

            SortedDictionary<string, List<Paragraph>> dict = new SortedDictionary<string, List<Paragraph>>();

            internal Paragraph[] Release(string k) {
                var v = dict[k];
                dict.Remove(k);
                return v.ToArray();
            }

            internal void Style(Paragraph[] alp, string style) {
                System.Collections.SortedList dict = new System.Collections.SortedList();
                foreach (string row in style.Split(';')) {
                    string[] cols = row.Split(new char[] { ':' }, 2);
                    if (cols.Length == 2) {
                        dict[cols[0].Trim()] = cols[1].Trim();
                    }
                }

                //foreach (var kv in dict) System.Diagnostics.Debug.WriteLine("# " + kv.Key + ": " + kv.Value);

                {
                    Match M = Regex.Match("" + dict["color"], "rgb\\s*\\(\\s*(?<r>\\d+)\\s*,\\s*(?<g>\\d+)\\s*,\\s*(?<b>\\d+)\\s*\\)");
                    if (M.Success) {
                        Color c = Color.FromArgb(int.Parse(M.Groups["r"].Value), int.Parse(M.Groups["g"].Value), int.Parse(M.Groups["b"].Value));
                        foreach (var p in alp) p.Color(c);
                    }
                }
                {
                    Match M = Regex.Match("" + dict["background-color"], "rgb\\s*\\(\\s*(?<r>\\d+)\\s*,\\s*(?<g>\\d+)\\s*,\\s*(?<b>\\d+)\\s*\\)");
                    if (M.Success) {
                        Color c = Color.FromArgb(int.Parse(M.Groups["r"].Value), int.Parse(M.Groups["g"].Value), int.Parse(M.Groups["b"].Value));
                        foreach (var p in alp) p.Highlight(Highlight.yellow);
                    }
                }
            }

            internal void Link(string str, string href) {
                Para().AppendHyperlink(doc.AddHyperlink(str, new Uri(href))).Color(Color.Blue).UnderlineColor(Color.Blue);
            }

            List list;

            internal void L(bool bull) {
                list = doc.AddList(" ", 0, bull ? ListItemType.Bulleted : ListItemType.Numbered);
            }
            internal void LEnd() {
                doc.InsertList(list);
            }
            internal void LI() {
                if (lind != 0) {
                    list = doc.AddListItem(list, " ");
                    p = list.Items[list.Items.Count - 1];
                }
                else {
                    p = list.Items[lind];
                }
                lind++;
            }

            int lind = 0;


            internal void Blockquote() {
                //LPara().IndentationBefore += 0.4f;
            }

            Table tbl;
            int tbly = 0, tblx = 0;
            Row tr;

            internal void NewTbl() {
                tbl = doc.AddTable(1, 1);
                tbly = 0;
            }

            internal void Tr() {
                while (tbly >= tbl.RowCount) tbl.InsertRow();
                tr = tbl.Rows[tbly];
                tblx = 0;
                tbly++;
            }

            List<Paragraph> tpused = new List<Paragraph>();

            internal void Td() {
                while (tblx >= tbl.ColumnCount) tbl.InsertColumn();
                var p = tr.Cells[tblx].Paragraphs.First();
                tpused.Add(p);
                al.Add(p);
                tblx++;
            }

            internal void AddTbl() {
                doc.InsertTable(tbl);
                tbl = null;
                foreach (Paragraph p in tpused) al.Remove(p);
                tpused.Clear();
            }

            internal void Horzline() {
                Newp();
                Str("---");
            }
        }

        class WUt {
            public WDoc wd;
            public Res1[] reses;

            public void Walk(IEnumerable<XNode> nodes, int depth) {
                int y = 0;
                foreach (XNode xn in nodes) {
                    y++;
                    if (depth == 0 && y != 0) wd.Newp();
                    XElement el = xn as XElement;
                    XText xt = xn as XText;
                    if (el != null) {
                        if (el.Name == "div" || el.Name == "p") {
                            String k = wd.Capture();
                            Walk(el.Nodes(), depth + 1);
                            wd.Style(wd.Release(k), (el.Attribute("style") ?? new XAttribute("style", "")).Value);
                        }
                        else if (el.Name == "br") {
                            wd.Newl();
                        }
                        else if (el.Name == "b") {
                            String k = wd.Capture();
                            Walk(el.Nodes(), depth + 1);
                            foreach (var p in wd.Release(k)) p.Bold();
                        }
                        else if (el.Name == "u") {
                            String k = wd.Capture();
                            Walk(el.Nodes(), depth + 1);
                            foreach (var p in wd.Release(k)) p.UnderlineColor(Color.Black);
                        }
                        else if (el.Name == "i") {
                            String k = wd.Capture();
                            Walk(el.Nodes(), depth + 1);
                            foreach (var p in wd.Release(k)) p.Italic();
                        }
                        else if (el.Name == "en-media") {
                            String ty2 = el.Attribute("type").Value;
                            if (ty2 == "image/jpeg") {
                                String hash = el.Attribute("hash").Value;
                                var res1 = reses.FirstOrDefault(q => q.md5 == hash);
                                if (res1 != null) {
                                    wd.AddPic(res1.si, int.Parse(res1.height), int.Parse(res1.width));
                                }
                            }
                            else {
                                wd.Str("[未対応の形式: " + ty2 + "]");
                            }
                            wd.Newl();
                        }
                        else if (el.Name == "span") {
                            String k = wd.Capture();
                            Walk(el.Nodes(), depth + 1);
                            wd.Style(wd.Release(k), Utatt.Get(el, "style"));
                        }
                        else if (el.Name == "font") {
                            String k = wd.Capture();
                            Walk(el.Nodes(), depth + 1);
                            foreach (var p in wd.Release(k)) {
                                String face = Utatt.Get(el, "face");
                                if (face.Length != 0) p.Font(new FontFamily(face));
                                String color = Utatt.Get(el, "color");
                                if (color.Length != 0) {
                                    Match M = Regex.Match(color, "^(?<R>[0-9a-f]{2})(?<G>[0-9a-f]{2})(?<B>[0-9a-f]{2})$", RegexOptions.IgnoreCase);
                                    if (M.Success) {
                                        p.Color(Color.FromArgb(
                                            Convert.ToByte(M.Groups["R"].Value, 16),
                                            Convert.ToByte(M.Groups["G"].Value, 16),
                                            Convert.ToByte(M.Groups["B"].Value, 16)
                                            ));
                                    }
                                }
                            }
                        }
                        else if (el.Name == "strike") {
                            String k = wd.Capture();
                            Walk(el.Nodes(), depth + 1);
                            foreach (var p in wd.Release(k)) p.StrikeThrough(StrikeThrough.strike);
                        }
                        else if (el.Name == "ul") {
                            wd.L(true);
                            Walk(el.Nodes(), depth + 1);
                            wd.LEnd();
                        }
                        else if (el.Name == "ol") {
                            wd.L(false);
                            Walk(el.Nodes(), depth + 1);
                            wd.LEnd();
                        }
                        else if (el.Name == "li") {
                            wd.LI();
                            Walk(el.Nodes(), depth + 1);
                        }
                        else if (el.Name == "en-todo") {
                            String _checked = Utatt.Get(el, "checked");
                            if (_checked == "true" || _checked == "1") {
                                wd.Str("[X]");
                            }
                            else {
                                wd.Str("[ ]");
                            }
                        }
                        else if (el.Name == "blockquote") {
                            wd.Blockquote();
                            Walk(el.Nodes(), depth + 1);
                        }
                        else if (el.Name == "table") {
                            wd.NewTbl();
                            Walk(el.Nodes(), depth + 1);
                            wd.AddTbl();
                        }
                        else if (el.Name == "tbody") {
                            Walk(el.Nodes(), depth + 1);
                        }
                        else if (el.Name == "tr") {
                            wd.Tr();
                            Walk(el.Nodes(), depth + 1);
                        }
                        else if (el.Name == "td") {
                            wd.Td();
                            Walk(el.Nodes(), depth + 1);
                        }
                        else if (el.Name == "hr") {
                            wd.Horzline();
                        }
                        else if (el.Name == "a") {
                            String href = Utatt.Get(el, "href");
                            Walk_a(href, el);
                        }
                        else if (el.Name == "en-crypt") {
                            wd.Str("[未対応の形式: 暗号化されたブロック]");
                            wd.Newl();
                        }
                        else throw new NotSupportedException("" + el.Name);
                    }
                    else if (xt != null) {
                        wd.Str(xt.Value);
                    }
                    else throw new NotSupportedException("" + xn);
                }
            }

            private void Walk_a(string href, XElement el) {
                StringWriter wr = new StringWriter();
                foreach (XNode xn in el.Nodes()) {
                    XText xt = xn as XText;
                    if (xt != null) {
                        wr.Write(xt.Value);
                    }
                }
                wd.Link(wr.ToString(), href);
            }

        }

        class Utatt {
            public static string Get(XElement el, string a) {
                var xa = el.Attribute(a);
                return (xa != null) ? xa.Value : "";
            }
        }

        class Fnno {
            SortedDictionary<String, String> dict = new SortedDictionary<string, string>();

            public string Next(String fnwe0) {
                for (int x = 0; ; x++) {
                    String fnwe = fnwe0 + ((x == 0) ? "" : "~" + x);
                    if (dict.ContainsKey(fnwe)) continue;
                    dict[fnwe] = null;
                    return fnwe;
                }
            }
        }

        static void helpYa() {
            Console.Error.WriteLine("enex2docx <a.enex> <out folder>");
        }




        class XUt {
            XDocument xmlo = new XDocument();
            XElement elroot;

            public XUt Walk(IEnumerable<XNode> nodes) {
                xmlo.Add(elroot = new XElement("note"));
                Walk(nodes, null, 0);
                return this;
            }

            private void Walk(IEnumerable<XNode> nodes, XElement Welline, int d) {
                foreach (XNode xn in nodes) {
                    if (Welline == null) {
                        Welline = new XElement("line");
                        elroot.Add(Welline);
                    }

                    XElement el = xn as XElement;
                    XText xt = xn as XText;
                    if (el != null) {
                        if (el.Name == "en-todo") {
                            Welline.Add(new XElement(el));
                            continue;
                        }
                        XElement Wel2;
                        Welline.Add(Wel2 = ApplyStyle(new XElement("style"), el.Attribute("style")));

                        if (el.Name == "b") {
                            Wel2.SetAttributeValue("b", "1");
                        }
                        else if (el.Name == "i") {
                            Wel2.SetAttributeValue("i", "1");
                        }
                        else if (el.Name == "u") {
                            Wel2.SetAttributeValue("u", "1");
                        }
                        else if (el.Name == "strike") {
                            Wel2.SetAttributeValue("strike", "1");
                        }
                        else if (el.Name == "div" || el.Name == "span") {

                        }
                        else if (el.Name == "ul") {
                            Wel2.Name = "ul";
                        }
                        else if (el.Name == "ol") {
                            Wel2.Name = "ul";
                        }
                        else if (el.Name == "li") {
                            Wel2.Name = "li";
                        }
                        else if (el.Name == "br") {
                            Wel2.Name = "br";
                        }
                        else if (el.Name == "table") {
                            Wel2.Name = "table";
                        }
                        else if (el.Name == "tbody") {
                            Wel2.Name = "tbody";
                        }
                        else if (el.Name == "tr") {
                            Wel2.Name = "tr";
                        }
                        else if (el.Name == "td") {
                            Wel2.Name = "td";
                        }
                        else if (el.Name == "hr") {
                            Wel2.Name = "hr";
                        }
                        else if (el.Name == "a") {
                            Wel2.Name = "a";
                        }
                        else if (el.Name == "blockquote") {
                            Wel2.SetAttributeValue("indent", "1");
                        }
                        else if (el.Name == "font") {
                            String face = Utatt.Get(el, "face");
                            if (face.Length != 0) Wel2.SetAttributeValue("face", face);
                            String color = Utatt.Get(el, "color");
                            if (color.Length != 0) {
                                Match M = Regex.Match(color, "^(?<R>[0-9a-f]{2})(?<G>[0-9a-f]{2})(?<B>[0-9a-f]{2})$", RegexOptions.IgnoreCase);
                                if (M.Success) {
                                    Wel2.SetAttributeValue("fg", (Color.FromArgb(
                                        Convert.ToByte(M.Groups["R"].Value, 16),
                                        Convert.ToByte(M.Groups["G"].Value, 16),
                                        Convert.ToByte(M.Groups["B"].Value, 16)
                                        )));
                                }
                            }

                        }
                        else throw new NotSupportedException("" + el);

                        Walk(el.Nodes(), Wel2, d + 1);
                    }
                    else if (xt != null) {
                        Welline.Add(new XElement("text", new XText(xt)));
                    }
                    else throw new NotSupportedException("" + xn);
                }
            }

            private XElement ApplyStyle(XElement el, XAttribute xaStyle) {
                if (xaStyle != null) {
                    String style = xaStyle.Value;

                    System.Collections.SortedList dict = new System.Collections.SortedList();
                    foreach (string row in style.Split(';')) {
                        string[] cols = row.Split(new char[] { ':' }, 2);
                        if (cols.Length == 2) {
                            dict[cols[0].Trim()] = cols[1].Trim();
                        }
                    }

                    //foreach (var kv in dict) System.Diagnostics.Debug.WriteLine("# " + kv.Key + ": " + kv.Value);

                    {
                        Match M = Regex.Match("" + dict["color"], "rgb\\s*\\(\\s*(?<r>\\d+)\\s*,\\s*(?<g>\\d+)\\s*,\\s*(?<b>\\d+)\\s*\\)");
                        if (M.Success) {
                            Color c = Color.FromArgb(int.Parse(M.Groups["r"].Value), int.Parse(M.Groups["g"].Value), int.Parse(M.Groups["b"].Value));
                            el.SetAttributeValue("fg", c.ToArgb());
                        }
                    }
                    {
                        Match M = Regex.Match("" + dict["background-color"], "rgb\\s*\\(\\s*(?<r>\\d+)\\s*,\\s*(?<g>\\d+)\\s*,\\s*(?<b>\\d+)\\s*\\)");
                        if (M.Success) {
                            Color c = Color.FromArgb(int.Parse(M.Groups["r"].Value), int.Parse(M.Groups["g"].Value), int.Parse(M.Groups["b"].Value));
                            el.SetAttributeValue("bg", c.ToArgb());
                        }
                    }
                }
                return el;
            }
        }
    }
}
