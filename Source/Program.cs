// PuTTY Session Builder 1.0.0 - MIT License
// Windows 10/11, .NET Framework 4.8. No Office automation, no network access.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using ExcelDataReader;

namespace PuttySessionBuilder
{
    public sealed class Entry
    {
        public string[] Cells;
        public int Line;
        public Entry(string[] cells, int line) { Cells = cells; Line = line; }
    }
    public sealed class Session
    {
        public string Name, Key;
        public Dictionary<string, object> Values = new Dictionary<string, object>();
    }
    public static class Core
    {
        public const string RegistryPath = @"Software\SimonTatham\PuTTY\Sessions";
        public static readonly string[] Headers = {
            "Anzeigename", "Hostname/IP", "Port", "Font", "Fontvariante", "Fontgröße",
            "Fontquality", "Keepalive", "Keepalive Rate", "Autologin Username", "X11 Forwarding",
            "Protokoll", "TCP Keepalive", "X11 Display", "SSH Keyfile", "Scrollback" };
        public static readonly string[] Example = {
            "BEISPIEL - wird übersprungen", "192.0.2.10", "22", "Consolas", "Normal", "11",
            "ClearType", "TRUE", "5", "admin", "FALSE", "ssh", "FALSE", "localhost:0", "", "2000" };
        public static readonly string[] Quality = { "Standard", "Antialiasing", "Keine Glättung", "ClearType" };
        public static readonly string[] Styles = { "Normal", "Fett" };
        public static readonly string[] FontSuggestions = {
            "Consolas", "Courier New", "Lucida Console", "Cascadia Mono", "Cascadia Code",
            "DejaVu Sans Mono", "Liberation Mono", "JetBrains Mono", "Fira Code", "Source Code Pro" };

        public static string Text(object value)
        {
            if (value == null || value == DBNull.Value) return "";
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            return Convert.ToString(value, CultureInfo.InvariantCulture).Trim();
        }
        public static string Canon(string value)
        {
            return (value ?? "").Trim().ToLowerInvariant().Replace("ä", "ae").Replace("ö", "oe")
                .Replace("ü", "ue").Replace("ß", "ss").Replace(" ", "").Replace("_", "").Replace("-", "");
        }
        public static int Integer(string s, int def, int min, int max, string field)
        {
            if (String.IsNullOrWhiteSpace(s)) return def;
            int n;
            if (!Int32.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) || n < min || n > max)
                throw new Exception(field + ": ganze Zahl von " + min + " bis " + max + " erforderlich.");
            return n;
        }
        public static bool Boolean(string s, bool def, string field)
        {
            if (String.IsNullOrWhiteSpace(s)) return def;
            switch (Canon(s)) {
                case "true": case "wahr": case "ja": case "yes": case "1": return true;
                case "false": case "falsch": case "nein": case "no": case "0": return false;
                default: throw new Exception(field + ": TRUE/FALSE, WAHR/FALSCH, JA/NEIN oder 1/0 erforderlich.");
            }
        }
        public static string KeyName(string name, Encoding encoding)
        {
            // PuTTY uses byte-wise registry escaping: space, backslash, *, ?, %, controls,
            // non-ASCII bytes and a leading dot. Use the Windows ANSI code page (Reg*A).
            Encoding strict = Encoding.GetEncoding(encoding.CodePage, EncoderFallback.ExceptionFallback,
                DecoderFallback.ExceptionFallback);
            byte[] bytes;
            try { bytes = strict.GetBytes(name); }
            catch (EncoderFallbackException) {
                throw new Exception("Anzeigename enthält Zeichen außerhalb der Windows-Systemcodepage. Bitte einen anderen Namen verwenden.");
            }
            StringBuilder b = new StringBuilder();
            for (int i = 0; i < bytes.Length; i++) {
                byte c = bytes[i];
                if (c <= 32 || c >= 127 || c == '\\' || c == '*' || c == '?' || c == '%' || (i == 0 && c == '.'))
                    b.Append('%').Append(c.ToString("X2"));
                else b.Append((char)c);
            }
            return b.ToString();
        }
        public static Session Parse(Entry row, ICollection<string> availableFonts, Encoding encoding)
        {
            string[] c = row.Cells;
            if (c.Length != Headers.Length) throw new Exception("Unerwartete Spaltenanzahl.");
            string name = c[0].Trim(), host = c[1].Trim();
            if (name.Length == 0) throw new Exception("Anzeigename fehlt.");
            if (name.Length > 180 || name.Any(Char.IsControl)) throw new Exception("Anzeigename: maximal 180 Zeichen, keine Steuerzeichen.");
            if (name.Equals("Default Settings", StringComparison.OrdinalIgnoreCase))
                throw new Exception("Default Settings ist reserviert und wird nicht überschrieben.");
            if (host.Length == 0 || host.Any(Char.IsWhiteSpace) || host.Any(Char.IsControl) || host.Contains("://") || host.Contains("/") || host.Contains("\\"))
                throw new Exception("Hostname/IP fehlt oder enthält Leerzeichen, URL oder Pfad.");
            if (host.StartsWith("[", StringComparison.Ordinal) && host.EndsWith("]", StringComparison.Ordinal))
                host = host.Substring(1, host.Length - 2);
            if (host.Contains(":")) {
                System.Net.IPAddress ip;
                if (!System.Net.IPAddress.TryParse(host, out ip))
                    throw new Exception("Hostname/IP: IPv6-Adresse ungültig. Den Port separat eintragen.");
            }
            string protocol = String.IsNullOrWhiteSpace(c[11]) ? "ssh" : c[11].Trim().ToLowerInvariant();
            if (protocol != "ssh" && protocol != "telnet" && protocol != "raw") throw new Exception("Protokoll: ssh, telnet oder raw.");
            int port = Integer(c[2], protocol == "telnet" ? 23 : 22, 1, 65535, "Port");
            string font = String.IsNullOrWhiteSpace(c[3]) ? "Consolas" : c[3].Trim();
            if (availableFonts != null && !availableFonts.Any(f => f.Equals(font, StringComparison.OrdinalIgnoreCase)))
                throw new Exception("Font '" + font + "' ist hier nicht als installierte Monospace-Schrift verfügbar. Im Dropdown auswählen.");
            bool bold;
            switch (Canon(c[4])) {
                case "": case "normal": case "regular": case "standard": case "0": bold = false; break;
                case "fett": case "bold": case "1": bold = true; break;
                default: throw new Exception("Fontvariante: Normal oder Fett. PuTTY speichert keine kursive Grundschrift.");
            }
            int height = Integer(c[5], 11, 1, 96, "Fontgröße (pt)");
            int quality;
            switch (Canon(c[6])) {
                case "": case "standard": case "default": case "0": quality = 0; break;
                case "antialiasing": case "antialiased": case "1": quality = 1; break;
                case "keineglaettung": case "nonantialiased": case "2": quality = 2; break;
                case "cleartype": case "3": quality = 3; break;
                default: throw new Exception("Fontquality: Standard, Keine Glättung, Antialiasing oder ClearType (0–3).");
            }
            bool keep = Boolean(c[7], true, "Keepalive");
            int rate = Integer(c[8], 5, keep ? 1 : 0, 86400, "Keepalive Rate (Sekunden)");
            bool x11 = Boolean(c[10], false, "X11 Forwarding");
            bool tcp = Boolean(c[12], false, "TCP Keepalive");
            if (x11 && protocol != "ssh") throw new Exception("X11 Forwarding benötigt das Protokoll ssh.");
            if (c[9].Any(Char.IsControl) || c[13].Any(Char.IsControl) || c[14].Any(Char.IsControl))
                throw new Exception("Username, Display und Keyfile dürfen keine Steuerzeichen enthalten.");
            Session s = new Session { Name = name, Key = KeyName(name, encoding) };
            if (s.Key.Length > 255) throw new Exception("Anzeigename ergibt einen zu langen Registry-Schlüssel (maximal 255 Zeichen nach PuTTY-Kodierung).");
            s.Values["Present"] = 1; s.Values["HostName"] = host; s.Values["PortNumber"] = port;
            s.Values["Protocol"] = protocol; s.Values["Font"] = font; s.Values["FontIsBold"] = bold ? 1 : 0;
            s.Values["FontCharSet"] = 0; s.Values["FontHeight"] = height; s.Values["FontQuality"] = quality;
            s.Values["PingInterval"] = keep ? rate / 60 : 0; s.Values["PingIntervalSecs"] = keep ? rate % 60 : 0;
            s.Values["UserName"] = c[9].Trim(); s.Values["UserNameFromEnvironment"] = 0;
            s.Values["X11Forward"] = x11 ? 1 : 0; s.Values["X11Display"] = c[13].Trim();
            s.Values["TCPKeepalives"] = tcp ? 1 : 0; s.Values["PublicKeyFile"] = c[14].Trim();
            s.Values["ScrollbackLines"] = Integer(c[15], 2000, 0, 1000000, "Scrollback");
            s.Values["LineCodePage"] = "UTF-8";
            return s;
        }
        public static List<Entry> Read(string file)
        {
            string ext = Path.GetExtension(file).ToLowerInvariant();
            if (ext != ".csv" && ext != ".xls" && ext != ".xlsx") throw new Exception("CSV, XLS oder XLSX auswählen.");
            using (FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (IExcelDataReader reader = ext == ".csv" ? ExcelReaderFactory.CreateCsvReader(stream,
                new ExcelReaderConfiguration { FallbackEncoding = Encoding.GetEncoding(1252), AutodetectSeparators = new[] {';', ',', '\t'} })
                : ExcelReaderFactory.CreateReader(stream)) {
                do {
                    // For workbooks prefer the Sessions sheet. A foreign workbook may use its first sheet.
                    if (ext != ".csv" && !reader.Name.Equals("Sessions", StringComparison.OrdinalIgnoreCase) && reader.ResultsCount > 1)
                        continue;
                    return ReadSheet(reader);
                } while (reader.NextResult());
                throw new Exception("Bei Dateien mit mehreren Tabellenblättern muss das Datenblatt 'Sessions' heißen.");
            }
        }
        static List<Entry> ReadSheet(IExcelDataReader r)
        {
            if (!r.Read()) throw new Exception("Die Datei ist leer.");
            Dictionary<string, int> cols = new Dictionary<string, int>();
            for (int i = 0; i < r.FieldCount; i++) {
                string h = Canon(Text(r.GetValue(i)));
                if (h.Length == 0) continue;
                if (cols.ContainsKey(h)) throw new Exception("Doppelte Spaltenüberschrift: " + Text(r.GetValue(i)));
                cols.Add(h, i);
            }
            for (int i = 0; i <= 10; i++) if (!cols.ContainsKey(Canon(Headers[i])))
                throw new Exception("Pflichtspalte fehlt: " + Headers[i] + ". Die mitgelieferte Vorlage benutzen.");
            List<Entry> rows = new List<Entry>(); int line = 1;
            while (r.Read()) {
                line++;
                string[] values = new string[Headers.Length];
                for (int i = 0; i < values.Length; i++) {
                    int col; values[i] = cols.TryGetValue(Canon(Headers[i]), out col) ? Text(r.GetValue(col)) : "";
                }
                // Keep physical row 2 even if emptied. It is always the example, never a session.
                if (line == 2 || values.Any(v => v.Length != 0)) rows.Add(new Entry(values, line));
                if (rows.Count > 100000) throw new Exception("Maximal 100.000 Zeilen pro Import.");
            }
            if (rows.Count == 0 || rows[0].Line != 2) throw new Exception("Beispielzeile fehlt. Zeile 2 unter den Überschriften muss erhalten bleiben.");
            return rows;
        }
        public static string CsvField(string value)
        {
            string s = value ?? "";
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }
        public static List<Session> ValidateSessions(IEnumerable<Entry> rows, ICollection<string> fonts, Encoding encoding, IDictionary<int,string> errors)
        {
            List<Session> sessions=new List<Session>();
            HashSet<string> names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(Entry row in rows) {
                if(row.Line==2) continue;
                try {
                    Session session=Parse(row,fonts,encoding);
                    if(!names.Add(session.Key)) throw new Exception("Doppelter Anzeigename (Windows ignoriert Groß-/Kleinschreibung).");
                    sessions.Add(session);
                } catch(Exception ex) { errors[row.Line]=ex.Message; }
            }
            return sessions;
        }
        public static void SaveCsv(string path, IEnumerable<Entry> rows)
        {
            using (StreamWriter w = new StreamWriter(path, false, new UTF8Encoding(true))) {
                w.WriteLine(String.Join(";", Headers.Select(CsvField)));
                foreach (Entry row in rows) w.WriteLine(String.Join(";", row.Cells.Select(CsvField)));
            }
        }
        public static List<string> InstalledFonts()
        {
            List<string> result = new List<string>();
            using (InstalledFontCollection fonts = new InstalledFontCollection())
            using (Bitmap bmp = new Bitmap(1, 1))
            using (Graphics g = Graphics.FromImage(bmp))
            using (StringFormat fmt = (StringFormat)StringFormat.GenericTypographic.Clone()) {
                fmt.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
                foreach (FontFamily ff in fonts.Families) {
                    try {
                        FontStyle style = ff.IsStyleAvailable(FontStyle.Regular) ? FontStyle.Regular : FontStyle.Bold;
                        using (Font font = new Font(ff, 12, style)) {
                            float a = g.MeasureString("iiiiiiii", font, 2000, fmt).Width;
                            float b = g.MeasureString("WWWWWWWW", font, 2000, fmt).Width;
                            float c = g.MeasureString("00000000", font, 2000, fmt).Width;
                            if (a > 0 && Math.Abs(a-b) < 0.1 && Math.Abs(a-c) < 0.1) result.Add(ff.Name);
                        }
                    } catch (ArgumentException) { }
                }
            }
            result.Sort(StringComparer.OrdinalIgnoreCase); return result;
        }
    }

    public static class RegistryImport
    {
        static string Esc(string s) { return s.Replace("\\", "\\\\").Replace("\"", "\\\""); }
        static void Hex(StringBuilder text, string prefix, byte[] data)
        {
            text.Append(prefix);
            for (int i=0; i<data.Length; i++) {
                if (i>0) text.Append(',');
                if (i>0 && i%24==0) text.Append("\\\r\n  ");
                text.Append(data[i].ToString("x2"));
            }
            text.Append("\r\n");
        }
        static void Snapshot(StringBuilder text, string path, RegistryKey key)
        {
            // Delete the exact target on restore, then rebuild its complete former contents.
            text.Append("[-HKEY_CURRENT_USER\\").Append(path).Append("]\r\n\r\n");
            if (key != null) Dump(text, path, key);
        }
        static void Dump(StringBuilder text, string path, RegistryKey key)
        {
            text.Append("[HKEY_CURRENT_USER\\").Append(path).Append("]\r\n");
            foreach (string name in key.GetValueNames()) {
                string prefix = name.Length == 0 ? "@=" : "\"" + Esc(name) + "\"=";
                object v = key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                switch (key.GetValueKind(name)) {
                    case RegistryValueKind.DWord:
                        text.Append(prefix).Append("dword:").Append(unchecked((uint)(int)v).ToString("x8")).Append("\r\n"); break;
                    case RegistryValueKind.String:
                        if (((string)v).Any(Char.IsControl)) Hex(text,prefix+"hex(1):",Encoding.Unicode.GetBytes((string)v+"\0"));
                        else text.Append(prefix).Append('"').Append(Esc((string)v)).Append("\"\r\n"); break;
                    case RegistryValueKind.ExpandString: Hex(text,prefix+"hex(2):",Encoding.Unicode.GetBytes((string)v+"\0")); break;
                    case RegistryValueKind.MultiString: Hex(text,prefix+"hex(7):",Encoding.Unicode.GetBytes(String.Join("\0",(string[])v)+"\0\0")); break;
                    case RegistryValueKind.QWord: Hex(text,prefix+"hex(b):",BitConverter.GetBytes((long)v)); break;
                    case RegistryValueKind.Binary: Hex(text,prefix+"hex:",(byte[])v); break;
                    case RegistryValueKind.None: Hex(text,prefix+"hex(0):",(byte[])v); break;
                    default: throw new Exception("Sicherung: Registry-Typ nicht unterstützt: " + name);
                }
            }
            text.Append("\r\n");
            foreach(string child in key.GetSubKeyNames()) using(RegistryKey k=key.OpenSubKey(child)) Dump(text,path+"\\"+child,k);
        }
        public static bool Exists(string key, string root)
        {
            using (RegistryKey r = Registry.CurrentUser.OpenSubKey(root+"\\"+key)) return r != null;
        }
        public static string Write(List<Session> sessions, bool overwrite, string backupDir, string root)
        {
            List<Session> targets = sessions.Where(s => overwrite || !Exists(s.Key,root)).ToList();
            if (targets.Count == 0) throw new Exception("Keine neuen Sessions. Bestehende Namen werden übersprungen.");
            StringBuilder backup = new StringBuilder("Windows Registry Editor Version 5.00\r\n\r\n");
            foreach (Session s in targets) using (RegistryKey old = Registry.CurrentUser.OpenSubKey(root+"\\"+s.Key))
                Snapshot(backup,root+"\\"+s.Key,old);
            Directory.CreateDirectory(backupDir);
            string backupFile = Path.Combine(backupDir,"PuTTY-vor-Import-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N").Substring(0,6)+".reg");
            File.WriteAllText(backupFile,backup.ToString(),Encoding.Unicode); // Nothing is written before a successful backup.
            List<string> done = new List<string>();
            try {
                using (RegistryKey parent = Registry.CurrentUser.CreateSubKey(root))
                using (RegistryKey defaults = parent.OpenSubKey("Default%20Settings")) {
                    foreach(Session s in targets) {
                        bool existing = Exists(s.Key,root);
                        using (RegistryKey key = parent.CreateSubKey(s.Key)) {
                            if (!existing && defaults != null) foreach(string name in defaults.GetValueNames())
                                key.SetValue(name,defaults.GetValue(name,null,RegistryValueOptions.DoNotExpandEnvironmentNames),defaults.GetValueKind(name));
                            foreach(KeyValuePair<string,object> v in s.Values)
                                key.SetValue(v.Key,v.Value,v.Value is int ? RegistryValueKind.DWord : RegistryValueKind.String);
                            foreach(KeyValuePair<string,object> v in s.Values)
                                if (!Object.Equals(key.GetValue(v.Key),v.Value)) throw new Exception("Rücklesen fehlgeschlagen: "+s.Name+" / "+v.Key);
                        }
                        done.Add(s.Name);
                    }
                }
            } catch(Exception ex) {
                throw new Exception("Import abgebrochen. Fertig geschriebene Sessions: "+done.Count+". Weitere Sessions können teilweise geschrieben sein.\r\nWiederherstellung per Doppelklick auf:\r\n"+backupFile+"\r\n\r\n"+ex.Message,ex);
            }
            return backupFile;
        }
        public static void Export(List<Session> sessions, string file)
        {
            StringBuilder text=new StringBuilder("Windows Registry Editor Version 5.00\r\n\r\n");
            foreach(Session s in sessions) {
                text.Append("[HKEY_CURRENT_USER\\").Append(Core.RegistryPath).Append('\\').Append(s.Key).Append("]\r\n");
                foreach(KeyValuePair<string,object> v in s.Values) {
                    text.Append('"').Append(Esc(v.Key)).Append("\"=");
                    if(v.Value is int) text.Append("dword:").Append(((int)v.Value).ToString("x8"));
                    else text.Append('"').Append(Esc((string)v.Value)).Append('"');
                    text.Append("\r\n");
                }
                text.Append("\r\n");
            }
            File.WriteAllText(file,text.ToString(),Encoding.Unicode);
        }
    }

    public sealed class MainForm : Form
    {
        readonly DataGridView grid = new DataGridView();
        readonly TextBox log = new TextBox();
        readonly Label info = new Label();
        readonly CheckBox overwrite = new CheckBox();
        readonly List<string> fonts;
        readonly string appDir=AppDomain.CurrentDomain.BaseDirectory;
        bool dirty;
        public MainForm()
        {
            Text="PuTTY Session Builder 1.0"; Font=new Font("Segoe UI",9);
            StartPosition=FormStartPosition.CenterScreen;
            Size=new Size(Math.Min(1320,Screen.PrimaryScreen.WorkingArea.Width-40),Math.Min(760,Screen.PrimaryScreen.WorkingArea.Height-40));
            MinimumSize=new Size(960,570);
            fonts=Core.InstalledFonts();
            TableLayoutPanel layout=new TableLayoutPanel { Dock=DockStyle.Fill,ColumnCount=1,RowCount=5,Padding=new Padding(12) };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,66));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,76));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,37));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute,110));
            Controls.Add(layout);
            Label title=new Label { Text="PuTTY Sessions aus einer Tabelle",Font=new Font("Segoe UI",18,FontStyle.Bold),Dock=DockStyle.Fill,AutoSize=false };
            title.Text+="\r\n"; // Separate short workflow description belongs below.
            info.Text="Zeile 1: Überschriften. Zeile 2: Beispiel, wird IMMER übersprungen. Eigene Sessions ab Zeile 3. Keepalive Rate in Sekunden.";
            info.Dock=DockStyle.Fill; info.TextAlign=ContentAlignment.MiddleLeft;
            TableLayoutPanel heading=new TableLayoutPanel { Dock=DockStyle.Fill,RowCount=2,ColumnCount=1 };
            heading.RowStyles.Add(new RowStyle(SizeType.Absolute,38)); heading.RowStyles.Add(new RowStyle(SizeType.Percent,100));
            heading.Controls.Add(title,0,0); heading.Controls.Add(info,0,1); layout.Controls.Add(heading,0,0);
            FlowLayoutPanel toolbar=new FlowLayoutPanel { Dock=DockStyle.Fill,WrapContents=true,Padding=new Padding(0,6,0,0) };
            AddButton(toolbar,"Datei laden…",LoadFile); AddButton(toolbar,"Neue Tabelle",NewTable);
            AddButton(toolbar,"CSV speichern…",SaveFile); AddButton(toolbar,"Vorlagen öffnen",OpenTemplates);
            AddButton(toolbar,"Prüfen",ValidateClick); AddButton(toolbar,"REG exportieren…",ExportClick);
            AddButton(toolbar,"Sessions importieren",ImportClick); AddButton(toolbar,"Anleitung",HelpClick);
            layout.Controls.Add(toolbar,0,1);
            grid.Dock=DockStyle.Fill; grid.AllowUserToAddRows=true; grid.AllowUserToDeleteRows=true;
            grid.AutoGenerateColumns=false; grid.RowHeadersWidth=60; grid.BackgroundColor=Color.White;
            grid.AlternatingRowsDefaultCellStyle.BackColor=Color.FromArgb(245,248,252);
            grid.EnableHeadersVisualStyles=false; grid.ColumnHeadersDefaultCellStyle.BackColor=Color.FromArgb(32,54,78);
            grid.ColumnHeadersDefaultCellStyle.ForeColor=Color.White; grid.ColumnHeadersHeight=42;
            grid.ColumnHeadersHeightSizeMode=DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ColumnHeadersDefaultCellStyle.WrapMode=DataGridViewTriState.True;
            for(int i=0;i<Core.Headers.Length;i++) {
                DataGridViewColumn col;
                if(i==3 || i==4 || i==6 || i==7 || i==10 || i==11 || i==12) {
                    DataGridViewComboBoxColumn cb=new DataGridViewComboBoxColumn { FlatStyle=FlatStyle.Flat };
                    IEnumerable<string> options=i==3 ? (IEnumerable<string>)fonts : i==4 ? Core.Styles : i==6 ? Core.Quality : i==11 ? new[]{"ssh","telnet","raw"} : new[]{"TRUE","FALSE"};
                    foreach(string option in options) cb.Items.Add(option);
                    col=cb;
                } else col=new DataGridViewTextBoxColumn();
                col.Name="C"+i; col.HeaderText=Core.Headers[i]; col.SortMode=DataGridViewColumnSortMode.NotSortable;
                col.Width=i==0 ? 215 : i==1 ? 185 : i==3 ? 155 : i==9 || i==14 ? 170 : 110;
                grid.Columns.Add(col);
            }
            grid.Columns[0].Frozen=true;
            grid.DefaultValuesNeeded+=(s,e)=> { string[] values=Defaults(); for(int i=0;i<values.Length;i++) e.Row.Cells[i].Value=values[i]; };
            grid.CellValueChanged+=(s,e)=> { if(e.RowIndex>=0) dirty=true; };
            grid.RowsRemoved+=(s,e)=>dirty=true;
            grid.UserDeletingRow+=(s,e)=> { if(e.Row.Index==0) { e.Cancel=true; MessageBox.Show("Die Beispielzeile muss erhalten bleiben."); } };
            grid.DataError+=(s,e)=> { e.ThrowException=false; grid.Rows[e.RowIndex].ErrorText="Ungültige Dropdown-Auswahl. Bitte einen angebotenen Wert wählen."; };
            grid.RowPostPaint+=(s,e)=> { TextRenderer.DrawText(e.Graphics,(e.RowIndex+2).ToString(),Font,new Rectangle(e.RowBounds.Left,e.RowBounds.Top,48,e.RowBounds.Height),Color.DimGray,TextFormatFlags.VerticalCenter|TextFormatFlags.Right); };
            layout.Controls.Add(grid,0,2);
            FlowLayoutPanel settings=new FlowLayoutPanel { Dock=DockStyle.Fill,Padding=new Padding(0,6,0,0) };
            overwrite.Text="Bestehende Sessions gleichen Namens aktualisieren"; overwrite.AutoSize=true;
            settings.Controls.Add(overwrite);
            settings.Controls.Add(new Label { AutoSize=true,Text=@"Ziel: HKCU\Software\SimonTatham\PuTTY\Sessions",Padding=new Padding(16,3,0,0) });
            layout.Controls.Add(settings,0,3);
            log.Dock=DockStyle.Fill; log.Multiline=true; log.ReadOnly=true; log.ScrollBars=ScrollBars.Vertical;
            log.BackColor=Color.FromArgb(245,248,252); layout.Controls.Add(log,0,4);
            FormClosing+=(s,e)=> { if(dirty && MessageBox.Show(this,"Änderungen wurden noch nicht als CSV gespeichert. Trotzdem schließen?","Schließen",MessageBoxButtons.YesNo,MessageBoxIcon.Question)!=DialogResult.Yes) e.Cancel=true; };
            Fill(new List<Entry>{new Entry(Core.Example.ToArray(),2)});
            SetLog("Bereit. Vorlagen öffnen, eine Kopie befüllen und Datei laden. Oder eigene Zeilen direkt hier eingeben.\r\n"+fonts.Count+" installierte Monospace-Schriften im Font-Dropdown. Kein Administrator erforderlich.");
        }
        void AddButton(FlowLayoutPanel panel,string text,EventHandler action)
        {
            Button b=new Button { Text=text,AutoSize=true,Height=30,Margin=new Padding(0,0,8,6),Padding=new Padding(6,0,6,0) };
            b.Click+=action; panel.Controls.Add(b);
        }
        string SelectedDefaultFont() { return fonts.Contains("Consolas") ? "Consolas" : fonts.FirstOrDefault() ?? "Consolas"; }
        string[] Defaults()
        {
            string[] d=Core.Example.ToArray(); d[0]=""; d[1]=""; d[3]=SelectedDefaultFont(); d[9]=""; return d;
        }
        void SetLog(string text) { log.Text=text; }
        void Error(Exception ex) { SetLog(ex.Message); MessageBox.Show(this,ex.Message,"PuTTY Session Builder",MessageBoxButtons.OK,MessageBoxIcon.Error); }
        bool Discard() { return !dirty || MessageBox.Show(this,"Ungespeicherte Änderungen verwerfen?","Tabelle wechseln",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes; }
        void Fill(List<Entry> rows)
        {
            grid.Rows.Clear();
            for(int ri=0;ri<rows.Count;ri++) {
                string[] c=rows[ri].Cells.ToArray();
                if(ri>0) {
                    try {
                        Session s=Core.Parse(new Entry(c,rows[ri].Line),null,Encoding.Default);
                        c[2]=s.Values["PortNumber"].ToString(); c[3]=(string)s.Values["Font"];
                        c[4]=(int)s.Values["FontIsBold"]==1 ? "Fett" : "Normal"; c[5]=s.Values["FontHeight"].ToString();
                        c[6]=Core.Quality[(int)s.Values["FontQuality"]]; c[7]=Core.Boolean(c[7],true,"Keepalive") ? "TRUE" : "FALSE";
                        c[8]=String.IsNullOrEmpty(c[8]) ? "5" : c[8]; c[10]=(int)s.Values["X11Forward"]==1 ? "TRUE" : "FALSE";
                        c[11]=(string)s.Values["Protocol"]; c[12]=(int)s.Values["TCPKeepalives"]==1 ? "TRUE" : "FALSE";
                        c[15]=s.Values["ScrollbackLines"].ToString();
                    } catch(Exception) { /* Preserve invalid input so the user can fix it. */ }
                }
                for(int i=0;i<c.Length;i++) {
                    DataGridViewComboBoxColumn cb=grid.Columns[i] as DataGridViewComboBoxColumn;
                    if(cb!=null && c[i].Length>0 && !cb.Items.Contains(c[i])) cb.Items.Add(c[i]);
                }
                int n=grid.Rows.Add(c.Cast<object>().ToArray());
                if(ri==0) { grid.Rows[n].ReadOnly=true; grid.Rows[n].DefaultCellStyle.BackColor=Color.FromArgb(255,242,204); grid.Rows[n].DefaultCellStyle.ForeColor=Color.FromArgb(112,89,27); }
            }
            dirty=false;
        }
        List<Entry> Rows()
        {
            grid.EndEdit(); List<Entry> rows=new List<Entry>();
            foreach(DataGridViewRow row in grid.Rows) {
                if(row.IsNewRow) continue;
                string[] c=row.Cells.Cast<DataGridViewCell>().Select(x=>Core.Text(x.Value)).ToArray();
                if(row.Index==0 || c.Any(x=>x.Length>0)) rows.Add(new Entry(c,row.Index+2));
            }
            return rows;
        }
        List<Session> ValidateRows()
        {
            List<Entry> rows=Rows(); Dictionary<int,string> errors=new Dictionary<int,string>();
            foreach(DataGridViewRow row in grid.Rows) row.ErrorText="";
            List<Session> sessions=Core.ValidateSessions(rows,fonts,Encoding.Default,errors);
            foreach(var error in errors) grid.Rows[error.Key-2].ErrorText=error.Value;
            if(errors.Count>0) throw new Exception("Bitte zuerst diese Fehler korrigieren:\r\n"+String.Join("\r\n",errors.Select(x=>"Zeile "+x.Key+": "+x.Value)));
            if(sessions.Count==0) throw new Exception("Keine Sessions: eigene Daten ab Zeile 3 eintragen. Die Beispielzeile wird nicht importiert.");
            return sessions;
        }
        void LoadFile(object sender,EventArgs e)
        {
            if(!Discard()) return;
            using(OpenFileDialog dlg=new OpenFileDialog { Filter="Sessiontabellen (*.csv;*.xls;*.xlsx)|*.csv;*.xls;*.xlsx",InitialDirectory=Path.Combine(appDir,"Vorlagen") })
                if(dlg.ShowDialog(this)==DialogResult.OK) try { List<Entry> rows=Core.Read(dlg.FileName); Fill(rows); SetLog(Path.GetFileName(dlg.FileName)+": "+(rows.Count-1)+" Datenzeilen geladen. Beispielzeile übersprungen. Änderungen hier gelten zunächst nur im Programm."); } catch(Exception ex) { Error(ex); }
        }
        void NewTable(object sender,EventArgs e) { if(Discard()) { Fill(new List<Entry>{new Entry(Core.Example.ToArray(),2)}); SetLog("Neue Tabelle. Eigene Daten in der leeren Zeile unter dem Beispiel eintragen."); } }
        void SaveFile(object sender,EventArgs e)
        {
            using(SaveFileDialog dlg=new SaveFileDialog { Filter="CSV (UTF-8, Semikolon)|*.csv",FileName="Meine-PuTTY-Sessions.csv" })
                if(dlg.ShowDialog(this)==DialogResult.OK) try { Core.SaveCsv(dlg.FileName,Rows()); dirty=false; SetLog("CSV gespeichert: "+dlg.FileName); } catch(Exception ex) { Error(ex); }
        }
        void OpenTemplates(object sender,EventArgs e) { try { System.Diagnostics.Process.Start(Path.Combine(appDir,"Vorlagen")); } catch(Exception ex) { Error(ex); } }
        void HelpClick(object sender,EventArgs e) { try { System.Diagnostics.Process.Start(Path.Combine(appDir,"Anleitung.html")); } catch(Exception ex) { Error(ex); } }
        void ValidateClick(object sender,EventArgs e)
        {
            try {
                List<Session> s=ValidateRows(); int existing=s.Count(x=>RegistryImport.Exists(x.Key,Core.RegistryPath));
                SetLog(s.Count+" gültige Sessions. "+existing+" bereits vorhanden. "+(overwrite.Checked ? "Diese werden aktualisiert." : "Diese werden beim Import übersprungen.")+"\r\n"+
                    String.Join("\r\n",s.Take(5).Select(x=>x.Name+" → "+x.Values["HostName"]+":"+x.Values["PortNumber"])));
            } catch(Exception ex) { Error(ex); }
        }
        string BackupDir() { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"PuTTYSessionBuilder","Backups"); }
        void ImportClick(object sender,EventArgs e)
        {
            try {
                List<Session> sessions=ValidateRows(); int existing=sessions.Count(x=>RegistryImport.Exists(x.Key,Core.RegistryPath));
                int count=overwrite.Checked ? sessions.Count : sessions.Count-existing;
                if(count==0) { SetLog("Alle Sessions sind bereits vorhanden und werden übersprungen. Zum Aktualisieren die Checkbox aktivieren."); return; }
                if(MessageBox.Show(this,count+" Sessions in deinem Windows-Benutzerprofil speichern?\r\n"+
                    (overwrite.Checked ? existing+" bestehende Sessions werden aktualisiert." : existing+" bestehende Sessions werden übersprungen.")+"\r\n\r\nVorher wird automatisch eine Wiederherstellungsdatei angelegt.","Registry-Import",MessageBoxButtons.YesNo,MessageBoxIcon.Question)!=DialogResult.Yes) return;
                string backup=RegistryImport.Write(sessions,overwrite.Checked,BackupDir(),Core.RegistryPath);
                SetLog(count+" Sessions importiert und durch Rücklesen geprüft.\r\nSicherung: "+backup+"\r\nPuTTY neu öffnen, um die Sessions anzuzeigen. Wiederherstellung: Sicherungsdatei per Doppelklick importieren.");
                MessageBox.Show(this,count+" Sessions gespeichert. PuTTY neu öffnen.\r\n\r\nSicherung:\r\n"+backup,"Import fertig",MessageBoxButtons.OK,MessageBoxIcon.Information);
            } catch(Exception ex) { Error(ex); }
        }
        void ExportClick(object sender,EventArgs e)
        {
            try {
                List<Session> sessions=ValidateRows();
                using(SaveFileDialog dlg=new SaveFileDialog { Filter="Registry-Datei (*.reg)|*.reg",FileName="PuTTY-Sessions.reg" })
                    if(dlg.ShowDialog(this)==DialogResult.OK) { RegistryImport.Export(sessions,dlg.FileName); SetLog("REG exportiert: "+dlg.FileName+"\r\nDer spätere Doppelklick auf diese Datei schreibt ohne Sicherung und überschreibt enthaltene Werte gleicher Sessionnamen. Für gesicherten Import den Programm-Button verwenden."); }
            } catch(Exception ex) { Error(ex); }
        }
    }
    public static class Program
    {
        [STAThread] public static void Main()
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            try { Application.Run(new MainForm()); }
            catch(Exception ex) { MessageBox.Show(ex.ToString(),"Startfehler",MessageBoxButtons.OK,MessageBoxIcon.Error); }
        }
    }
}
