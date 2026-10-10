using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using PuttySessionBuilder;
using ExcelDataReader;

public static class Tests
{
    static int checks;
    static void Check(bool ok,string label) { if(!ok) throw new Exception("FAIL: "+label); checks++; }
    static void Reject(Action action,string label) { bool failed=false; try { action(); } catch(Exception) { failed=true; } Check(failed,label); }
    static string[] Data() { var x=Core.Example.ToArray(); x[0]="Server A"; x[1]="server.example.net"; return x; }
    static Session Parse(string[] x) { return Core.Parse(new Entry(x,3),null,Encoding.GetEncoding(1252)); }
    public static void Main(string[] args)
    {
        string app=args[0], tmp=Path.Combine(Path.GetTempPath(),"PuttyBuilderTests-"+Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmp);
        string root=@"Software\PuTTYSessionBuilder.Tests\"+Guid.NewGuid().ToString("N");
        try {
            var x=Data(); var s=Parse(x);
            Check((int)s.Values["PortNumber"]==22,"SSH port");
            Check((int)s.Values["FontHeight"]==11,"font points");
            Check((int)s.Values["FontCharSet"]==0,"font charset");
            Check((int)s.Values["FontQuality"]==3,"ClearType");
            Check((int)s.Values["PingInterval"]==0 && (int)s.Values["PingIntervalSecs"]==5,"5 second keepalive");
            Check((string)s.Values["UserName"]=="admin","autologin username");
            Check((int)s.Values["X11Forward"]==0,"X11 false");
            Check(s.Key=="Server%20A","space escaping");
            Check(Core.KeyName(".Büro\\100%*?",Encoding.GetEncoding(1252))=="%2EB%FCro%5C100%25%2A%3F","PuTTY ANSI escaping");
            Check(Core.KeyName("a.b",Encoding.GetEncoding(1252))=="a.b","nonleading dot");
            Reject(()=>Core.KeyName("Server 🐧",Encoding.GetEncoding(1252)),"unrepresentable session name");
            x=Data(); x[8]="125"; s=Parse(x); Check((int)s.Values["PingInterval"]==2 && (int)s.Values["PingIntervalSecs"]==5,"125 second split");
            x[7]="FALSE"; x[8]="0"; s=Parse(x); Check((int)s.Values["PingInterval"]==0 && (int)s.Values["PingIntervalSecs"]==0,"disabled pings");
            x=Data(); for(int i=2;i<x.Length;i++) x[i]=""; s=Parse(x);
            Check((int)s.Values["PingIntervalSecs"]==5,"blank rate defaults to 5");
            Check((string)s.Values["Font"]=="Consolas" && (int)s.Values["FontHeight"]==11,"blank font defaults");
            Check((int)s.Values["ScrollbackLines"]==2000,"blank scrollback");
            Check((int)s.Values["X11Forward"]==0,"blank X11");
            Check(Core.Boolean("WAHR",false,"bool") && !Core.Boolean("nein",true,"bool"),"German booleans");
            Check(Core.Boolean("1",false,"bool") && !Core.Boolean("0",true,"bool"),"numeric booleans");
            for(int i=0;i<4;i++) { x=Data(); x[6]=Core.Quality[i]; Check((int)Parse(x).Values["FontQuality"]==i,"font quality "+i); }
            x=Data(); x[4]="bold"; Check((int)Parse(x).Values["FontIsBold"]==1,"bold style");
            x=Data(); x[11]="telnet"; x[2]=""; Check((int)Parse(x).Values["PortNumber"]==23,"telnet default");
            x=Data(); x[10]="ja"; x[12]="true"; s=Parse(x); Check((int)s.Values["X11Forward"]==1 && (int)s.Values["TCPKeepalives"]==1,"X11 and TCP true");
            x=Data(); x[1]="[2001:db8::1]"; Check((string)Parse(x).Values["HostName"]=="2001:db8::1","IPv6 brackets");
            x=Data(); x[0]="Default Settings"; Reject(()=>Parse(x),"reserved default settings");
            x=Data(); x[0]=""; Reject(()=>Parse(x),"missing name");
            x=Data(); x[1]=""; Reject(()=>Parse(x),"missing host");
            x=Data(); x[1]="ssh://host"; Reject(()=>Parse(x),"host URL");
            x=Data(); x[1]="host:22"; Reject(()=>Parse(x),"embedded port");
            x=Data(); x[2]="65536"; Reject(()=>Parse(x),"port overflow");
            x=Data(); x[2]="1"; Check((int)Parse(x).Values["PortNumber"]==1,"port minimum");
            x=Data(); x[2]="65535"; Check((int)Parse(x).Values["PortNumber"]==65535,"port maximum");
            x=Data(); x[4]="Italic"; Reject(()=>Parse(x),"unsupported italic");
            x=Data(); x[7]="maybe"; Reject(()=>Parse(x),"invalid boolean");
            x=Data(); x[8]="0"; Reject(()=>Parse(x),"enabled zero rate");
            x=Data(); x[11]="telnet"; x[10]="TRUE"; Reject(()=>Parse(x),"X11 requires SSH");
            x=Data(); Reject(()=>Core.Parse(new Entry(x,3),new[]{"Courier New"},Encoding.GetEncoding(1252)),"uninstalled font");
            x=Data(); x[0]=new string('ü',100); Reject(()=>Parse(x),"encoded registry name too long");
            var rows=Core.Read(Path.Combine(app,"Vorlagen","PuTTY-Sessions-Vorlage.csv"));
            Check(rows.Count==1 && rows[0].Line==2,"CSV example only");
            rows=Core.Read(Path.Combine(app,"Vorlagen","PuTTY-Sessions-Vorlage.xlsx"));
            Check(rows.Count==1 && rows[0].Line==2,"XLSX example only and Sessions sheet");
            x=Data(); x[0]="Büro; \"Nord\""; string csv=Path.Combine(tmp,"sessions.csv");
            Core.SaveCsv(csv,new[]{new Entry(Core.Example.ToArray(),2),new Entry(x,3)});
            rows=Core.Read(csv); Check(rows.Count==2 && rows[1].Cells[0]==x[0],"CSV quoted roundtrip");
            Check(rows.Skip(1).Count()==1,"example excluded");
            var issues=new Dictionary<int,string>();
            var validated=Core.ValidateSessions(rows,null,Encoding.GetEncoding(1252),issues);
            Check(validated.Count==1 && issues.Count==0,"batch skips first example");
            var copy=x.ToArray(); copy[0]=x[0].ToLowerInvariant(); rows.Add(new Entry(copy,4));
            issues.Clear(); validated=Core.ValidateSessions(rows,null,Encoding.GetEncoding(1252),issues);
            Check(issues.ContainsKey(4) && validated.Count==1,"case-insensitive duplicate detection");
            byte[] utf8=File.ReadAllBytes(csv); Check(utf8[0]==239 && utf8[1]==187 && utf8[2]==191,"UTF8 BOM");
            string ansi=Path.Combine(tmp,"ansi.csv");
            File.WriteAllText(ansi,File.ReadAllText(csv),Encoding.GetEncoding(1252));
            Check(Core.Read(ansi)[1].Cells[0]==x[0],"legacy ANSI CSV");
            string comma=Path.Combine(tmp,"comma.csv");
            File.WriteAllText(comma,String.Join(",",Core.Headers.Select(Core.CsvField))+"\r\n"+String.Join(",",Core.Example.Select(Core.CsvField))+"\r\n"+String.Join(",",x.Select(Core.CsvField)),new UTF8Encoding(true));
            Check(Core.Read(comma)[1].Cells[0]==x[0],"comma CSV detection");
            string bad=Path.Combine(tmp,"bad.csv"); File.WriteAllText(bad,"Anzeigename;Hostname/IP\nExample;host\n");
            Reject(()=>Core.Read(bad),"missing header");
            // Isolated registry branch only, never real PuTTY sessions.
            using(RegistryKey defaults=Registry.CurrentUser.CreateSubKey(root+"\\Default%20Settings")) {
                defaults.SetValue("InheritedSetting","yes"); defaults.SetValue("HostName","default-host");
            }
            x=Data(); s=Parse(x);
            string backup=RegistryImport.Write(new List<Session>{s},false,tmp,root);
            using(RegistryKey k=Registry.CurrentUser.OpenSubKey(root+"\\"+s.Key)) {
                Check((string)k.GetValue("InheritedSetting")=="yes","inherits defaults");
                Check((string)k.GetValue("HostName")==x[1],"explicit host overrides default");
                Check(k.GetValueKind("PortNumber")==RegistryValueKind.DWord,"registry DWORD");
                Check(k.GetValueKind("Font")==RegistryValueKind.String,"registry string");
            }
            Check(File.ReadAllText(backup).Contains("[-HKEY_CURRENT_USER\\"+root+"\\"+s.Key+"]"),"backup deletes newly created session on restore");
            x[1]="changed.example.net"; Session updated=Parse(x);
            Reject(()=>RegistryImport.Write(new List<Session>{updated},false,tmp,root),"existing sessions skipped");
            using(RegistryKey k=Registry.CurrentUser.OpenSubKey(root+"\\"+s.Key)) Check((string)k.GetValue("HostName")=="server.example.net","skip preserves host");
            backup=RegistryImport.Write(new List<Session>{updated},true,tmp,root);
            Check(File.ReadAllText(backup).Contains("server.example.net"),"backup preserves original value");
            using(RegistryKey k=Registry.CurrentUser.OpenSubKey(root+"\\"+s.Key)) {
                Check((string)k.GetValue("HostName")=="changed.example.net","overwrite host");
                Check((string)k.GetValue("InheritedSetting")=="yes","overwrite preserves other values");
            }
            string export=Path.Combine(tmp,"export.reg"); RegistryImport.Export(new List<Session>{s},export);
            Check(File.ReadAllText(export).Contains("\"PortNumber\"=dword:00000016"),"REG DWORD encoding");
            Check(File.ReadAllText(export).Contains(Core.RegistryPath+"\\Server%20A"),"REG path");
            // Optional upstream BIFF .xls fixture.
            if(args.Length>1) using(var stream=File.OpenRead(args[1])) using(var reader=ExcelReaderFactory.CreateReader(stream))
                Check(reader.Read() && reader.FieldCount>0,"legacy binary XLS reader");
            Console.WriteLine("PASS: "+checks+" checks. CSV/XLSX, validation, PuTTY mapping, isolated registry writes, backup and export.");
        } finally { Registry.CurrentUser.DeleteSubKeyTree(root,false); Directory.Delete(tmp,true); }
    }
}
