// Lumen Office desktop host: a real Windows program (own .exe, icon and taskbar/dock entry) showing a Lumen app in
// WebView2. Self-contained: it serves the app files itself on the private address https://lumen.local/ (nothing listens
// on the network), reads and writes the documents you open, and installs updates from GitHub Releases. No Python needed.
// One source builds Lumen Write.exe, Lumen Cells.exe, Lumen Slides.exe and Lumen Office.exe; the file name picks the app.
//
//   "Lumen Write.exe"                  new document
//   "Lumen Write.exe" "C:\a\b.docx"    open a file (what Explorer runs); Ctrl+S saves back to it
//   "Lumen Office.exe" --updates       start page, checking for updates
//
// Written for the C# 5 compiler that ships with Windows (.NET Framework 4.x).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

// each program carries its own name (shown by Windows in Default apps, Open with and Task Manager)
#if APP_WRITE
[assembly: System.Reflection.AssemblyTitle("Lumen Write")]
#elif APP_CELLS
[assembly: System.Reflection.AssemblyTitle("Lumen Cells")]
#elif APP_SLIDES
[assembly: System.Reflection.AssemblyTitle("Lumen Slides")]
#else
[assembly: System.Reflection.AssemblyTitle("Lumen Office")]
#endif
[assembly: System.Reflection.AssemblyProduct("Lumen Office")]
[assembly: System.Reflection.AssemblyCompany("Lumen Office")]
[assembly: System.Reflection.AssemblyVersion("1.3.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.3.0.0")]

namespace Lumen
{
    static class Program
    {
        public const string Host = "lumen.local";
        public const string Base = "https://lumen.local/";
        public static string Root = Path.GetDirectoryName(Application.ExecutablePath);
        public static string Data = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Lumen Office");
        public static string App = "office";
        public static CoreWebView2Environment Env;
        public static readonly Dictionary<string, string> Opened = new Dictionary<string, string>();
        static readonly HashSet<string> OpenTypes = new HashSet<string> { ".docx", ".docm", ".dotx", ".txt", ".xlsx", ".xlsm", ".csv", ".pptx" };

        [DllImport("shell32.dll", SetLastError = true)]
        static extern void SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string id);

        public static string PageOf(string app)
        {
            if (app == "write") return "write/index.html";
            if (app == "cells") return "cells/index.html";
            if (app == "slides") return "slides/index.html";
            return "index.html";
        }

        public static string ExeOf(string app)
        {
            string n = app == "write" ? "Lumen Write" : app == "cells" ? "Lumen Cells" : app == "slides" ? "Lumen Slides" : "Lumen Office";
            return Path.Combine(Root, n + ".exe");
        }

        public static string Register(string file)
        {
            try
            {
                string p = Path.GetFullPath(file);
                if (!File.Exists(p) || !OpenTypes.Contains(Path.GetExtension(p).ToLowerInvariant())) return null;
                lock (Opened)
                {
                    foreach (var kv in Opened) if (string.Equals(kv.Value, p, StringComparison.OrdinalIgnoreCase)) return kv.Key;
                    byte[] b = new byte[18];
                    using (var rng = RandomNumberGenerator.Create()) rng.GetBytes(b);
                    string t = Convert.ToBase64String(b).Replace('+', '-').Replace('/', '_').TrimEnd('=');
                    Opened[t] = p;
                    return t;
                }
            }
            catch { return null; }
        }

        [STAThread]
        static void Main(string[] args)
        {
            string me = Path.GetFileNameWithoutExtension(Application.ExecutablePath).ToLowerInvariant();
            if (me.Contains("write")) App = "write"; else if (me.Contains("cells")) App = "cells"; else if (me.Contains("slides")) App = "slides";
            try { SetCurrentProcessExplicitAppUserModelID("LumenOffice." + char.ToUpper(App[0]) + App.Substring(1)); } catch { }
            try { ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072; } catch { }  // TLS 1.2 for GitHub
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Updater.CleanupOld();

            string file = null; bool updates = false;
            foreach (string a in args) { if (a == "--updates") updates = true; else if (!a.StartsWith("--")) file = a; }
            string token = null;
            if (file != null)
            {
                token = Register(file);
                if (token == null) MessageBox.Show("Lumen Office cannot open this file:\n" + file, "Lumen Office", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (App == "office") App = ext == ".xlsx" || ext == ".xlsm" || ext == ".csv" ? "cells" : ext == ".pptx" ? "slides" : "write";
            }
            string url = Base + PageOf(App) + (token != null ? "?open=" + token : "") + (updates && App == "office" ? "#updates" : "");
            Application.Run(new HostForm(url, null, App == "office"));
        }

        /* ---------------- the private web address: app files + document and update services ---------------- */
        static readonly Dictionary<string, string> Mime = new Dictionary<string, string> {
            { ".html", "text/html; charset=utf-8" }, { ".js", "text/javascript; charset=utf-8" }, { ".css", "text/css; charset=utf-8" },
            { ".json", "application/json" }, { ".png", "image/png" }, { ".ico", "image/x-icon" }, { ".svg", "image/svg+xml" },
            { ".webmanifest", "application/manifest+json" }, { ".txt", "text/plain; charset=utf-8" }, { ".woff2", "font/woff2" } };

        static CoreWebView2WebResourceResponse Respond(int code, string reason, byte[] body, string headers)
        {
            return Env.CreateWebResourceResponse(new MemoryStream(body ?? new byte[0]), code, reason, (headers ?? "") + "Cache-Control: no-cache\r\nX-Content-Type-Options: nosniff");
        }

        public static CoreWebView2WebResourceResponse Json(int code, object o)
        {
            return Respond(code, code == 200 ? "OK" : "Error", Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(o)), "Content-Type: application/json\r\n");
        }

        public static void Serve(CoreWebView2WebResourceRequestedEventArgs e)
        {
            Uri u = new Uri(e.Request.Uri);
            string path = Uri.UnescapeDataString(u.AbsolutePath);
            string method = e.Request.Method.ToUpperInvariant();
            try
            {
                if (path == "/_ping") { e.Response = Json(200, new Dictionary<string, object> { { "app", "lumen-office" }, { "version", Updater.CurrentVersion() } }); return; }
                if (path.StartsWith("/_file/"))
                {
                    string tok = path.Substring(7), p;
                    lock (Opened) { if (!Opened.TryGetValue(tok, out p)) p = null; }
                    if (p == null) { e.Response = Json(404, new Dictionary<string, object> { { "error", "unknown file" } }); return; }
                    if (method == "GET")
                    {
                        e.Response = Respond(200, "OK", File.ReadAllBytes(p), "Content-Type: application/octet-stream\r\nX-Lumen-Name: " + Uri.EscapeDataString(Path.GetFileName(p)) + "\r\n");
                        return;
                    }
                    if (method == "PUT")
                    {
                        if (e.Request.Headers.GetHeader("X-Lumen") != "1") { e.Response = Json(403, new Dictionary<string, object> { { "error", "forbidden" } }); return; }
                        var ms = new MemoryStream();
                        if (e.Request.Content != null) e.Request.Content.CopyTo(ms);
                        string tmp = p + ".lumen-tmp";
                        try
                        {
                            File.WriteAllBytes(tmp, ms.ToArray());
                            if (File.Exists(p)) File.Replace(tmp, p, null); else File.Move(tmp, p);
                        }
                        catch (IOException ex) { try { File.Delete(tmp); } catch { } e.Response = Json(423, new Dictionary<string, object> { { "error", "The file is open in another program or read-only. " + ex.Message } }); return; }
                        catch (UnauthorizedAccessException) { try { File.Delete(tmp); } catch { } e.Response = Json(423, new Dictionary<string, object> { { "error", "No permission to write this file (read-only?)." } }); return; }
                        e.Response = Json(200, new Dictionary<string, object> { { "saved", Path.GetFileName(p) } });
                        return;
                    }
                }
                if (path == "/_update/status" || path == "/_update/apply")
                {
                    var deferral = e.GetDeferral();
                    bool apply = path == "/_update/apply" && method == "POST" && e.Request.Headers.GetHeader("X-Lumen") == "1";
                    bool force = u.Query.Contains("refresh=1");
                    var ui = System.Threading.SynchronizationContext.Current;
                    Task.Run(() =>
                    {
                        Dictionary<string, object> res; int code = 200;
                        try
                        {
                            if (apply) { string v = Updater.Apply(); res = new Dictionary<string, object> { { "version", v }, { "restarting", true } }; }
                            else res = Updater.Check(force);
                        }
                        catch (Exception ex) { code = 409; res = new Dictionary<string, object> { { "error", ex.Message } }; }
                        ui.Post(_ => { e.Response = Json(code, res); deferral.Complete(); }, null);
                    });
                    return;
                }
                if (path.StartsWith("/_") || method != "GET") { e.Response = Json(404, new Dictionary<string, object> { { "error", "not found" } }); return; }
                // static files of the apps (no service worker inside the program; never anything outside the install folder)
                if (path == "/") path = "/index.html";
                if (path == "/sw.js" || path.StartsWith("/tools/") || path.StartsWith("/host/")) { e.Response = Respond(404, "Not Found", null, ""); return; }
                string full = Path.GetFullPath(Path.Combine(Root, path.TrimStart('/').Replace('/', '\\')));
                if (!full.StartsWith(Root + "\\", StringComparison.OrdinalIgnoreCase) || !File.Exists(full)) { e.Response = Respond(404, "Not Found", Encoding.UTF8.GetBytes("Not found"), "Content-Type: text/plain\r\n"); return; }
                string mime;
                if (!Mime.TryGetValue(Path.GetExtension(full).ToLowerInvariant(), out mime)) mime = "application/octet-stream";
                e.Response = Respond(200, "OK", File.ReadAllBytes(full), "Content-Type: " + mime + "\r\n");
            }
            catch (Exception ex) { e.Response = Json(500, new Dictionary<string, object> { { "error", ex.Message } }); }
        }
    }

    /* ---------------- updates from GitHub Releases (same rules as tools/updater.py) ---------------- */
    static class Updater
    {
        static string CacheFile { get { return Path.Combine(Program.Data, "update-status.json"); } }
        static readonly HashSet<string> Keep = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "update.json", "python.path" };

        static Dictionary<string, object> ReadJson(string path)
        {
            try { return new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(path)); }  // ReadAllText strips a BOM
            catch { return new Dictionary<string, object>(); }
        }
        static string Str(Dictionary<string, object> d, string k) { object o; return d != null && d.TryGetValue(k, out o) && o != null ? o.ToString() : ""; }

        public static string CurrentVersion() { string v = Str(ReadJson(Path.Combine(Program.Root, "version.json")), "version"); return v == "" ? "0.0.0" : v; }

        static int[] V(string s)
        {
            var m = Regex.Matches(s ?? "", "\\d+"); int[] r = new int[3];
            for (int i = 0; i < 3 && i < m.Count; i++) int.TryParse(m[i].Value, out r[i]);
            return r;
        }
        static int Cmp(string a, string b) { int[] x = V(a), y = V(b); for (int i = 0; i < 3; i++) if (x[i] != y[i]) return x[i].CompareTo(y[i]); return 0; }

        static byte[] Get(string url, string accept)
        {
            var rq = (HttpWebRequest)WebRequest.Create(url);
            rq.UserAgent = "LumenOffice-Updater"; rq.Accept = accept; rq.Timeout = 30000;
            using (var rs = rq.GetResponse()) using (var s = rs.GetResponseStream()) using (var ms = new MemoryStream()) { s.CopyTo(ms); return ms.ToArray(); }
        }

        public static Dictionary<string, object> Check(bool force)
        {
            var cfg = ReadJson(Path.Combine(Program.Root, "update.json"));
            string repo = Str(cfg, "repo").Trim().Trim('/'), cur = CurrentVersion();
            var st = new Dictionary<string, object> { { "current", cur }, { "repo", repo }, { "available", false }, { "checked", DateTimeOffset.UtcNow.ToUnixTimeSeconds() } };
            if (repo == "" || !Regex.IsMatch(repo, "^[\\w.-]+/[\\w.-]+$")) { st["error"] = "No GitHub repository is set. Put it in update.json as \"owner/name\"."; return st; }
            double hours = 24; double.TryParse(Str(cfg, "check_every_hours"), out hours); if (hours <= 0) hours = 24;
            var cached = File.Exists(CacheFile) ? ReadJson(CacheFile) : null;
            long when; long.TryParse(Str(cached, "checked"), out when);
            if (!force && cached != null && Str(cached, "repo") == repo && Str(cached, "current") == cur && DateTimeOffset.UtcNow.ToUnixTimeSeconds() - when < hours * 3600) return cached;
            try
            {
                var rel = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(Encoding.UTF8.GetString(Get("https://api.github.com/repos/" + repo + "/releases/latest", "application/vnd.github+json")));
                string tag = Str(rel, "tag_name");
                string zip = "", sha = "";
                object assetsObj; rel.TryGetValue("assets", out assetsObj);
                var assets = assetsObj as System.Collections.ArrayList;
                if (assets != null)
                {
                    foreach (object ao in assets) { var a = ao as Dictionary<string, object>; string n = Str(a, "name"); if (Regex.IsMatch(n, "^LumenOffice-[\\w.-]+\\.zip$")) zip = Str(a, "browser_download_url"); }
                    foreach (object ao in assets) { var a = ao as Dictionary<string, object>; string n = Str(a, "name"); if (zip != "" && n == Path.GetFileName(new Uri(zip).AbsolutePath) + ".sha256") sha = Str(a, "browser_download_url"); }
                }
                st["latest"] = tag.TrimStart('v', 'V'); string notes = Str(rel, "body"); st["notes"] = notes.Length > 2000 ? notes.Substring(0, 2000) : notes;
                st["page"] = Str(rel, "html_url"); st["zip"] = zip; st["sha256"] = sha;
                st["available"] = Cmp(tag, cur) > 0 && zip != "";
                if (Cmp(tag, cur) > 0 && zip == "") st["error"] = "Release " + tag + " has no LumenOffice-*.zip file attached.";
            }
            catch (Exception ex) { st["error"] = "Could not check GitHub: " + ex.Message; }
            try { Directory.CreateDirectory(Program.Data); File.WriteAllText(CacheFile, new JavaScriptSerializer().Serialize(st)); } catch { }
            return st;
        }

        public static string Apply()
        {
            var st = Check(true);
            if (!(st["available"] is bool) || !(bool)st["available"]) throw new Exception(st.ContainsKey("error") ? Str(st, "error") : "Lumen Office is up to date.");
            if (Str(st, "sha256") == "") throw new Exception("The release has no .sha256 checksum file, so the download cannot be verified. Update refused.");
            byte[] blob = Get(Str(st, "zip"), "application/octet-stream");
            string want = Encoding.ASCII.GetString(Get(Str(st, "sha256"), "application/octet-stream")).Trim().Split(' ', '\t')[0].ToLowerInvariant();
            string got; using (var sha = SHA256.Create()) got = BitConverter.ToString(sha.ComputeHash(blob)).Replace("-", "").ToLowerInvariant();
            if (got != want) throw new Exception("The download does not match its checksum. Update refused.");
            string backup = Path.Combine(Program.Data, "previous-version");
            try { if (Directory.Exists(backup)) Directory.Delete(backup, true); } catch { }
            using (var za = new ZipArchive(new MemoryStream(blob), ZipArchiveMode.Read))
            {
                string prefix = null;
                foreach (var en in za.Entries)
                {
                    string n = en.FullName.Replace('\\', '/');
                    if (n.Contains("..") || n.StartsWith("/") || n.Contains(":")) throw new Exception("The update contains an unsafe file path. Update refused.");
                    if (n.EndsWith("version.json") && (prefix == null || n.Length < prefix.Length + 12)) prefix = n.Substring(0, n.Length - "version.json".Length);
                }
                if (prefix == null) throw new Exception("The download is not a Lumen Office release. Update refused.");
                foreach (var en in za.Entries)
                {
                    string n = en.FullName.Replace('\\', '/');
                    if (!n.StartsWith(prefix) || n.EndsWith("/") || en.Name == "") continue;
                    string rel = n.Substring(prefix.Length).Replace('/', '\\');
                    if (Keep.Contains(rel) && File.Exists(Path.Combine(Program.Root, rel))) continue;
                    string dst = Path.Combine(Program.Root, rel);
                    Directory.CreateDirectory(Path.GetDirectoryName(dst));
                    if (File.Exists(dst)) { string b = Path.Combine(backup, rel); Directory.CreateDirectory(Path.GetDirectoryName(b)); File.Copy(dst, b, true); }
                    string tmp = dst + ".new";
                    using (var src = en.Open()) using (var fs = File.Create(tmp)) src.CopyTo(fs);
                    try { if (File.Exists(dst)) File.Delete(dst); File.Move(tmp, dst); }
                    catch (Exception)
                    {
                        // a running program or loaded DLL cannot be overwritten but can be renamed; it is deleted on the next start
                        File.Move(dst, dst + ".lumen-old-" + DateTime.Now.Ticks);
                        File.Move(tmp, dst);
                    }
                }
            }
            try { File.Delete(CacheFile); } catch { }
            return CurrentVersion();
        }

        public static void CleanupOld()
        {
            try { foreach (string f in Directory.GetFiles(Program.Root, "*.lumen-old-*", SearchOption.AllDirectories)) { try { File.Delete(f); } catch { } } } catch { }
        }
    }

    /* ---------------- Windows 11 window chrome ---------------- */
    static class Win11
    {
        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
        const int DARK = 20, CORNERS = 33, BORDER = 34, CAPTION = 35, TEXT = 36;

        public static bool AppsDark()
        {
            try
            {
                using (var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize"))
                { object v = k == null ? null : k.GetValue("AppsUseLightTheme"); return v is int && (int)v == 0; }
            }
            catch { return false; }
        }
        static int Rgb(Color c) { return c.R | (c.G << 8) | (c.B << 16); }
        // title bar in the same colour as the Lumen toolbar, so window and app read as one surface
        public static Color Chrome(bool dark) { return dark ? Color.FromArgb(36, 36, 38) : Color.FromArgb(250, 250, 252); }
        public static Color Canvas(bool dark) { return dark ? Color.FromArgb(27, 27, 29) : Color.FromArgb(242, 242, 245); }
        public static void Apply(Form f)
        {
            bool dark = AppsDark();
            try
            {
                int on = dark ? 1 : 0; DwmSetWindowAttribute(f.Handle, DARK, ref on, 4);
                int cap = Rgb(Chrome(dark)); DwmSetWindowAttribute(f.Handle, CAPTION, ref cap, 4);
                int txt = Rgb(dark ? Color.FromArgb(245, 245, 247) : Color.FromArgb(29, 29, 31)); DwmSetWindowAttribute(f.Handle, TEXT, ref txt, 4);
                int border = Rgb(dark ? Color.FromArgb(58, 58, 62) : Color.FromArgb(218, 218, 222)); DwmSetWindowAttribute(f.Handle, BORDER, ref border, 4);
                int round = 2; DwmSetWindowAttribute(f.Handle, CORNERS, ref round, 4);
            }
            catch { }  // older Windows: keep the default title bar
            f.BackColor = Canvas(dark);
        }
    }

    /* ---------------- remembers each app window's size and place ---------------- */
    static class Placement
    {
        static string FileOf(string app) { return Path.Combine(Program.Data, "window-" + app + ".txt"); }
        public static void Restore(Form f, string app)
        {
            try
            {
                string[] p = File.ReadAllText(FileOf(app)).Split(',');
                var r = new Rectangle(int.Parse(p[0]), int.Parse(p[1]), int.Parse(p[2]), int.Parse(p[3]));
                foreach (var s in Screen.AllScreens)  // only if that place is still on a connected screen
                    if (s.WorkingArea.IntersectsWith(r) && r.Width >= 400 && r.Height >= 300)
                    {
                        f.StartPosition = FormStartPosition.Manual; f.Bounds = r;
                        if (p.Length > 4 && p[4] == "max") f.WindowState = FormWindowState.Maximized;
                        return;
                    }
            }
            catch { }
        }
        public static void Save(Form f, string app)
        {
            try
            {
                var r = f.WindowState == FormWindowState.Normal ? f.Bounds : f.RestoreBounds;
                Directory.CreateDirectory(Program.Data);
                File.WriteAllText(FileOf(app), r.X + "," + r.Y + "," + r.Width + "," + r.Height + "," + (f.WindowState == FormWindowState.Maximized ? "max" : "normal"));
            }
            catch { }
        }
    }

    class HostForm : Form
    {
        readonly WebView2 web = new WebView2();
        readonly string startUrl;
        readonly bool isStartPage;
        CoreWebView2NewWindowRequestedEventArgs pendingPopup;
        CoreWebView2Deferral pendingDeferral;
        bool allowClose;
        FormWindowState savedState; FormBorderStyle savedBorder;

        public HostForm(string url, CoreWebView2NewWindowRequestedEventArgs popup, bool startPage)
        {
            startUrl = url; isStartPage = startPage;
            Text = startPage ? "Lumen Office" : "Lumen";
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            float k = 1f; try { using (var g = CreateGraphics()) k = Math.Max(1f, g.DpiX / 96f); } catch { }  // follow the Windows display scale
            Size = startPage ? new Size((int)(460 * k), (int)(250 * k)) : new Size((int)Math.Min(1360 * k, Screen.PrimaryScreen.WorkingArea.Width * .9), (int)Math.Min(860 * k, Screen.PrimaryScreen.WorkingArea.Height * .9));
            if (startPage) { MaximizeBox = false; MinimizeBox = false; FormBorderStyle = FormBorderStyle.FixedSingle; StartPosition = FormStartPosition.CenterScreen; }
            StartPosition = FormStartPosition.WindowsDefaultLocation;
            bool dark = Win11.AppsDark();
            BackColor = Win11.Canvas(dark);
            web.Dock = DockStyle.Fill;
            web.DefaultBackgroundColor = Win11.Canvas(dark);  // no white flash while a dark-mode page loads
            if (popup == null && !startPage) Placement.Restore(this, Program.App);
            HandleCreated += (s, e) => Win11.Apply(this);
            Microsoft.Win32.SystemEvents.UserPreferenceChanged += OnThemeChanged;
            FormClosed += (s, e) => Microsoft.Win32.SystemEvents.UserPreferenceChanged -= OnThemeChanged;
            Controls.Add(web);
            if (popup != null)
            {
                pendingPopup = popup; pendingDeferral = popup.GetDeferral();
                var f = popup.WindowFeatures;
                if (f.HasSize) Size = new Size((int)Math.Max(320, f.Width), (int)Math.Max(240, f.Height));
                if (f.HasPosition) { StartPosition = FormStartPosition.Manual; Location = new Point((int)f.Left, (int)f.Top); }
            }
            Load += async (s, e) => await Init();
            FormClosing += OnClosing;
        }

        async Task Init()
        {
            try
            {
                if (Program.Env == null)
                    Program.Env = await CoreWebView2Environment.CreateAsync(null, Path.Combine(Program.Data, "WebView2"), null);
                await web.EnsureCoreWebView2Async(Program.Env);
            }
            catch (WebView2RuntimeNotFoundException)
            {
                MessageBox.Show("The Microsoft Edge WebView2 Runtime is missing. Install it from https://go.microsoft.com/fwlink/p/?LinkId=2124703 and start Lumen again.", "Lumen Office", MessageBoxButtons.OK, MessageBoxIcon.Error);
                allowClose = true; Close(); return;
            }
            var wv = web.CoreWebView2;
            wv.Settings.AreDevToolsEnabled = false;
            wv.Settings.IsStatusBarEnabled = false;
            wv.AddWebResourceRequestedFilter(Program.Base + "*", CoreWebView2WebResourceContext.All);
            wv.WebResourceRequested += (s, e) => Program.Serve(e);
            wv.DocumentTitleChanged += (s, e) => { Text = string.IsNullOrEmpty(wv.DocumentTitle) ? "Lumen" : wv.DocumentTitle; };
            wv.NewWindowRequested += OnNewWindow;
            wv.WindowCloseRequested += (s, e) => { allowClose = true; Close(); };
            wv.ContainsFullScreenElementChanged += (s, e) => SetFullScreen(wv.ContainsFullScreenElement);
            wv.NavigationStarting += OnNavigationStarting;
            if (pendingPopup != null)
            {
                pendingPopup.NewWindow = wv;
                pendingPopup.Handled = true;
                pendingDeferral.Complete();
                pendingPopup = null; pendingDeferral = null;
            }
            else web.Source = new Uri(startUrl);
            // automated check used when building releases: LUMEN_SELFTEST=save adds a line, presses Ctrl+S, closes
            if (Environment.GetEnvironmentVariable("LUMEN_SELFTEST") == "save")
            {
                await Task.Delay(4000);
                await wv.ExecuteScriptAsync("(function(){var d=document.getElementById('doc');if(d){var p=document.createElement('p');p.textContent='Saved by the Lumen self-test.';d.appendChild(p);}document.dispatchEvent(new KeyboardEvent('keydown',{key:'s',ctrlKey:true,bubbles:true}));})()");
                await Task.Delay(3000);
                allowClose = true; Close();
            }
        }

        // only Lumen's own pages load in the window; web links open in the normal browser.
        // From the start page, the app cards open their own program windows.
        void OnNavigationStarting(object sender, CoreWebView2NavigationStartingEventArgs e)
        {
            Uri u;
            if (!Uri.TryCreate(e.Uri, UriKind.Absolute, out u)) return;
            if (u.Scheme == "about" || u.Scheme == "data" || u.Scheme == "blob") return;
            if (u.Host != Program.Host)
            {
                e.Cancel = true;
                if (u.Scheme == "http" || u.Scheme == "https" || u.Scheme == "mailto") { try { Process.Start(new ProcessStartInfo(u.AbsoluteUri) { UseShellExecute = true }); } catch { } }
                return;
            }
            string cur = web.CoreWebView2 != null ? web.CoreWebView2.Source ?? "" : "";
            bool onStart = cur.StartsWith(Program.Base) && !cur.Contains("/write/") && !cur.Contains("/cells/") && !cur.Contains("/slides/");
            if (isStartPage && onStart)
            {
                string app = u.AbsolutePath.StartsWith("/write/") ? "write" : u.AbsolutePath.StartsWith("/cells/") ? "cells" : u.AbsolutePath.StartsWith("/slides/") ? "slides" : null;
                if (app != null && File.Exists(Program.ExeOf(app)))
                {
                    e.Cancel = true;
                    try { Process.Start(new ProcessStartInfo(Program.ExeOf(app)) { UseShellExecute = false, WorkingDirectory = Program.Root }); } catch { }
                    // the chooser is only a chooser: it goes away once an app is picked
                    allowClose = true; BeginInvoke(new Action(Close));
                }
            }
        }

        void OnNewWindow(object sender, CoreWebView2NewWindowRequestedEventArgs e)
        {
            Uri u;
            if (!string.IsNullOrEmpty(e.Uri) && e.Uri != "about:blank" && Uri.TryCreate(e.Uri, UriKind.Absolute, out u) && u.Host != Program.Host)
            {
                e.Handled = true;
                try { Process.Start(new ProcessStartInfo(u.AbsoluteUri) { UseShellExecute = true }); } catch { }
                return;
            }
            // the app's own popups (e.g. the slide show audience window) become program windows too
            new HostForm(null, e, false).Show();
        }

        // follow Windows when you switch between light and dark mode
        void OnThemeChanged(object sender, Microsoft.Win32.UserPreferenceChangedEventArgs e)
        {
            if (IsDisposed || !IsHandleCreated) return;
            try { BeginInvoke(new Action(() => { Win11.Apply(this); web.DefaultBackgroundColor = Win11.Canvas(Win11.AppsDark()); Invalidate(); })); } catch { }
        }

        void SetFullScreen(bool on)
        {
            if (on)
            {
                savedState = WindowState; savedBorder = FormBorderStyle;
                FormBorderStyle = FormBorderStyle.None; WindowState = FormWindowState.Normal; WindowState = FormWindowState.Maximized;
            }
            else { FormBorderStyle = savedBorder; WindowState = savedState; }
        }

        // like Office: ask before closing a window with unsaved work
        async void OnClosing(object sender, FormClosingEventArgs e)
        {
            if (allowClose || web.CoreWebView2 == null) return;
            e.Cancel = true;
            string dirty = "";
            try { dirty = await web.CoreWebView2.ExecuteScriptAsync("(function(){var d=document.getElementById('dirty');return d?d.textContent.trim():''})()"); } catch { }
            if (dirty != null && dirty.Trim('"', ' ').Length > 0)
            {
                var r = MessageBox.Show(this, "You have unsaved changes in \"" + Text + "\".\n\nClose without saving?", "Lumen Office", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (r != DialogResult.Yes) return;
            }
            allowClose = true;
            if (startUrl != null) Placement.Save(this, Program.App);
            BeginInvoke(new Action(Close));
        }
    }
}
