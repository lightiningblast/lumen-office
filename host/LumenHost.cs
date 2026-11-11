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

[assembly: System.Reflection.AssemblyTitle("Lumen Office")]
[assembly: System.Reflection.AssemblyProduct("Lumen Office")]
[assembly: System.Reflection.AssemblyCompany("Lumen Office")]
[assembly: System.Reflection.AssemblyVersion("1.1.0.0")]

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
            Size = startPage ? new Size(1000, 720) : new Size(1360, 860);
            StartPosition = FormStartPosition.WindowsDefaultLocation;
            BackColor = Color.FromArgb(243, 243, 243);
            web.Dock = DockStyle.Fill;
            web.DefaultBackgroundColor = Color.FromArgb(243, 243, 243);
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
            BeginInvoke(new Action(Close));
        }
    }
}
