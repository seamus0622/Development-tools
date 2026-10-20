// 在本機執行釋疑簡報的 Python 程式（產生簡報、對照圖說）。程式內嵌在本 DLL，隨外掛更新。C# 5 語法。
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace RfiDeckAgent.Core
{
    public class PyResult
    {
        public int ExitCode = -1;
        public string Output = "";
        public string Error = "";
        public Dictionary<string, object> Result;   // 程式最後一行 RESULT {...}
    }

    public static class PyRunner
    {
        static string SettingFile
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"DT_AIAgent\python.txt"); }
        }

        const string ResourcePrefix = "RfiDeckAgent.python.";
        static string _extracted;

        /// <summary>
        /// Python 程式位置。優先順序：
        /// 1. %APPDATA%\DT_AIAgent\scripts_dir.txt 指定的資料夾（開發者測試未發版的程式用）
        /// 2. 隨本 DLL 內嵌、解壓到 %LOCALAPPDATA%\DT_AIAgent\python\&lt;內容雜湊&gt;\ 的版本（隨外掛更新）
        /// 3. 舊版做法：AI AGENT 資料夾內的 01_AI運行本體\程式
        /// </summary>
        public static string ScriptDir
        {
            get
            {
                try
                {
                    string f = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"DT_AIAgent\scripts_dir.txt");
                    if (File.Exists(f))
                    {
                        string d = File.ReadAllText(f, Encoding.UTF8).Trim().Trim('"');
                        if (Directory.Exists(d)) return d;
                    }
                }
                catch { }
                string x = ExtractEmbedded();
                if (x != null) return x;
                return Path.Combine(Config.Root, @"01_AI運行本體\程式");
            }
        }

        /// <summary>把內嵌的 Python 程式解壓到依內容雜湊命名的資料夾；同版本只解壓一次，舊版本保留</summary>
        static string ExtractEmbedded()
        {
            if (_extracted != null && Directory.Exists(_extracted)) return _extracted;
            try
            {
                Assembly asm = typeof(PyRunner).Assembly;
                string[] names = asm.GetManifestResourceNames().Where(n => n.StartsWith(ResourcePrefix, StringComparison.Ordinal)).OrderBy(n => n, StringComparer.Ordinal).ToArray();
                if (names.Length == 0) return null;
                Dictionary<string, byte[]> files = new Dictionary<string, byte[]>();
                using (SHA256 sha = SHA256.Create())
                {
                    foreach (string n in names)
                    {
                        byte[] data;
                        using (Stream st = asm.GetManifestResourceStream(n))
                        using (MemoryStream ms = new MemoryStream())
                        {
                            st.CopyTo(ms);
                            data = ms.ToArray();
                        }
                        string file = n.Substring(ResourcePrefix.Length);
                        files[file] = data;
                        byte[] nb = Encoding.UTF8.GetBytes(file);
                        sha.TransformBlock(nb, 0, nb.Length, null, 0);
                        sha.TransformBlock(data, 0, data.Length, null, 0);
                    }
                    sha.TransformFinalBlock(new byte[0], 0, 0);
                    string hash = BitConverter.ToString(sha.Hash).Replace("-", "").Substring(0, 12).ToLowerInvariant();
                    string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"DT_AIAgent\python", hash);
                    string done = Path.Combine(dir, ".complete");
                    if (!File.Exists(done))
                    {
                        string tmp = dir + ".tmp" + Guid.NewGuid().ToString("N").Substring(0, 6);
                        Directory.CreateDirectory(tmp);
                        foreach (KeyValuePair<string, byte[]> kv in files) File.WriteAllBytes(Path.Combine(tmp, kv.Key), kv.Value);
                        File.WriteAllText(Path.Combine(tmp, ".complete"), hash);
                        try
                        {
                            // 同雜湊資料夾若存在但沒有 .complete，是先前中斷的解壓；另一個 Revit 剛解壓完成的則保留不動
                            if (Directory.Exists(dir) && !File.Exists(done)) Directory.Delete(dir, true);
                            if (!Directory.Exists(dir)) Directory.Move(tmp, dir);
                        }
                        finally
                        {
                            if (Directory.Exists(tmp)) { try { Directory.Delete(tmp, true); } catch { } }
                        }
                        if (!File.Exists(done)) return null;
                    }
                    _extracted = dir;
                    return dir;
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>找 python.exe；回傳 null 表示找不到</summary>
        public static string FindPython(out string prefixArgs)
        {
            prefixArgs = "";
            try
            {
                if (File.Exists(SettingFile))
                {
                    string s = File.ReadAllText(SettingFile, Encoding.UTF8).Trim().Trim('"');
                    if (File.Exists(s)) return s;
                }
            }
            catch { }
            List<string> roots = new List<string>();
            roots.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python"));
            roots.Add(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles));
            roots.Add(@"C:\");
            foreach (string r in roots)
            {
                if (!Directory.Exists(r)) continue;
                string[] dirs;
                try { dirs = Directory.GetDirectories(r, "Python3*"); } catch { continue; }
                foreach (string d in dirs.OrderByDescending(x => VersionKey(Path.GetFileName(x))))
                {
                    string exe = Path.Combine(d, "python.exe");
                    if (File.Exists(exe)) return exe;
                }
            }
            string[] launchers = {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Python\Launcher\py.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "py.exe") };
            foreach (string l in launchers)
            {
                if (File.Exists(l)) { prefixArgs = "-3 "; return l; }
            }
            return null;
        }

        static int VersionKey(string name)
        {
            // Python312 → 312, Python39 → 39 → 309
            string digits = new string(name.Where(char.IsDigit).ToArray());
            int v;
            if (!int.TryParse(digits, out v)) return 0;
            if (digits.Length == 2) v = (v / 10) * 100 + v % 10;
            return v;
        }

        public static PyResult Run(string args, int timeoutMs)
        {
            PyResult res = new PyResult();
            string prefix;
            string py = FindPython(out prefix);
            if (py == null)
            {
                res.Error = "找不到 Python。請安裝 Python 3（python.org），或在 %APPDATA%\\DT_AIAgent\\python.txt 填 python.exe 的完整路徑。";
                return res;
            }
            ProcessStartInfo psi = new ProcessStartInfo(py, prefix + args);
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;
            psi.StandardOutputEncoding = Encoding.UTF8;
            psi.StandardErrorEncoding = Encoding.UTF8;
            psi.WorkingDirectory = Directory.Exists(ScriptDir) ? ScriptDir : Path.GetTempPath();
            psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            psi.EnvironmentVariables["PYTHONUTF8"] = "1";
            psi.EnvironmentVariables["PYTHONDONTWRITEBYTECODE"] = "1";
            // 程式不一定放在 AI AGENT 資料夾內（可能是外掛解壓的位置），由這裡告訴 Python 資料在哪裡
            if (Config.IsConfigured) psi.EnvironmentVariables["DT_AIAGENT_ROOT"] = Config.Root;
            StringBuilder err = new StringBuilder();
            using (Process p = new Process())
            {
                p.StartInfo = psi;
                p.ErrorDataReceived += (s, e) => { if (e.Data != null) lock (err) err.AppendLine(e.Data); };
                p.Start();
                p.BeginErrorReadLine();
                res.Output = p.StandardOutput.ReadToEnd();
                if (!p.WaitForExit(timeoutMs))
                {
                    try { p.Kill(); } catch { }
                    res.Error = "執行逾時";
                    return res;
                }
                p.WaitForExit();
                res.ExitCode = p.ExitCode;
            }
            lock (err) res.Error = err.ToString();
            foreach (string line in res.Output.Split('\n').Reverse())
            {
                string t = line.Trim();
                if (t.StartsWith("RESULT "))
                {
                    try { res.Result = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(t.Substring(7)); }
                    catch { }
                    break;
                }
            }
            return res;
        }

        /// <summary>確認 python-pptx / openpyxl 已安裝，沒有就自動 pip install --user</summary>
        public static string EnsurePackages(Action<string> status)
        {
            PyResult r = Run("-c \"import pptx, openpyxl, PIL, lxml\"", 60000);
            if (r.ExitCode == 0) return null;
            if (r.ExitCode == -1 && r.Error.StartsWith("找不到")) return r.Error;
            if (status != null) status("第一次使用：安裝簡報套件（python-pptx）中，約 1 分鐘…");
            PyResult ins = Run("-m pip install --user --disable-pip-version-check python-pptx openpyxl", 300000);
            r = Run("-c \"import pptx, openpyxl, PIL, lxml\"", 60000);
            if (r.ExitCode == 0) return null;
            return "安裝 python-pptx 失敗：\n" + ins.Output + "\n" + ins.Error;
        }

        public static string Q(string s) { return "\"" + s.Replace("\"", "\\\"") + "\""; }
    }
}
