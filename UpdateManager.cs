using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MssqlPatientHelper
{
    public class ReleaseInfo
    {
        public string TagName { get; set; }
        public string Version { get; set; }
        public string Title { get; set; }
        public string Body { get; set; }
        public string DownloadUrl { get; set; }
        public bool HasUpdate { get; set; }
    }

    public static class UpdateManager
    {
        public const string CurrentVersion = "5.20.0";
        public const string RepoOwner = "terapark3-commits";
        public const string RepoName = "pm-helper";

        public static string LatestApiUrl
        {
            get { return string.Format("https://api.github.com/repos/{0}/{1}/releases/latest", RepoOwner, RepoName); }
        }

        public static void CheckForUpdatesAsync(Form owner, Action<ReleaseInfo> onCompleted = null, bool silent = true)
        {
            Task.Run(() =>
            {
                ReleaseInfo info = CheckForUpdatesInternal();

                if (owner != null && !owner.IsDisposed && owner.IsHandleCreated)
                {
                    owner.BeginInvoke((Action)(() =>
                    {
                        if (onCompleted != null)
                        {
                            onCompleted(info);
                        }

                        if (info != null && info.HasUpdate)
                        {
                            ShowUpdatePrompt(owner, info);
                        }
                        else if (!silent)
                        {
                            MessageBox.Show(
                                string.Format("현재 최신 버전(v{0})을 사용하고 있습니다.", CurrentVersion),
                                "업데이트 확인",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                    }));
                }
            });
        }

        public static ReleaseInfo CheckForUpdatesInternal()
        {
            ReleaseInfo info = new ReleaseInfo { HasUpdate = false };

            try
            {
                // Enable TLS 1.2 for modern GitHub API
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | SecurityProtocolType.Tls;

                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(LatestApiUrl);
                req.Method = "GET";
                req.UserAgent = "pm-helper-updater";
                req.Timeout = 10000;
                req.Accept = "application/vnd.github.v3+json";

                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (StreamReader reader = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                {
                    string json = reader.ReadToEnd();
                    ParseReleaseJson(json, info);
                }

                if (!string.IsNullOrEmpty(info.Version))
                {
                    info.HasUpdate = IsNewer(info.Version, CurrentVersion);
                }
            }
            catch (Exception ex)
            {
                info.Body = "업데이트 확인 실패: " + ex.Message;
            }

            return info;
        }

        private static void ParseReleaseJson(string json, ReleaseInfo info)
        {
            // Extract tag_name
            Match mTag = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"([^\"]+)\"");
            if (mTag.Success)
            {
                info.TagName = mTag.Groups[1].Value;
                info.Version = info.TagName.TrimStart('v', 'V');
            }

            // Extract title / name
            Match mName = Regex.Match(json, "\"name\"\\s*:\\s*\"([^\"]+)\"");
            if (mName.Success)
            {
                info.Title = mName.Groups[1].Value;
            }

            // Extract body (release notes)
            Match mBody = Regex.Match(json, "\"body\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
            if (mBody.Success)
            {
                string rawBody = mBody.Groups[1].Value;
                info.Body = Regex.Unescape(rawBody);
            }

            // Extract exe download url
            Match mUrl = Regex.Match(json, "\"browser_download_url\"\\s*:\\s*\"(https://[^\"]+?\\.exe)\"");
            if (mUrl.Success)
            {
                info.DownloadUrl = mUrl.Groups[1].Value;
            }
            else
            {
                // Fallback to releases page URL if direct exe asset isn't attached
                info.DownloadUrl = string.Format("https://github.com/{0}/{1}/releases/latest", RepoOwner, RepoName);
            }
        }

        public static bool IsNewer(string remoteVer, string localVer)
        {
            try
            {
                Version vRemote = NormalizeVersion(remoteVer);
                Version vLocal = NormalizeVersion(localVer);
                return vRemote > vLocal;
            }
            catch
            {
                return string.Compare(remoteVer, localVer, StringComparison.OrdinalIgnoreCase) > 0;
            }
        }

        private static Version NormalizeVersion(string vStr)
        {
            string clean = Regex.Replace(vStr ?? "0.0.0", "[^0-9.]", "").Trim('.');
            string[] parts = clean.Split('.');
            int major = parts.Length > 0 ? int.Parse(parts[0]) : 0;
            int minor = parts.Length > 1 ? int.Parse(parts[1]) : 0;
            int build = parts.Length > 2 ? int.Parse(parts[2]) : 0;
            int rev = parts.Length > 3 ? int.Parse(parts[3]) : 0;
            return new Version(major, minor, build, rev);
        }

        public static void ShowUpdatePrompt(Form owner, ReleaseInfo info)
        {
            using (Form dlg = new Form())
            {
                dlg.Text = "새 버전 업데이트 안내";
                dlg.Size = new Size(520, 420);
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;
                dlg.BackColor = Color.FromArgb(15, 23, 42); // Slate 900
                dlg.ForeColor = Color.FromArgb(248, 250, 252);
                dlg.Font = new Font("맑은 고딕", 9.5F);

                try
                {
                    dlg.Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                }
                catch { }

                Label lblHeader = new Label
                {
                    Text = string.Format("🎉 새로운 버전 [ v{0} ] 이 출시되었습니다!", info.Version),
                    Location = new Point(20, 18),
                    Size = new Size(460, 28),
                    Font = new Font("맑은 고딕", 12F, FontStyle.Bold),
                    ForeColor = Color.FromArgb(99, 102, 241) // Indigo 500
                };
                dlg.Controls.Add(lblHeader);

                Label lblSub = new Label
                {
                    Text = string.Format("현재 버전: v{0}  ➔  최신 버전: v{1}", CurrentVersion, info.Version),
                    Location = new Point(22, 50),
                    Size = new Size(460, 20),
                    ForeColor = Color.FromArgb(148, 163, 184) // Slate 400
                };
                dlg.Controls.Add(lblSub);

                Label lblNotes = new Label
                {
                    Text = "■ 업데이트 변경 내역 (Release Notes):",
                    Location = new Point(22, 80),
                    Size = new Size(460, 20),
                    Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                    ForeColor = Color.FromArgb(226, 232, 240)
                };
                dlg.Controls.Add(lblNotes);

                TextBox txtNotes = new TextBox
                {
                    Location = new Point(22, 105),
                    Size = new Size(460, 200),
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    BackColor = Color.FromArgb(30, 41, 59),
                    ForeColor = Color.FromArgb(226, 232, 240),
                    BorderStyle = BorderStyle.FixedSingle,
                    Text = string.IsNullOrEmpty(info.Body) ? "상세 변경 내역이 제공되지 않았습니다." : info.Body.Replace("\n", "\r\n")
                };
                dlg.Controls.Add(txtNotes);

                Button btnUpdateNow = new Button
                {
                    Text = "⚡ 지금 자동 업데이트",
                    Location = new Point(190, 325),
                    Size = new Size(180, 36),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(16, 185, 129), // Emerald 500
                    ForeColor = Color.White,
                    Font = new Font("맑은 고딕", 9.5F, FontStyle.Bold),
                    DialogResult = DialogResult.OK
                };
                btnUpdateNow.FlatAppearance.BorderSize = 0;
                dlg.Controls.Add(btnUpdateNow);

                Button btnLater = new Button
                {
                    Text = "나중에",
                    Location = new Point(380, 325),
                    Size = new Size(100, 36),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.FromArgb(51, 65, 85), // Slate 700
                    ForeColor = Color.White,
                    DialogResult = DialogResult.Cancel
                };
                btnLater.FlatAppearance.BorderSize = 0;
                dlg.Controls.Add(btnLater);

                if (dlg.ShowDialog(owner) == DialogResult.OK)
                {
                    StartDownloadAndUpdate(owner, info);
                }
            }
        }

        private static void StartDownloadAndUpdate(Form owner, ReleaseInfo info)
        {
            if (string.IsNullOrEmpty(info.DownloadUrl) || !info.DownloadUrl.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            {
                // Open browser if direct exe is not linked
                Process.Start(info.DownloadUrl);
                return;
            }

            Form progDlg = new Form
            {
                Text = "새 버전 다운로드 중...",
                Size = new Size(420, 160),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(15, 23, 42),
                ForeColor = Color.White,
                Font = new Font("맑은 고딕", 9.5F)
            };

            Label lblStatus = new Label
            {
                Text = "최신 버전을 다운로드하고 있습니다. 잠시만 기다려주십시오...",
                Location = new Point(20, 20),
                Size = new Size(370, 25),
                ForeColor = Color.FromArgb(148, 163, 184)
            };
            progDlg.Controls.Add(lblStatus);

            ProgressBar pb = new ProgressBar
            {
                Location = new Point(20, 50),
                Size = new Size(365, 25),
                Style = ProgressBarStyle.Continuous
            };
            progDlg.Controls.Add(pb);

            string tempDir = Path.Combine(Path.GetTempPath(), "pm_helper_update");
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
            string tempExe = Path.Combine(tempDir, "pm+helper_new.exe");

            WebClient client = new WebClient();
            client.Headers.Add("User-Agent", "pm-helper-updater");

            client.DownloadProgressChanged += (s, e) =>
            {
                pb.Value = e.ProgressPercentage;
                lblStatus.Text = string.Format("다운로드 중: {0}% ({1:N1} MB / {2:N1} MB)",
                    e.ProgressPercentage,
                    (double)e.BytesReceived / 1024 / 1024,
                    (double)e.TotalBytesToReceive / 1024 / 1024);
            };

            client.DownloadFileCompleted += (s, e) =>
            {
                progDlg.Close();

                if (e.Error != null)
                {
                    MessageBox.Show("업데이트 다운로드 실패:\n" + e.Error.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Execute safe swap
                ExecuteInPlaceSwap(tempExe);
            };

            progDlg.Shown += (s, e) =>
            {
                try
                {
                    client.DownloadFileAsync(new Uri(info.DownloadUrl), tempExe);
                }
                catch (Exception ex)
                {
                    progDlg.Close();
                    MessageBox.Show("다운로드 시작 실패: " + ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            progDlg.ShowDialog(owner);
        }

        public static void ExecuteInPlaceSwap(string newExePath)
        {
            string currentExe = Application.ExecutablePath;
            int currentPid = Process.GetCurrentProcess().Id;
            string updaterBat = Path.Combine(Path.GetTempPath(), "pm_updater_" + currentPid + ".bat");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("@echo off");
            sb.AppendLine("chcp 65001 > nul");
            sb.AppendLine("setlocal");
            sb.AppendLine(string.Format("set \"TARGET={0}\"", currentExe));
            sb.AppendLine(string.Format("set \"NEW_FILE={0}\"", newExePath));
            sb.AppendLine(string.Format("set \"PID={0}\"", currentPid));
            sb.AppendLine("");
            sb.AppendLine(":: 1. 기존 프로세스 종료 대기");
            sb.AppendLine(":wait_loop");
            sb.AppendLine("timeout /t 1 /nobreak > nul");
            sb.AppendLine("tasklist /fi \"pid eq %PID%\" 2>nul | find \"%PID%\" > nul");
            sb.AppendLine("if not errorlevel 1 goto wait_loop");
            sb.AppendLine("");
            sb.AppendLine(":: 2. 기존 실행 파일 백업");
            sb.AppendLine("if exist \"%TARGET%.bak\" del /f /q \"%TARGET%.bak\"");
            sb.AppendLine("if exist \"%TARGET%\" move /y \"%TARGET%\" \"%TARGET%.bak\" > nul");
            sb.AppendLine("");
            sb.AppendLine(":: 3. 새 실행 파일 교체");
            sb.AppendLine("move /y \"%NEW_FILE%\" \"%TARGET%\" > nul");
            sb.AppendLine("");
            sb.AppendLine(":: 4. 새 버전 자동 실행");
            sb.AppendLine("start \"\" \"%TARGET%\"");
            sb.AppendLine("");
            sb.AppendLine(":: 5. 임시 updater 스크립트 자체 삭제");
            sb.AppendLine("(goto) 2>nul & del /f /q \"%~f0\"");

            File.WriteAllText(updaterBat, sb.ToString(), Encoding.GetEncoding(949));

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/c \"" + updaterBat + "\"",
                CreateNoWindow = true,
                UseShellExecute = false
            };

            Process.Start(psi);
            Application.Exit();
        }
    }
}

