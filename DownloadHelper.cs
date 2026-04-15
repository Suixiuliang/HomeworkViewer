// Copyright (c) 2026 MaxSui 隋修梁. All rights reserved.
// Licensed under the GPL3.0 License. See LICENSE in the project root for license information.
#nullable disable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HomeworkViewer
{
    public class DownloadProgressInfo
    {
        public long BytesReceived { get; set; }
        public long TotalBytes { get; set; }
        public int Percentage { get; set; }
        public double SpeedKBps { get; set; }
        public TimeSpan TimeRemaining { get; set; }
    }

    public class ReleaseAsset
    {
        public string Name { get; set; }
        public string DownloadUrl { get; set; }
        public long Size { get; set; }
        public string ContentType { get; set; }
    }

    public class DownloadHelper
    {
        private readonly HttpClient _httpClient;
        private readonly MirrorManager _mirrorManager;
        private readonly string _tempPath;

        public DownloadHelper()
        {
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromMinutes(10);
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "HomeworkViewer-Updater");

            _mirrorManager = new MirrorManager();

            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _tempPath = Path.Combine(appData, "HomeworkViewerUpgrader");
            if (!Directory.Exists(_tempPath))
            {
                Directory.CreateDirectory(_tempPath);
            }
        }

        public async Task<List<ReleaseAsset>> GetLatestReleaseAssetsAsync(string owner, string repo)
        {
            string apiUrl = $"https://api.github.com/repos/{owner}/{repo}/releases/latest";
            var result = new List<ReleaseAsset>();

            try
            {
                string json = await _httpClient.GetStringAsync(apiUrl);
                using JsonDocument doc = JsonDocument.Parse(json);
                JsonElement root = doc.RootElement;
                if (root.TryGetProperty("assets", out JsonElement assets))
                {
                    foreach (var asset in assets.EnumerateArray())
                    {
                        var releaseAsset = new ReleaseAsset
                        {
                            Name = asset.GetProperty("name").GetString(),
                            DownloadUrl = asset.GetProperty("browser_download_url").GetString(),
                            Size = asset.GetProperty("size").GetInt64(),
                            ContentType = asset.GetProperty("content_type").GetString()
                        };
                        result.Add(releaseAsset);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取Release信息失败: {ex.Message}");
            }
            return result;
        }

        public ReleaseAsset FindBestMatchAsset(List<ReleaseAsset> assets)
        {
            if (assets == null || assets.Count == 0) return null;

            // 简单匹配：优先选择包含 windows/win 且扩展名为 .exe/.msi 的文件
            var candidates = assets.Where(a =>
            {
                string name = a.Name.ToLower();
                return (name.Contains("windows") || name.Contains("win")) &&
                       (name.EndsWith(".exe") || name.EndsWith(".msi") || name.EndsWith(".zip"));
            }).ToList();

            if (candidates.Count == 0)
                candidates = assets.Where(a => a.Name.ToLower().EndsWith(".exe") || a.Name.ToLower().EndsWith(".msi")).ToList();

            return candidates.OrderByDescending(a => a.Name.Length).FirstOrDefault();
        }

        public async Task<string> DownloadFileAsync(string downloadUrl, string fileName, string expectedHash = null, IProgress<DownloadProgressInfo> progress = null)
        {
            string localPath = Path.Combine(_tempPath, fileName);
            if (File.Exists(localPath)) try { File.Delete(localPath); } catch { }

            string mirroredUrl = await _mirrorManager.GetMirroredUrlAsync(downloadUrl);
            int retry = 0;
            Exception lastException = null;

            while (retry < 3)
            {
                try
                {
                    using (var response = await _httpClient.GetAsync(mirroredUrl, HttpCompletionOption.ResponseHeadersRead))
                    {
                        response.EnsureSuccessStatusCode();
                        long totalBytes = response.Content.Headers.ContentLength ?? -1;
                        using (var contentStream = await response.Content.ReadAsStreamAsync())
                        using (var fileStream = new FileStream(localPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                        {
                            var buffer = new byte[8192];
                            long totalRead = 0;
                            int bytesRead;
                            var lastReportTime = DateTime.Now;
                            long lastReportBytes = 0;

                            while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                            {
                                await fileStream.WriteAsync(buffer, 0, bytesRead);
                                totalRead += bytesRead;

                                if (progress != null && totalBytes > 0)
                                {
                                    var now = DateTime.Now;
                                    if ((now - lastReportTime).TotalMilliseconds >= 100)
                                    {
                                        int percentage = (int)((totalRead * 100) / totalBytes);
                                        double speedKBps = 0;
                                        TimeSpan timeRemaining = TimeSpan.Zero;
                                        if (lastReportBytes > 0)
                                        {
                                            double bytesPerSec = (totalRead - lastReportBytes) / (now - lastReportTime).TotalSeconds;
                                            speedKBps = bytesPerSec / 1024.0;
                                            if (speedKBps > 0)
                                            {
                                                double remainingBytes = totalBytes - totalRead;
                                                timeRemaining = TimeSpan.FromSeconds(remainingBytes / (speedKBps * 1024));
                                            }
                                        }
                                        var info = new DownloadProgressInfo
                                        {
                                            BytesReceived = totalRead,
                                            TotalBytes = totalBytes,
                                            Percentage = percentage,
                                            SpeedKBps = speedKBps,
                                            TimeRemaining = timeRemaining
                                        };
                                        progress.Report(info);
                                        lastReportBytes = totalRead;
                                        lastReportTime = now;
                                    }
                                }
                            }
                            if (progress != null && totalBytes > 0)
                            {
                                var finalInfo = new DownloadProgressInfo { BytesReceived = totalBytes, TotalBytes = totalBytes, Percentage = 100, SpeedKBps = 0, TimeRemaining = TimeSpan.Zero };
                                progress.Report(finalInfo);
                            }
                        }
                    }

                    // 哈希校验
                    if (!string.IsNullOrEmpty(expectedHash))
                    {
                        string actualHash = await Task.Run(() => ComputeSHA256(localPath));
                        if (!string.Equals(actualHash, expectedHash, StringComparison.OrdinalIgnoreCase))
                        {
                            File.Delete(localPath);
                            throw new Exception($"哈希校验失败：期望 {expectedHash}，实际 {actualHash}");
                        }
                    }
                    return localPath;
                }
                catch (HttpRequestException ex) when (mirroredUrl != downloadUrl)
                {
                    // 镜像失败，回退到原始URL
                    mirroredUrl = downloadUrl;
                    lastException = ex;
                    retry++;
                    await Task.Delay(1000);
                    continue;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    retry++;
                    if (retry >= 3) throw;
                    await Task.Delay(1000 * retry);
                }
            }
            throw lastException ?? new Exception("下载失败");
        }

        private string ComputeSHA256(string filePath)
        {
            using (var sha256 = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hashBytes = sha256.ComputeHash(stream);
                return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
            }
        }

        public void OpenOrInstallFile(string filePath)
        {
            if (!File.Exists(filePath)) return;
            string ext = Path.GetExtension(filePath).ToLower();
            if (ext == ".exe" || ext == ".msi")
            {
                Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true, Verb = "runas" });
            }
            else
            {
                Process.Start("explorer.exe", $"/select,\"{filePath}\"");
            }
        }

        public void Cleanup()
        {
            try
            {
                if (Directory.Exists(_tempPath))
                {
                    foreach (var file in Directory.GetFiles(_tempPath)) try { File.Delete(file); } catch { }
                }
            }
            catch { }
        }

        public string GetTempPath() => _tempPath;
    }
}