using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HomeworkViewer
{
    /// <summary>
    /// 系统架构信息
    /// </summary>
    public class SystemArchInfo
    {
        public string OS { get; set; }      // windows, linux, osx
        public string Arch { get; set; }     // x86, x64, arm64, arm
        public string Extension { get; set; } // exe, msi, zip, dmg, AppImage, deb, rpm

        public override string ToString()
        {
            return $"{OS}-{Arch}.{Extension}";
        }
    }

    /// <summary>
    /// GitHub Release资产信息
    /// </summary>
    public class ReleaseAsset
    {
        public string Name { get; set; }
        public string DownloadUrl { get; set; }
        public long Size { get; set; }
        public string ContentType { get; set; }
    }

    /// <summary>
    /// 下载辅助类
    /// </summary>
    public class DownloadHelper
    {
        private readonly HttpClient _httpClient;
        private readonly MirrorManager _mirrorManager;
        private readonly string _tempPath;

        public DownloadHelper()
        {
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromMinutes(5); // 下载超时5分钟
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "HomeworkViewer-Updater");
            
            _mirrorManager = new MirrorManager();
            
            // 创建临时目录
            _tempPath = Path.Combine(Path.GetTempPath(), "HomeworkViewer_Updates");
            if (!Directory.Exists(_tempPath))
            {
                Directory.CreateDirectory(_tempPath);
            }
        }

        /// <summary>
        /// 获取当前系统的架构信息
        /// </summary>
        public SystemArchInfo GetCurrentSystemArch()
        {
            var info = new SystemArchInfo();

            // 判断操作系统
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                info.OS = "windows";
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                info.OS = "osx";
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                info.OS = "linux";
            }
            else
            {
                info.OS = "unknown";
            }

            // 判断架构
            var arch = RuntimeInformation.OSArchitecture;
            if (arch == Architecture.X64)
            {
                info.Arch = "x64";
            }
            else if (arch == Architecture.X86)
            {
                info.Arch = "x86";
            }
            else if (arch == Architecture.Arm64)
            {
                info.Arch = "arm64";
            }
            else if (arch == Architecture.Arm)
            {
                info.Arch = "arm";
            }
            else
            {
                info.Arch = "unknown";
            }

            // 默认扩展名（后续会根据实际文件判断）
            info.Extension = "exe";

            return info;
        }

        /// <summary>
        /// 从GitHub API获取最新的Release信息
        /// </summary>
        public async Task<List<ReleaseAsset>> GetLatestReleaseAssetsAsync(string owner, string repo)
        {
            string apiUrl = $"https://api.github.com/repos/{owner}/{repo}/releases/latest";
            
            try
            {
                string json = await _httpClient.GetStringAsync(apiUrl);
                
                // 解析JSON获取assets
                using JsonDocument doc = JsonDocument.Parse(json);
                JsonElement root = doc.RootElement;
                
                if (root.TryGetProperty("assets", out JsonElement assets))
                {
                    var result = new List<ReleaseAsset>();
                    
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
                    
                    return result;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取Release信息失败: {ex.Message}");
            }
            
            return new List<ReleaseAsset>();
        }

        /// <summary>
        /// 根据当前系统匹配最合适的安装包
        /// </summary>
        public ReleaseAsset FindBestMatchAsset(List<ReleaseAsset> assets, SystemArchInfo archInfo)
        {
            if (assets == null || assets.Count == 0)
                return null;

            // 构建匹配模式
            var patterns = new List<string>();

            // 操作系统匹配模式
            patterns.Add(archInfo.OS);
            
            // 架构匹配模式
            patterns.Add(archInfo.Arch);
            
            // 扩展名匹配模式（根据操作系统）
            switch (archInfo.OS)
            {
                case "windows":
                    patterns.AddRange(new[] { "exe", "msi", "zip", "7z" });
                    break;
                case "osx":
                    patterns.AddRange(new[] { "dmg", "zip", "pkg", "app" });
                    break;
                case "linux":
                    patterns.AddRange(new[] { "AppImage", "deb", "rpm", "tar.gz", "tar.xz", "snap" });
                    break;
            }

            // 评分系统：给每个资产打分，分数越高越匹配
            var scoredAssets = assets.Select(asset => new
            {
                Asset = asset,
                Score = CalculateMatchScore(asset.Name, patterns, archInfo)
            }).Where(x => x.Score > 0)
              .OrderByDescending(x => x.Score)
              .ToList();

            return scoredAssets.FirstOrDefault()?.Asset;
        }

        /// <summary>
        /// 计算匹配分数
        /// </summary>
        private int CalculateMatchScore(string fileName, List<string> patterns, SystemArchInfo archInfo)
        {
            int score = 0;
            string lowerName = fileName.ToLower();

            // 操作系统完全匹配 +30
            if (lowerName.Contains(archInfo.OS))
                score += 30;

            // 架构完全匹配 +20
            if (lowerName.Contains(archInfo.Arch))
                score += 20;

            // 架构近似匹配（如x64匹配amd64）+10
            if (archInfo.Arch == "x64" && (lowerName.Contains("amd64") || lowerName.Contains("x86_64")))
                score += 10;
            if (archInfo.Arch == "x86" && lowerName.Contains("i386"))
                score += 10;
            if (archInfo.Arch == "arm64" && lowerName.Contains("aarch64"))
                score += 10;

            // 扩展名匹配 +15
            foreach (var ext in patterns.Skip(2)) // 跳过OS和Arch
            {
                if (lowerName.EndsWith(ext) || lowerName.Contains("." + ext))
                {
                    score += 15;
                    break;
                }
            }

            // 包含"setup"或"install" +5
            if (lowerName.Contains("setup") || lowerName.Contains("install"))
                score += 5;

            // 排除源码包
            if (lowerName.Contains("source") || lowerName.Contains("src"))
                score -= 50;

            return score;
        }

        /// <summary>
        /// 下载文件（支持镜像加速）
        /// </summary>
        public async Task<string> DownloadFileAsync(string downloadUrl, string fileName, IProgress<int> progress = null)
        {
            string localPath = Path.Combine(_tempPath, fileName);
            
            // 如果文件已存在，先删除
            if (File.Exists(localPath))
            {
                try { File.Delete(localPath); } catch { }
            }

            // 使用镜像加速
            string mirroredUrl = await _mirrorManager.GetMirroredUrlAsync(downloadUrl);

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
                        
                        while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, bytesRead);
                            totalRead += bytesRead;
                            
                            if (progress != null && totalBytes > 0)
                            {
                                int percentage = (int)((totalRead * 100) / totalBytes);
                                progress.Report(percentage);
                            }
                        }
                    }
                }
                
                return localPath;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"下载失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 打开文件或执行安装程序
        /// </summary>
        public void OpenOrInstallFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                MessageBox.Show("安装文件不存在，请手动下载。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string extension = Path.GetExtension(filePath).ToLower();
            var archInfo = GetCurrentSystemArch();

            try
            {
                // Windows平台
                if (archInfo.OS == "windows")
                {
                    if (extension == ".exe" || extension == ".msi")
                    {
                        // 启动安装程序
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = filePath,
                            UseShellExecute = true,
                            Verb = "runas" // 请求管理员权限
                        });
                        
                        // 询问是否退出当前应用
                        var result = MessageBox.Show(
                            "安装程序已启动。是否关闭当前应用以完成安装？",
                            "更新",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);
                            
                        if (result == DialogResult.Yes)
                        {
                            Application.Exit();
                        }
                    }
                    else
                    {
                        // 打开所在文件夹
                        Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                        MessageBox.Show($"文件已下载到：{filePath}\n请手动运行安装。", "下载完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                // macOS平台
                else if (archInfo.OS == "osx")
                {
                    if (extension == ".dmg" || extension == ".pkg")
                    {
                        Process.Start("open", $"\"{filePath}\"");
                    }
                    else
                    {
                        Process.Start("open", $"\"{Path.GetDirectoryName(filePath)}\"");
                    }
                }
                // Linux平台
                else
                {
                    Process.Start("xdg-open", $"\"{Path.GetDirectoryName(filePath)}\"");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法启动安装程序：{ex.Message}\n请手动打开文件：{filePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 清理临时文件
        /// </summary>
        public void Cleanup()
        {
            try
            {
                if (Directory.Exists(_tempPath))
                {
                    var files = Directory.GetFiles(_tempPath);
                    foreach (var file in files)
                    {
                        try { File.Delete(file); } catch { }
                    }
                }
            }
            catch { }
        }
    }
}