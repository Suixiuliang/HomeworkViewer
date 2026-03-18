using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace HomeworkViewer
{
    /// <summary>
    /// GitHub镜像站管理器，自动测试可用镜像并选择最快的镜像
    /// </summary>
    public class MirrorManager
    {
        // 镜像站列表（根据搜索结果整理，可定期更新）
        private readonly List<string> _mirrorSites = new List<string>
        {
            // 直接访问型镜像站（支持浏览仓库）
            "https://bgithub.xyz",
            "https://kkgithub.com",
            "https://github.ur1.fun",
            
            // 文件加速型镜像站（专用于下载Release文件）
            "https://ghp.ci",
            "https://moeyy.cn/gh-proxy",
            "https://ghproxy.net",
            "https://ghproxy.homeboyc.cn",
            
            // 备用镜像
            "https://gitclone.com",
            "https://hub.fastgit.org"
        };

        // 缓存测试结果，避免频繁测试（30分钟缓存）
        private List<string> _workingMirrors;
        private DateTime _lastTestTime = DateTime.MinValue;
        private readonly TimeSpan _cacheTTL = TimeSpan.FromMinutes(30);
        private readonly HttpClient _httpClient;

        public MirrorManager()
        {
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(3); // 超时时间3秒
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "HomeworkViewer-MirrorChecker");
        }

        /// <summary>
        /// 获取最快的可用镜像站
        /// </summary>
        public async Task<string> GetFastestMirrorAsync()
        {
            var mirrors = await GetWorkingMirrorsAsync();
            return mirrors.FirstOrDefault() ?? "https://github.com"; // 如果没有可用镜像，返回原始GitHub
        }

        /// <summary>
        /// 获取所有可用的镜像站（按延迟排序）
        /// </summary>
        public async Task<List<string>> GetWorkingMirrorsAsync()
        {
            // 如果缓存未过期，直接返回
            if (_workingMirrors != null && DateTime.Now - _lastTestTime < _cacheTTL)
            {
                return _workingMirrors;
            }

            // 并发测试所有镜像
            var testTasks = _mirrorSites.Select(async mirror =>
            {
                var latency = await TestMirrorLatencyAsync(mirror);
                return (mirror, latency);
            });

            var results = await Task.WhenAll(testTasks);
            
            // 过滤出可用的镜像并按延迟排序
            _workingMirrors = results
                .Where(r => r.latency >= 0)
                .OrderBy(r => r.latency)
                .Select(r => r.mirror)
                .ToList();

            _lastTestTime = DateTime.Now;
            return _workingMirrors;
        }

        /// <summary>
        /// 测试单个镜像的延迟，返回毫秒数，失败返回-1
        /// </summary>
        private async Task<int> TestMirrorLatencyAsync(string mirror)
        {
            try
            {
                var startTime = DateTime.Now;
                
                // 使用HEAD请求测试，避免下载实际内容
                var request = new HttpRequestMessage(HttpMethod.Head, mirror);
                var response = await _httpClient.SendAsync(request);
                
                if (response.IsSuccessStatusCode)
                {
                    return (int)(DateTime.Now - startTime).TotalMilliseconds;
                }
                return -1;
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// 获取适用于指定GitHub URL的镜像URL
        /// </summary>
        public async Task<string> GetMirroredUrlAsync(string originalUrl)
        {
            if (string.IsNullOrEmpty(originalUrl) || !originalUrl.Contains("github.com"))
                return originalUrl;

            var fastestMirror = await GetFastestMirrorAsync();
            if (fastestMirror == "https://github.com")
                return originalUrl;

            // 根据镜像类型构造URL
            if (fastestMirror.Contains("ghproxy") || fastestMirror.Contains("gh-proxy") || fastestMirror.Contains("ghp.ci"))
            {
                // 文件加速型：需要添加代理前缀
                return $"{fastestMirror}/{originalUrl}";
            }
            else
            {
                // 直接访问型：替换域名
                return originalUrl.Replace("https://github.com", fastestMirror);
            }
        }

        /// <summary>
        /// 清除缓存，强制重新测试
        /// </summary>
        public void ClearCache()
        {
            _workingMirrors = null;
            _lastTestTime = DateTime.MinValue;
        }
    }
}