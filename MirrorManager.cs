#nullable disable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace HomeworkViewer
{
    public class MirrorManager
    {
        private readonly List<string> _mirrorSites = new List<string>
        {
            "https://bgithub.xyz",
            "https://kkgithub.com",
            "https://github.ur1.fun",
            "https://ghp.ci",
            "https://moeyy.cn/gh-proxy",
            "https://ghproxy.net",
            "https://ghproxy.homeboyc.cn",
            "https://gitclone.com",
            "https://hub.fastgit.org"
        };

        private List<string> _workingMirrors;
        private DateTime _lastTestTime = DateTime.MinValue;
        private readonly TimeSpan _cacheTTL = TimeSpan.FromMinutes(30);
        private readonly HttpClient _httpClient;

        public MirrorManager()
        {
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(3);
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "HomeworkViewer-MirrorChecker");
        }

        public async Task<string> GetFastestMirrorAsync()
        {
            var mirrors = await GetWorkingMirrorsAsync();
            return mirrors.FirstOrDefault() ?? "https://github.com";
        }

        public async Task<List<string>> GetWorkingMirrorsAsync()
        {
            if (_workingMirrors != null && DateTime.Now - _lastTestTime < _cacheTTL)
            {
                return _workingMirrors;
            }

            var testTasks = _mirrorSites.Select(async mirror =>
            {
                var latency = await TestMirrorLatencyAsync(mirror);
                return (mirror, latency);
            });

            var results = await Task.WhenAll(testTasks);

            _workingMirrors = results
                .Where(r => r.latency >= 0)
                .OrderBy(r => r.latency)
                .Select(r => r.mirror)
                .ToList();

            _lastTestTime = DateTime.Now;
            return _workingMirrors;
        }

        private async Task<int> TestMirrorLatencyAsync(string mirror)
        {
            try
            {
                var startTime = DateTime.Now;
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

        public async Task<string> GetMirroredUrlAsync(string originalUrl)
        {
            if (string.IsNullOrEmpty(originalUrl) || !originalUrl.Contains("github.com"))
                return originalUrl;

            var fastestMirror = await GetFastestMirrorAsync();
            if (fastestMirror == "https://github.com")
                return originalUrl;

            if (fastestMirror.Contains("ghproxy") || fastestMirror.Contains("gh-proxy") || fastestMirror.Contains("ghp.ci"))
            {
                return $"{fastestMirror}/{originalUrl}";
            }
            else
            {
                return originalUrl.Replace("https://github.com", fastestMirror);
            }
        }

        public void ClearCache()
        {
            _workingMirrors = null;
            _lastTestTime = DateTime.MinValue;
        }
    }
}