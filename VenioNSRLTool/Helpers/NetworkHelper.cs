#nullable enable
using System;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VenioNSRLTool.Helpers
{
    public static class NetworkHelper
    {
        public static async Task<string> GetLatestNSRLVersion(HttpClient httpClient, TextBox? log = null)
        {
            string nistUrl = "https://www.nist.gov/itl/ssd/software-quality-group/national-software-reference-library-nsrl/nsrl-download/current-rds";
            string html = await httpClient.GetStringAsync(nistUrl);
            var match = Regex.Match(html, @"RDS Version\s*([\d.]+)");
            return match.Success ? match.Groups[1].Value : "Unknown";
        }

        public static async Task<string> DownloadFile(HttpClient httpClient, string url, TextBox? log = null, ProgressBar? progressBar = null)
        {
            string zipPath = Path.GetTempFileName() + ".zip";
            // Check disk space via HEAD request
            try
            {
                var headRequest = new HttpRequestMessage(HttpMethod.Head, url);
                using var headResponse = await httpClient.SendAsync(headRequest);
                headResponse.EnsureSuccessStatusCode();
                if (headResponse.Content.Headers.ContentLength.HasValue)
                {
                    long requiredDownloadSpace = headResponse.Content.Headers.ContentLength.Value;
                    log?.AppendText($"Required space for download: {requiredDownloadSpace / (1024 * 1024)} MB\n");
                    string tempDir = Path.GetTempPath();
                    var driveInfo = new DriveInfo(Path.GetPathRoot(tempDir)!);
                    if (driveInfo.AvailableFreeSpace < requiredDownloadSpace * 1.1)
                    {
                        log?.AppendText("Insufficient disk space for download.\n");
                        return "";
                    }
                }
            }
            catch (Exception ex)
            {
                log?.AppendText($"Failed to check file size: {ex.Message}\n");
            }
            // Stream download with progress
            try
            {
                log?.AppendText($"Downloading NSRL data from {url}...\n");
                using var response = await httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
                response.EnsureSuccessStatusCode();
                long? totalBytes = response.Content.Headers.ContentLength;
                long downloaded = 0;
                byte[] buffer = new byte[8192];
                await using var stream = await response.Content.ReadAsStreamAsync();
                await using var fs = new FileStream(zipPath, FileMode.Create, FileAccess.Write);
                while (true)
                {
                    int read = await stream.ReadAsync(buffer);
                    if (read == 0) break;
                    await fs.WriteAsync(buffer.AsMemory(0, read));
                    downloaded += read;
                    if (totalBytes.HasValue && totalBytes > 0 && progressBar != null)
                    {
                        int progress = (int)((downloaded * 100L) / totalBytes.Value);
                        progressBar.Invoke((MethodInvoker)(() => progressBar.Value = progress));
                    }
                }
                log?.AppendText("Download complete.\n");
                return zipPath;
            }
            catch (Exception ex)
            {
                log?.AppendText($"Download failed: {ex.Message}\n");
                return "";
            }
        }
    }
}
