using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace Xls2Dat.Core.Readers
{
    public sealed class LibreOfficeReader : ISpreadsheetReader
    {
        private readonly string _tempDir;
        private readonly string _convertedXlsx;
        private readonly OpenXmlReader _inner;

        public LibreOfficeReader(string inputPath, string sofficePath, TimeSpan? timeout = null)
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "xls2dat-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);

            try
            {
                _convertedXlsx = ConvertToXlsx(inputPath, sofficePath, _tempDir, timeout ?? TimeSpan.FromMinutes(5));
                _inner = new OpenXmlReader(_convertedXlsx);
            }
            catch
            {
                SafeCleanup();
                throw;
            }
        }

        public IEnumerable<Sheet> ReadSheets() => _inner.ReadSheets();

        public void Dispose()
        {
            _inner.Dispose();
            SafeCleanup();
        }

        private static string ConvertToXlsx(string inputPath, string sofficePath, string outDir, TimeSpan timeout)
        {
            var profileDir = Path.Combine(outDir, "profile");
            Directory.CreateDirectory(profileDir);

            var psi = new ProcessStartInfo(sofficePath)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            };
            psi.ArgumentList.Add($"-env:UserInstallation=file://{profileDir.Replace('\\', '/')}");
            psi.ArgumentList.Add("--headless");
            psi.ArgumentList.Add("--norestore");
            psi.ArgumentList.Add("--nologo");
            psi.ArgumentList.Add("--nofirststartwizard");
            psi.ArgumentList.Add("--convert-to");
            psi.ArgumentList.Add("xlsx");
            psi.ArgumentList.Add("--outdir");
            psi.ArgumentList.Add(outDir);
            psi.ArgumentList.Add(inputPath);

            using var proc = Process.Start(psi) ?? throw new InvalidOperationException("Failed to launch LibreOffice.");
            if (!proc.WaitForExit((int)timeout.TotalMilliseconds))
            {
                try { proc.Kill(entireProcessTree: true); } catch { }
                throw new TimeoutException($"LibreOffice conversion timed out after {timeout.TotalSeconds:N0}s.");
            }

            if (proc.ExitCode != 0)
            {
                var stderr = proc.StandardError.ReadToEnd();
                throw new InvalidOperationException($"LibreOffice exited with code {proc.ExitCode}: {stderr.Trim()}");
            }

            var produced = Path.Combine(outDir, Path.GetFileNameWithoutExtension(inputPath) + ".xlsx");
            if (!File.Exists(produced))
                throw new FileNotFoundException("LibreOffice did not produce the expected output.", produced);

            return produced;
        }

        private void SafeCleanup()
        {
            try { if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, recursive: true); } catch { }
        }
    }
}
