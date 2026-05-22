using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Xls2Dat.Core.Readers
{
    public static class LibreOfficeLocator
    {
        public static string? Locate(string? explicitPath = null)
        {
            if (!string.IsNullOrWhiteSpace(explicitPath))
                return File.Exists(explicitPath) ? explicitPath : null;

            var envPath = Environment.GetEnvironmentVariable("SOFFICE_PATH");
            if (!string.IsNullOrWhiteSpace(envPath) && File.Exists(envPath))
                return envPath;

            foreach (var candidate in WellKnownPaths())
            {
                if (File.Exists(candidate)) return candidate;
            }

            return FindOnPath(IsWindows ? "soffice.exe" : "soffice");
        }

        private static System.Collections.Generic.IEnumerable<string> WellKnownPaths()
        {
            if (IsWindows)
            {
                yield return @"C:\Program Files\LibreOffice\program\soffice.exe";
                yield return @"C:\Program Files (x86)\LibreOffice\program\soffice.exe";
            }
            else if (IsMacOS)
            {
                yield return "/Applications/LibreOffice.app/Contents/MacOS/soffice";
            }
            else
            {
                yield return "/usr/bin/soffice";
                yield return "/usr/local/bin/soffice";
                yield return "/snap/bin/libreoffice";
            }
        }

        private static string? FindOnPath(string name)
        {
            var pathVar = Environment.GetEnvironmentVariable("PATH");
            if (string.IsNullOrEmpty(pathVar)) return null;

            foreach (var dir in pathVar!.Split(Path.PathSeparator))
            {
                var candidate = Path.Combine(dir, name);
                if (File.Exists(candidate)) return candidate;
            }
            return null;
        }

        private static bool IsWindows => RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
        private static bool IsMacOS => RuntimeInformation.IsOSPlatform(OSPlatform.OSX);
    }
}
