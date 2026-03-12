#nullable enable
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace VenioNSRLTool.Helpers
{
    public static class IniHelper
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetPrivateProfileString(
            string lpAppName,
            string lpKeyName,
            string lpDefault,
            StringBuilder lpReturnedString,
            int nSize,
            string lpFileName);

        public static string GetIni(string iniPath, string section, string key)
        {
            var sb = new StringBuilder(255);
            GetPrivateProfileString(section, key, "", sb, 255, iniPath);
            return sb.ToString().Trim();
        }

        public static string DecryptPassword(string encrypted, TextBox? log = null)
        {
            log?.AppendText($"\ud83d\udd0d Encrypted value from VenioSetup.ini: {encrypted}\n");
            if (string.IsNullOrWhiteSpace(encrypted))
            {
                log?.AppendText("\u274c Password field is empty\n");
                return "";
            }
            // Try real VenioUtility.dll
            try
            {
                string dllPath = Path.Combine(
                    Path.GetDirectoryName(Application.ExecutablePath) ?? ".",
                    "VenioUtility.dll");
                if (File.Exists(dllPath))
                {
                    log?.AppendText($" Loading VenioUtility.dll...\n");
                    var asm = Assembly.LoadFrom(dllPath);
                    var type = asm.GetType("VenioUtility.Security.CommonSecurityProvider");
                    if (type != null)
                    {
                        log?.AppendText($" Found CommonSecurityProvider\n");
                        var method = type.GetMethod("DecrptPassword", BindingFlags.Public | BindingFlags.Static);
                        if (method != null && method.GetParameters().Length == 2)
                        {
                            log?.AppendText($" Found DecrptPassword method\n");
                            var enumType = asm.GetType("VenioUtility.Security.PasswordType");
                            if (enumType != null)
                            {
                                object? enumValue = Enum.Parse(enumType, "VENIO_SETUP_INI");
                                object? result = method.Invoke(null, new object[] { encrypted, enumValue! });
                                if (result is string plain && !string.IsNullOrEmpty(plain))
                                {
                                    log?.AppendText($"\u2705 SUCCESS! Decrypted using Venio method\n");
                                    return plain;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log?.AppendText($" VenioUtility reflection failed: {ex.Message}\n");
            }
            // Fallback DPAPI
            log?.AppendText(" Trying DPAPI fallback...\n");
            string[] candidates = { encrypted.TrimStart('+'), encrypted };
            foreach (var candidate in candidates)
            {
                try
                {
                    byte[] data = Convert.FromBase64String(candidate);
                    log?.AppendText($" Base64 OK ({data.Length} bytes)\n");
                    try
                    {
                        byte[] plain = ProtectedData.Unprotect(data, null, DataProtectionScope.LocalMachine);
                        log?.AppendText($"\u2705 SUCCESS! LocalMachine worked\n");
                        return Encoding.UTF8.GetString(plain);
                    }
                    catch (Exception ex1)
                    {
                        log?.AppendText($" LocalMachine: {ex1.Message}\n");
                    }
                    try
                    {
                        byte[] plain = ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser);
                        log?.AppendText($"\u2705 SUCCESS! CurrentUser worked\n");
                        return Encoding.UTF8.GetString(plain);
                    }
                    catch (Exception ex2)
                    {
                        log?.AppendText($" CurrentUser: {ex2.Message}\n");
                    }
                }
                catch { }
            }
            log?.AppendText("\u274c All decryption failed. Will prompt for plain password.\n");
            return "";
        }
    }
}
