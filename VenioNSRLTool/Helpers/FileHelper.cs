#nullable enable
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;

namespace VenioNSRLTool.Helpers
{
    public static class FileHelper
    {
        public static async Task ExtractAndImport(
            string zipPath,
            string connStr,
            string versionToUse,
            bool isDownloaded,
            TextBox? txtLog = null,
            ProgressBar? progressBar = null,
            TextBox? mainLog = null)
        {
            // Check space for extraction
            long requiredExtractSpace = 0;
            try
            {
                using ZipArchive archive = ZipFile.OpenRead(zipPath);
                requiredExtractSpace = archive.Entries.Sum(e => e.Length);
                txtLog?.AppendText($"Required space for extraction: {requiredExtractSpace / (1024 * 1024)} MB\n");
                string tempDir = Path.GetTempPath();
                var driveInfo = new DriveInfo(Path.GetPathRoot(tempDir)!);
                if (driveInfo.AvailableFreeSpace < requiredExtractSpace * 1.1)
                {
                    txtLog?.AppendText("Insufficient disk space for extraction.\n");
                    if (isDownloaded) File.Delete(zipPath);
                    return;
                }
            }
            catch (Exception ex)
            {
                txtLog?.AppendText($"Failed to check extraction space: {ex.Message}. Proceeding with caution.\n");
            }

            string extractPath = Path.Combine(Path.GetTempPath(), "NSRLExtract_" + Guid.NewGuid().ToString());
            Directory.CreateDirectory(extractPath);
            string dbPath = "";

            try
            {
                txtLog?.AppendText("Extracting modern_minimal RDS...\n");
                if (progressBar != null)
                    progressBar.Value = 0;
                using ZipArchive archive = ZipFile.OpenRead(zipPath);
                long totalBytes = archive.Entries.Sum(e => e.Length);
                long extracted = 0;
                byte[] buffer = new byte[8192];
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    string fullPath = Path.Combine(extractPath, entry.FullName);
                    Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
                    if (entry.Name == "") continue;
                    await using var entryStream = entry.Open();
                    await using var fs = new FileStream(fullPath, FileMode.Create);
                    int read;
                    while ((read = await entryStream.ReadAsync(buffer)) > 0)
                    {
                        await fs.WriteAsync(buffer.AsMemory(0, read));
                        extracted += read;
                        if (progressBar != null)
                        {
                            int prog = (int)(extracted * 100L / totalBytes);
                            progressBar.Invoke((MethodInvoker)(() => progressBar.Value = prog));
                        }
                    }
                }
                dbPath = Directory.GetFiles(extractPath, "*.db", SearchOption.AllDirectories).FirstOrDefault() ?? "";
                if (string.IsNullOrEmpty(dbPath))
                {
                    txtLog?.AppendText("\u274c No .db file found in ZIP!\n");
                    return;
                }
                txtLog?.AppendText($"Extraction complete \u2192 {Path.GetFileName(dbPath)}\n");
            }
            catch (Exception ex)
            {
                txtLog?.AppendText($"Extraction failed: {ex.Message}\n");
                return;
            }

            // Validate file hash if signatures.txt exists
            ValidateFileHashes(extractPath, dbPath, txtLog);

            if (string.IsNullOrEmpty(dbPath)) return;
            txtLog?.AppendText("\ud83d\ude80 Starting import into VenioNSRL...\n");
            if (progressBar != null)
                progressBar.Value = 0;

            using var sqlConn = new SqlConnection(connStr);
            await sqlConn.OpenAsync();

            // 1. Clear old hashes (full replace on every RDS update)
            await DatabaseHelper.ExecuteNonQuery(sqlConn, "TRUNCATE TABLE [tbl_hs_MD5]");
            await DatabaseHelper.ExecuteNonQuery(sqlConn, "TRUNCATE TABLE [tbl_hs_SHA1]");
            txtLog?.AppendText("\u2705 Cleared old MD5/SHA1 hash sets\n");

            // 2. Open SQLite (read-only for speed)
            using var sqliteConn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
            await sqliteConn.OpenAsync();

            // Get exact distinct counts for accurate progress
            long distinctMD5 = (long)(await new SqliteCommand("SELECT COUNT(DISTINCT md5) FROM FILE", sqliteConn).ExecuteScalarAsync() ?? 0);
            long distinctSHA1 = (long)(await new SqliteCommand("SELECT COUNT(DISTINCT sha1) FROM FILE", sqliteConn).ExecuteScalarAsync() ?? 0);
            txtLog?.AppendText($"Importing {distinctMD5:N0} distinct MD5 + {distinctSHA1:N0} distinct SHA1 hashes...\n");

            // 3. Bulk-copy MD5 (0-50%)
            txtLog?.AppendText("Bulk-copying tbl_hs_MD5...\n");
            using (var readerMD5 = await new SqliteCommand("SELECT DISTINCT md5 FROM FILE WHERE md5 <> ''", sqliteConn).ExecuteReaderAsync())
            using (var bulkMD5 = new SqlBulkCopy(sqlConn))
            {
                bulkMD5.DestinationTableName = "[tbl_hs_MD5]";
                bulkMD5.BatchSize = 100_000;
                bulkMD5.NotifyAfter = 100_000;
                bulkMD5.SqlRowsCopied += (s, e) =>
                {
                    if (progressBar != null)
                    {
                        int percent = (int)((e.RowsCopied * 50L) / distinctMD5);
                        progressBar.Invoke((MethodInvoker)(() => progressBar.Value = Math.Min(percent, 50)));
                    }
                };
                await bulkMD5.WriteToServerAsync(readerMD5);
            }
            txtLog?.AppendText($"\u2705 Imported {distinctMD5:N0} MD5 hashes\n");

            // 4. Bulk-copy SHA1 (50-100%)
            txtLog?.AppendText("Bulk-copying tbl_hs_SHA1...\n");
            using (var readerSHA1 = await new SqliteCommand("SELECT DISTINCT sha1 FROM FILE WHERE sha1 <> ''", sqliteConn).ExecuteReaderAsync())
            using (var bulkSHA1 = new SqlBulkCopy(sqlConn))
            {
                bulkSHA1.DestinationTableName = "[tbl_hs_SHA1]";
                bulkSHA1.BatchSize = 100_000;
                bulkSHA1.NotifyAfter = 100_000;
                bulkSHA1.SqlRowsCopied += (s, e) =>
                {
                    if (progressBar != null)
                    {
                        int percent = 50 + (int)((e.RowsCopied * 50L) / distinctSHA1);
                        progressBar.Invoke((MethodInvoker)(() => progressBar.Value = Math.Min(percent, 100)));
                    }
                };
                await bulkSHA1.WriteToServerAsync(readerSHA1);
            }
            txtLog?.AppendText($"\u2705 Imported {distinctSHA1:N0} SHA1 hashes\n");

            // 5. Insert version row into tbl_sys_Version
            using var versionCmd = new SqlCommand("SELECT ISNULL(MAX(Id),0) + 1 FROM [tbl_sys_Version]", sqlConn);
            object? scalarVersion = await versionCmd.ExecuteScalarAsync();
            long nextVersionId = (long)scalarVersion!;
            DateTime releaseDate = DateTime.UtcNow;
            using (var cmdDate = new SqliteCommand("SELECT release_date FROM VERSION ORDER BY build_date DESC LIMIT 1", sqliteConn))
            {
                var obj = await cmdDate.ExecuteScalarAsync();
                if (obj is DateTime dt) releaseDate = dt;
            }
            using (var insertVer = new SqlCommand(
                "INSERT INTO [tbl_sys_Version] (Id, Version, ReleaseDate, AppliedDate) VALUES (@id,@ver,@rel,@app)", sqlConn))
            {
                insertVer.Parameters.AddWithValue("@id", nextVersionId);
                insertVer.Parameters.AddWithValue("@ver", versionToUse);
                insertVer.Parameters.AddWithValue("@rel", releaseDate);
                insertVer.Parameters.AddWithValue("@app", DateTime.UtcNow);
                await insertVer.ExecuteNonQueryAsync();
            }
            txtLog?.AppendText($"\u2705 Version record inserted: {versionToUse} (Id {nextVersionId})\n");

            // 6. Log the import in tbl_hs_NSRLInfo
            using var infoCmd = new SqlCommand("SELECT ISNULL(MAX(NSRLInfoID),0) + 1 FROM [tbl_hs_NSRLInfo]", sqlConn);
            object? scalarInfo = await infoCmd.ExecuteScalarAsync();
            long nextInfoId = (long)scalarInfo!;
            using (var insertInfo = new SqlCommand(@"
                INSERT INTO [tbl_hs_NSRLInfo] 
                    (NSRLInfoID, ISOFilePath, Status, IncludeSHA1, IncludeMD5, CreatedBy, CreatedDateTime, 
                     isDeleteSHA1, isDeleteMD5, DeleteBy, DeleteDateTime)
                VALUES (@id, @path, 'Imported', 1, 1, 0, GETDATE(), 0, 0, NULL, NULL)", sqlConn))
            {
                insertInfo.Parameters.AddWithValue("@id", nextInfoId);
                insertInfo.Parameters.AddWithValue("@path", $"RDS_{versionToUse}_modern_minimal.zip");
                await insertInfo.ExecuteNonQueryAsync();
            }
            txtLog?.AppendText($"\u2705 NSRLInfo record created (Id {nextInfoId})\n");

            // Cleanup
            if (isDownloaded) File.Delete(zipPath);
            Directory.Delete(extractPath, true);
            txtLog?.AppendText("\ud83c\udf89 Import finished successfully!\n");
            mainLog?.AppendText(txtLog?.Text ?? "");
        }

        private static void ValidateFileHashes(string extractPath, string dbPath, TextBox? txtLog)
        {
            string? signaturesPath = Directory.GetFiles(extractPath, "signatures.txt", SearchOption.AllDirectories).FirstOrDefault();
            if (!string.IsNullOrEmpty(signaturesPath))
            {
                txtLog?.AppendText("Validating file hashes...\n");
                bool validated = false;
                string dbFileName = Path.GetFileName(dbPath);
                var lines = File.ReadAllLines(signaturesPath);
                string? expectedSHA1 = null;
                string? expectedMD5 = null;
                foreach (var line in lines)
                {
                    if (line.StartsWith("SHA1(") && line.Contains(dbFileName))
                    {
                        var parts = line.Split('=');
                        if (parts.Length == 2)
                            expectedSHA1 = parts[1].Trim().ToLower();
                    }
                    else if (line.StartsWith("MD5(") && line.Contains(dbFileName))
                    {
                        var parts = line.Split('=');
                        if (parts.Length == 2)
                            expectedMD5 = parts[1].Trim().ToLower();
                    }
                }
                // Compute SHA1
                if (!string.IsNullOrEmpty(expectedSHA1))
                {
                    using var sha1 = SHA1.Create();
                    using var fileStream = File.OpenRead(dbPath);
                    byte[] hashBytes = sha1.ComputeHash(fileStream);
                    string computedSHA1 = BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
                    if (computedSHA1 == expectedSHA1)
                    {
                        txtLog?.AppendText("SHA1 validation successful.\n");
                        validated = true;
                    }
                    else
                    {
                        txtLog?.AppendText("SHA1 validation failed.\n");
                        return;
                    }
                }
                // Compute MD5 if present
                if (!string.IsNullOrEmpty(expectedMD5))
                {
                    using var md5 = MD5.Create();
                    using var fileStream = File.OpenRead(dbPath);
                    byte[] hashBytes = md5.ComputeHash(fileStream);
                    string computedMD5 = BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
                    if (computedMD5 == expectedMD5)
                    {
                        txtLog?.AppendText("MD5 validation successful.\n");
                        validated = true;
                    }
                    else
                    {
                        txtLog?.AppendText("MD5 validation failed.\n");
                        return;
                    }
                }
                if (!validated)
                {
                    txtLog?.AppendText("No applicable hashes found in signatures.txt for validation.\n");
                }
            }
            else
            {
                txtLog?.AppendText("No signatures.txt found; skipping validation.\n");
            }
        }
    }
}
