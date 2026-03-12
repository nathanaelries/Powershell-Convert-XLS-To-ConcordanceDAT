#nullable enable
using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace VenioNSRLTool.Helpers
{
    public static class DatabaseHelper
    {
        public static string BuildConnectionString(string server, string database, string username, string password)
        {
            return $"Server={server};Database={database};User ID={username};Password={password};TrustServerCertificate=True;Encrypt=True;";
        }

        public static async Task<bool> DatabaseExists(string masterConnStr, string dbName)
        {
            using var conn = new SqlConnection(masterConnStr);
            await conn.OpenAsync();
            using var cmd = new SqlCommand($"SELECT db_id('{dbName}')", conn);
            return await cmd.ExecuteScalarAsync() != DBNull.Value;
        }

        public static async Task CreateDatabase(string masterConnStr, string dbName, string? mdfPath, string? ldfPath, TextBox? log = null)
        {
            using var conn = new SqlConnection(masterConnStr);
            await conn.OpenAsync();
            string sql = $"CREATE DATABASE [{dbName}]";
            if (!string.IsNullOrEmpty(mdfPath))
            {
                sql += $" ON PRIMARY (NAME = '{dbName}', FILENAME = '{mdfPath}')";
                if (!string.IsNullOrEmpty(ldfPath))
                    sql += $" LOG ON (NAME = '{dbName}_log', FILENAME = '{ldfPath}')";
            }
            using var cmd = new SqlCommand(sql, conn);
            await cmd.ExecuteNonQueryAsync();
        }

        public static async Task CreateTables(string connStr, TextBox? log = null)
        {
            try
            {
                using var dbConn = new SqlConnection(connStr);
                await dbConn.OpenAsync();
                string createMD5 = @"
                    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tbl_hs_MD5' AND schema_id = SCHEMA_ID('dbo'))
                    BEGIN
                        CREATE TABLE [dbo].[tbl_hs_MD5] (
                            [md5] CHAR(32) NULL
                        )
                    END";
                using var cmdMD5 = new SqlCommand(createMD5, dbConn);
                await cmdMD5.ExecuteNonQueryAsync();
                string createSHA1 = @"
                    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tbl_hs_SHA1' AND schema_id = SCHEMA_ID('dbo'))
                    BEGIN
                        CREATE TABLE [dbo].[tbl_hs_SHA1] (
                            [SHA1] CHAR(40) NULL
                        )
                    END";
                using var cmdSHA1 = new SqlCommand(createSHA1, dbConn);
                await cmdSHA1.ExecuteNonQueryAsync();
                string createNSRLInfo = @"
                    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tbl_hs_NSRLInfo' AND schema_id = SCHEMA_ID('dbo'))
                    BEGIN
                        CREATE TABLE [dbo].[tbl_hs_NSRLInfo] (
                            [NSRLInfoID] BIGINT NOT NULL,
                            [ISOFilePath] NVARCHAR(100) NULL,
                            [Status] NVARCHAR(50) NULL,
                            [IncludeSHA1] BIT NULL,
                            [IncludeMD5] BIT NULL,
                            [CreatedBy] INT NULL,
                            [CreatedDateTime] DATETIME NULL,
                            [isDeleteSHA1] BIT NULL,
                            [isDeleteMD5] BIT NULL,
                            [DeleteBy] INT NULL,
                            [DeleteDateTime] DATETIME NULL
                        )
                    END";
                using var cmdNSRLInfo = new SqlCommand(createNSRLInfo, dbConn);
                await cmdNSRLInfo.ExecuteNonQueryAsync();
                string createVersion = @"
                    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tbl_sys_Version' AND schema_id = SCHEMA_ID('dbo'))
                    BEGIN
                        CREATE TABLE [dbo].[tbl_sys_Version] (
                            [Id] BIGINT NOT NULL,
                            [Version] VARCHAR(25) NULL,
                            [ReleaseDate] DATETIME NULL,
                            [AppliedDate] DATETIME NULL
                        )
                    END";
                using var cmdVersion = new SqlCommand(createVersion, dbConn);
                await cmdVersion.ExecuteNonQueryAsync();
                log?.AppendText("Tables created or already exist.\n");
            }
            catch (Exception ex)
            {
                log?.AppendText($"Table creation failed: {ex.Message}\n");
            }
        }

        public static async Task<string?> GetCurrentNSRLVersion(string connStr)
        {
            try
            {
                using var conn = new SqlConnection(connStr);
                await conn.OpenAsync();
                using var cmd = new SqlCommand("SELECT TOP 1 Version FROM [tbl_sys_Version] ORDER BY Id DESC", conn);
                return (string?)await cmd.ExecuteScalarAsync();
            }
            catch
            {
                return null;
            }
        }

        public static async Task ExecuteNonQuery(SqlConnection conn, string sql, TextBox? log = null)
        {
            await using var cmd = new SqlCommand(sql, conn);
            await cmd.ExecuteNonQueryAsync();
        }
    }
}
