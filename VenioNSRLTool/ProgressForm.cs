#nullable enable
using System;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using VenioNSRLTool.Helpers;

namespace VenioNSRLTool
{
    public class ProgressForm : Form
    {
        private Label lblCurrent = null!;
        private Label lblLatest = null!;
        private Button btnUpgrade = null!;
        private Button btnOffline = null!;
        private Button btnCancel = null!;
        private ProgressBar progressBar = null!;
        private TextBox txtLog = null!;
        private readonly string connStr;
        private readonly string? currentVersion;
        private readonly string latestVersion;
        private readonly HttpClient httpClient;
        private readonly TextBox mainLog;

        public ProgressForm(string connStr, string? currentVersion, string latestVersion, HttpClient httpClient, TextBox mainLog)
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Segoe UI", 9F);
            this.connStr = connStr;
            this.currentVersion = currentVersion;
            this.latestVersion = latestVersion;
            this.httpClient = httpClient;
            this.mainLog = mainLog;
            Text = "NSRL Database Status";
            ClientSize = new Size(500, 400);
            MinimumSize = new Size(300, 200);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            CreateControls();
        }

        private void CreateControls()
        {
            TableLayoutPanel mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                AutoSize = true
            };
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(mainPanel);

            lblCurrent = new Label { Text = $"Current Version: {currentVersion ?? "None"}", AutoSize = true };
            mainPanel.Controls.Add(lblCurrent, 0, 0);
            lblLatest = new Label { Text = $"Latest Version: {latestVersion}", AutoSize = true };
            mainPanel.Controls.Add(lblLatest, 0, 1);

            FlowLayoutPanel buttonsPanel = new FlowLayoutPanel { AutoSize = true };
            if (currentVersion == null || !string.Equals(currentVersion, latestVersion, StringComparison.OrdinalIgnoreCase))
            {
                btnUpgrade = new Button { Text = "Upgrade to Latest", Size = new Size(150, 30), FlatStyle = FlatStyle.Flat };
                btnUpgrade.Click += async (s, e) => await PerformUpgrade(latestVersion);
                buttonsPanel.Controls.Add(btnUpgrade);
                btnOffline = new Button { Text = "Offline Upgrade from ZIP", Size = new Size(180, 30), FlatStyle = FlatStyle.Flat };
                btnOffline.Click += async (s, e) => await PerformOfflineUpgrade();
                buttonsPanel.Controls.Add(btnOffline);
            }
            mainPanel.Controls.Add(buttonsPanel, 0, 2);

            progressBar = new ProgressBar { Dock = DockStyle.Fill, Visible = false };
            mainPanel.Controls.Add(progressBar, 0, 3);
            txtLog = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, ReadOnly = true, Font = new Font("Consolas", 9) };
            mainPanel.Controls.Add(txtLog, 0, 3);

            btnCancel = new Button { Text = "Close", Size = new Size(100, 30), FlatStyle = FlatStyle.Flat };
            btnCancel.Click += (s, e) => Close();
            mainPanel.Controls.Add(btnCancel, 0, 4);
        }

        private async Task PerformUpgrade(string versionToUse)
        {
            btnUpgrade.Enabled = false;
            if (btnOffline != null) btnOffline.Enabled = false;
            progressBar.Visible = true;
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            progressBar.Value = 0;
            txtLog.AppendText("Starting upgrade...\n");

            string downloadUrl = $"https://s3.amazonaws.com/rds.nsrl.nist.gov/RDS/rds_{versionToUse}/RDS_{versionToUse}_modern_minimal.zip";
            string zipPath = await NetworkHelper.DownloadFile(httpClient, downloadUrl, txtLog, progressBar);
            if (string.IsNullOrEmpty(zipPath))
            {
                btnUpgrade.Enabled = true;
                if (btnOffline != null) btnOffline.Enabled = true;
                progressBar.Visible = false;
                return;
            }

            await Task.Run(async () => await FileHelper.ExtractAndImport(zipPath, connStr, versionToUse, true, txtLog, progressBar, mainLog));
            btnUpgrade.Enabled = true;
            if (btnOffline != null) btnOffline.Enabled = true;
            progressBar.Visible = false;
        }

        private async Task PerformOfflineUpgrade()
        {
            btnUpgrade.Enabled = false;
            btnOffline.Enabled = false;
            progressBar.Visible = true;
            progressBar.Minimum = 0;
            progressBar.Maximum = 100;
            progressBar.Value = 0;
            txtLog.AppendText("Starting offline upgrade...\n");

            using var dialog = new OpenFileDialog { Filter = "ZIP Files (*.zip)|*.zip", Title = "Select NSRL ZIP File" };
            if (dialog.ShowDialog() != DialogResult.OK)
            {
                txtLog.AppendText("No file selected.\n");
                btnUpgrade.Enabled = true;
                btnOffline.Enabled = true;
                progressBar.Visible = false;
                return;
            }

            string zipPath = dialog.FileName;
            string fileName = Path.GetFileName(zipPath);
            var match = Regex.Match(fileName, @"RDS_([\d.]+)_modern_minimal\.zip", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                txtLog.AppendText("Invalid ZIP filename format. Expected: RDS_X.X.X_modern_minimal.zip\n");
                btnUpgrade.Enabled = true;
                btnOffline.Enabled = true;
                progressBar.Visible = false;
                return;
            }

            string detectedVersion = match.Groups[1].Value;
            txtLog.AppendText($"Detected version from filename: {detectedVersion}\n");
            txtLog.AppendText("Assuming provided ZIP is correct (no validation performed).\n");

            await Task.Run(async () => await FileHelper.ExtractAndImport(zipPath, connStr, detectedVersion, false, txtLog, progressBar, mainLog));
            btnUpgrade.Enabled = true;
            btnOffline.Enabled = true;
            progressBar.Visible = false;
        }
    }
}
