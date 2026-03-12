#nullable enable
using System;
using System.Drawing;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;
using VenioNSRLTool.Helpers;

namespace VenioNSRLTool
{
    public partial class MainForm : Form
    {
        private readonly string iniPath = @"C:\Program Files\Venio\VenioFPR\VenioSetup.ini";
        private string sqlPassword = "";
        private readonly HttpClient httpClient = new HttpClient();

        // Controls
        private CheckBox chkDefaultDB = null!;
        private TextBox txtDBName = null!;
        private CheckBox chkDefaultLocation = null!;
        private TextBox txtMdf = null!;
        private Button btnBrowseMdf = null!;
        private TextBox txtLdf = null!;
        private Button btnBrowseLdf = null!;
        private TextBox txtServer = null!;
        private TextBox txtUsername = null!;
        private TextBox txtPassword = null!;
        private Button btnTestConnection = null!;
        private Button btnNext = null!;
        private Button btnClose = null!;
        private Label lblStatus = null!;
        private TextBox txtLog = null!;

        public MainForm()
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Segoe UI", 9F);
            Text = "Create/Update Venio NSRL Database";
            ClientSize = new Size(620, 550);
            MinimumSize = new Size(400, 300);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            BackColor = Color.White;
            CreateScreenshotControls();
            LoadValuesFromIni();
        }

        private void CreateScreenshotControls()
        {
            TableLayoutPanel mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                AutoSize = true
            };
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(mainPanel);

            // Title
            var title = new Label { Text = "Create/Update Venio NSRL Database", AutoSize = true, Font = new Font("Segoe UI", 12, FontStyle.Bold) };
            mainPanel.Controls.Add(title, 0, 0);

            // Database options panel
            FlowLayoutPanel dbOptionsPanel = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown };
            chkDefaultDB = new CheckBox { Text = "Use Default Venio database", Checked = true, AutoSize = true };
            dbOptionsPanel.Controls.Add(chkDefaultDB);
            var dbNamePanel = new FlowLayoutPanel { AutoSize = true };
            var lblDBName = new Label { Text = "Database Name:", AutoSize = true };
            txtDBName = new TextBox { Text = "VenioNSRL", Size = new Size(200, 23) };
            dbNamePanel.Controls.Add(lblDBName);
            dbNamePanel.Controls.Add(txtDBName);
            dbOptionsPanel.Controls.Add(dbNamePanel);
            chkDefaultLocation = new CheckBox { Text = "Use Default database location", Checked = true, AutoSize = true };
            dbOptionsPanel.Controls.Add(chkDefaultLocation);
            mainPanel.Controls.Add(dbOptionsPanel, 0, 1);

            // File locations panel
            FlowLayoutPanel fileLocationsPanel = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown };
            var mdfPanel = new FlowLayoutPanel { AutoSize = true };
            var lblMdf = new Label { Text = "Data File (.mdf) Location", AutoSize = true };
            txtMdf = new TextBox { Size = new Size(300, 23), Enabled = false };
            btnBrowseMdf = new Button { Text = "...", Size = new Size(30, 23), Enabled = false, FlatStyle = FlatStyle.Flat };
            mdfPanel.Controls.Add(lblMdf);
            mdfPanel.Controls.Add(txtMdf);
            mdfPanel.Controls.Add(btnBrowseMdf);
            fileLocationsPanel.Controls.Add(mdfPanel);
            var ldfPanel = new FlowLayoutPanel { AutoSize = true };
            var lblLdf = new Label { Text = "Log File (.ldf) Location", AutoSize = true };
            txtLdf = new TextBox { Size = new Size(300, 23), Enabled = false };
            btnBrowseLdf = new Button { Text = "...", Size = new Size(30, 23), Enabled = false, FlatStyle = FlatStyle.Flat };
            ldfPanel.Controls.Add(lblLdf);
            ldfPanel.Controls.Add(txtLdf);
            ldfPanel.Controls.Add(btnBrowseLdf);
            fileLocationsPanel.Controls.Add(ldfPanel);
            mainPanel.Controls.Add(fileLocationsPanel, 0, 2);

            // SQL Authentication panel
            Panel authPanel = new Panel { BorderStyle = BorderStyle.FixedSingle, BackColor = Color.WhiteSmoke, AutoSize = true };
            var authLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 4, AutoSize = true };
            authPanel.Controls.Add(authLayout);
            var lblAuthTitle = new Label { Text = "SQL server authentication", AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) };
            authLayout.Controls.Add(lblAuthTitle, 0, 0);
            authLayout.SetColumnSpan(lblAuthTitle, 2);
            var lblServer = new Label { Text = "SQL database server name or instance", AutoSize = true };
            txtServer = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            authLayout.Controls.Add(lblServer, 0, 1);
            authLayout.Controls.Add(txtServer, 1, 1);
            var lblUser = new Label { Text = "Username", AutoSize = true };
            txtUsername = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            authLayout.Controls.Add(lblUser, 0, 2);
            authLayout.Controls.Add(txtUsername, 1, 2);
            var lblPass = new Label { Text = "Password", AutoSize = true };
            txtPassword = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, PasswordChar = '*' };
            authLayout.Controls.Add(lblPass, 0, 3);
            authLayout.Controls.Add(txtPassword, 1, 3);
            btnTestConnection = new Button { Text = "Test Connection", Size = new Size(180, 35), FlatStyle = FlatStyle.Flat };
            authLayout.Controls.Add(btnTestConnection, 1, 4);
            mainPanel.Controls.Add(authPanel, 0, 3);

            // Status + Log box
            FlowLayoutPanel statusPanel = new FlowLayoutPanel { AutoSize = true };
            lblStatus = new Label { Text = "Unavailable before 'Test Connection' Passes.", AutoSize = true, ForeColor = Color.Red, Font = new Font("Segoe UI", 10) };
            statusPanel.Controls.Add(lblStatus);
            mainPanel.Controls.Add(statusPanel, 0, 4);
            txtLog = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, ReadOnly = true, Font = new Font("Consolas", 9) };
            mainPanel.Controls.Add(txtLog, 0, 4);

            // Bottom buttons
            FlowLayoutPanel bottomButtons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
            btnClose = new Button { Text = "Close", Size = new Size(100, 35), FlatStyle = FlatStyle.Flat };
            btnNext = new Button { Text = "Next", Size = new Size(100, 35), Enabled = false, FlatStyle = FlatStyle.Flat };
            bottomButtons.Controls.Add(btnClose);
            bottomButtons.Controls.Add(btnNext);
            mainPanel.Controls.Add(bottomButtons, 0, 5);

            // Wire events
            btnTestConnection.Click += async (s, e) => await TestConnection();
            btnNext.Click += async (s, e) => await RunNISTImport();
            btnClose.Click += (s, e) => Close();
        }

        private void LoadValuesFromIni()
        {
            txtServer.Text = IniHelper.GetIni(iniPath, "DATABASE", "DatabaseServerName");
            txtUsername.Text = IniHelper.GetIni(iniPath, "DATABASE", "username");
            string encrypted = IniHelper.GetIni(iniPath, "DATABASE", "password");
            sqlPassword = IniHelper.DecryptPassword(encrypted, txtLog);
            if (string.IsNullOrEmpty(sqlPassword))
            {
                sqlPassword = PromptForPassword();
            }
            if (!string.IsNullOrEmpty(sqlPassword)) txtPassword.Text = "********";
        }

        private async Task TestConnection()
        {
            lblStatus.Text = "Testing connection...";
            lblStatus.ForeColor = Color.Blue;
            string connStr = DatabaseHelper.BuildConnectionString(txtServer.Text, "master", txtUsername.Text, sqlPassword);
            try
            {
                using var conn = new Microsoft.Data.SqlClient.SqlConnection(connStr);
                await conn.OpenAsync();
                lblStatus.Text = "\u2705 Connection successful";
                lblStatus.ForeColor = Color.Green;
                btnNext.Enabled = true;
            }
            catch (Exception ex)
            {
                lblStatus.Text = "\u274c Connection failed";
                lblStatus.ForeColor = Color.Red;
                txtLog.AppendText($"Connection error: {ex.Message}\n");
            }
        }

        private async Task RunNISTImport()
        {
            lblStatus.Text = "Checking NSRL database...";
            lblStatus.ForeColor = Color.Blue;
            txtLog.AppendText("Starting NIST check + import...\n");
            string dbName = chkDefaultDB.Checked ? "VenioNSRL" : txtDBName.Text;
            string masterConnStr = DatabaseHelper.BuildConnectionString(txtServer.Text, "master", txtUsername.Text, sqlPassword);
            string connStr = DatabaseHelper.BuildConnectionString(txtServer.Text, dbName, txtUsername.Text, sqlPassword);

            // Step 1: Check if DB exists, create if not
            bool dbExists = await DatabaseHelper.DatabaseExists(masterConnStr, dbName);
            if (!dbExists)
            {
                txtLog.AppendText("Creating database...\n");
                await DatabaseHelper.CreateDatabase(masterConnStr, dbName,
                    chkDefaultLocation.Checked ? null : txtMdf.Text,
                    chkDefaultLocation.Checked ? null : txtLdf.Text, txtLog);
                txtLog.AppendText("Database created.\n");
            }

            // Ensure tables exist
            await DatabaseHelper.CreateTables(connStr, txtLog);

            // Step 2: Get current version from DB
            string? currentVersion = await DatabaseHelper.GetCurrentNSRLVersion(connStr);

            // Step 3: Get latest version from NIST
            string latestVersion = await NetworkHelper.GetLatestNSRLVersion(httpClient, txtLog);
            txtLog.AppendText($"Current version: {currentVersion ?? "None"}\nLatest version: {latestVersion}\n");

            // Step 4: Show second form if update available or first time
            var progressForm = new ProgressForm(connStr, currentVersion, latestVersion, httpClient, txtLog);
            progressForm.ShowDialog();
            lblStatus.Text = "Ready";
            lblStatus.ForeColor = Color.Green;
        }

        private string PromptForPassword()
        {
            var form = new Form
            {
                Text = "Enter SQL Password",
                Size = new Size(400, 150),
                StartPosition = FormStartPosition.CenterParent,
                Font = new Font("Segoe UI", 9F)
            };
            var tb = new TextBox { Location = new Point(20, 20), Size = new Size(340, 25), PasswordChar = '*' };
            var btn = new Button { Text = "OK", Location = new Point(140, 60), FlatStyle = FlatStyle.Flat };
            form.Controls.Add(tb);
            form.Controls.Add(btn);
            btn.Click += (s, e) => form.DialogResult = DialogResult.OK;
            form.ShowDialog();
            return tb.Text;
        }
    }
}
