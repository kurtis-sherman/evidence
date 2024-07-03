using System;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace evidence
{
    public partial class Form1 : Form
    {
        private const int WM_HOTKEY = 0x0312; // F8
        private const int HOTKEY_ID = 1;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private WordUtilities wordUtils;


        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                HandleGlobalHotkey();
            }
        }

        private void HandleGlobalHotkey()
        {
            // Hide the form
            this.Hide();

            // Determine the selected radio button
            WordUtilities.AnnotationType annoType = WordUtilities.AnnotationType.None;
            if (radioFail.Checked)
                annoType = WordUtilities.AnnotationType.Fail;
            else if (radioPass.Checked)
                annoType = WordUtilities.AnnotationType.Pass;
            else if (radioInfo.Checked)
                annoType = WordUtilities.AnnotationType.Info;

            // Take screenshot and annotate
            wordUtils.AppendScreenshotToWord(annoType, textBox1.Text);

            if (!this.Visible)
                this.Show();
        }

        public Form1()
        {
            InitializeComponent();

            // Read hotkey configuration from app.config
            int hotkeyModifiers = int.Parse(ConfigurationManager.AppSettings["HotkeyModifiers"]);
            int hotkeyCode = int.Parse(ConfigurationManager.AppSettings["HotkeyCode"]);

            // Register hotkey using configuration
            RegisterHotKey(this.Handle, HOTKEY_ID, (uint)hotkeyModifiers, (uint)hotkeyCode);

            // Read monitor number from app.config
            int monitorNumber;
            try
            {
                monitorNumber = int.Parse(ConfigurationManager.AppSettings["MonitorNumber"]);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading monitor number from app.config: {ex.Message}");
                monitorNumber = 0; // default to primary monitor
            }

            // Ensure the monitor number is within valid range
            if (monitorNumber >= 0 && monitorNumber < Screen.AllScreens.Length)
            {
                Screen selectedScreen = Screen.AllScreens[monitorNumber];

                // Calculate the bottom right corner of the selected monitor's working area
                int invisibleWindowTopHeight = 40;
                int formWidth = this.Width;
                int formHeight = this.Height - invisibleWindowTopHeight;

                // Calculate position to anchor bottom right within working area
                int xPosition = selectedScreen.WorkingArea.Right - formWidth;
                int yPosition = selectedScreen.WorkingArea.Bottom - formHeight;

                // Ensure the form's position is within the working area
                xPosition = Math.Max(selectedScreen.WorkingArea.Left, xPosition);
                yPosition = Math.Max(selectedScreen.WorkingArea.Top, yPosition);

                // Set the form's location
                this.Location = new Point(xPosition, yPosition);
            }
            else
            {
                MessageBox.Show("Invalid monitor number specified in app.config. Defaulting to primary monitor.");
                // Position the form at the default location (e.g., primary monitor center)
                this.StartPosition = FormStartPosition.CenterScreen; // or adjust as per default behavior
            }

            // Other initialization code
            this.TopMost = true;

            // Create Word document
            wordUtils = new WordUtilities();
            
            LoadPowerShellScripts();
        }

        private void LoadPowerShellScripts()
        {
            string exeDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string[] scriptFiles = Directory.GetFiles(Path.Combine(exeDirectory, "scripts"), "*.ps1");

            foreach (string scriptFile in scriptFiles)
            {
                comboScripts.Items.Add(Path.GetFileName(scriptFile));
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            base.OnFormClosed(e);
        }


        private void comboScripts_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Hide the form
            this.Hide();

            string selectedScriptName = comboScripts.SelectedItem as string;
            if (selectedScriptName != null)
            {
                string scriptPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "scripts", selectedScriptName);

                // Execute PowerShell script and get its output
                string scriptOutput = ExecutePowerShellScript(scriptPath);

                // Add script output to Word document using WordUtilities
                wordUtils.AddTextToWordDocument($"Script:{selectedScriptName}");

                // Insert horizontal line
                wordUtils.AddHorizontalLineToWordDocument();

                // Add script output to Word document using WordUtilities
                wordUtils.AddTextToWordDocument(scriptOutput);
            }

            if (!this.Visible)
                this.Show();
        }

        private string ExecutePowerShellScript(string scriptPath)
        {
            // PowerShell command or script to execute
            string command = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"";

            // Create a new ProcessStartInfo
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = command,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            // Create and start the process
            using (Process process = new Process { StartInfo = psi })
            {
                process.Start();
                process.WaitForExit();

                // Read the output and errors (if needed)
                string output = process.StandardOutput.ReadToEnd();
                string errors = process.StandardError.ReadToEnd();

                if (!string.IsNullOrEmpty(errors))
                {
                    Console.WriteLine($"Error occurred while executing PowerShell script: {errors}");
                    // Handle error case if needed
                }

                return output;
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            wordUtils.SaveAndCloseWordDocument();
        }
    }
}
