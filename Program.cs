using System;
using System.Drawing;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PhotoFileViewer
{
    public class PhotoViewerForm : Form
    {
        private Panel imagePanel;
        private PictureBox pictureBox;
        private Label statusLabel;
        private bool isRunning = true;
        private const string PIPE_NAME = "PhotoFileViewerPipe";

        // Added controls for storage folder selection
        private TextBox storageFolderTextBox;
        private Button browseButton;
        private string storageFolderPath;

        public PhotoViewerForm(string initialFile)
        {
            InitializeComponent();

            if (!string.IsNullOrEmpty(initialFile))
            {
                OpenFile(initialFile);
            }
            else
            {
                UpdateStatus("Ready. Waiting for Photo files...");
            }

            // Start listening for files from other instances
            Task.Run(() => ListenForFiles());
        }

        private void InitializeComponent()
        {
            this.Text = "Photo File Viewer";
            this.Size = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Status label at top
            statusLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Padding = new Padding(10, 5, 10, 5),
                BackColor = Color.LightBlue
            };

            // Panel that contains the PictureBox
            imagePanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.Black
            };

            pictureBox = new PictureBox
            {
                Location = new Point(0, 0),
                // Fill the available area and scale the image to fit while preserving aspect ratio
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom
            };

            imagePanel.Controls.Add(pictureBox);

            // Panel for storage folder selection (bottom)
            Panel instructionsPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 80,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(10)
            };

            // Use a TableLayoutPanel so the Browse button stays visible regardless of window width
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                AutoSize = false
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F)); // label width
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); // textbox fills remaining space
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // button auto size

            Label storageLabel = new Label
            {
                Text = "Storage folder:",
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 3, 5, 3)
            };

            storageFolderTextBox = new TextBox
            {
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 3, 5, 3)
            };

            browseButton = new Button
            {
                Text = "Browse...",
                AutoSize = true,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0, 0, 0, 0)
            };
            browseButton.Click += (s, e) =>
            {
                using (var dlg = new FolderBrowserDialog())
                {
                    dlg.Description = "Select storage folder";
                    dlg.ShowNewFolderButton = true;
                    if (!string.IsNullOrEmpty(storageFolderPath) && Directory.Exists(storageFolderPath))
                    {
                        try { dlg.SelectedPath = storageFolderPath; } catch { }
                    }

                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        storageFolderPath = dlg.SelectedPath;
                        storageFolderTextBox.Text = storageFolderPath;
                        UpdateStatus($"Storage folder set: {storageFolderPath}");
                    }
                }
            };

            layout.Controls.Add(storageLabel, 0, 0);
            layout.Controls.Add(storageFolderTextBox, 1, 0);
            layout.Controls.Add(browseButton, 2, 0);

            instructionsPanel.Controls.Add(layout);

            this.Controls.Add(imagePanel);
            this.Controls.Add(statusLabel);
            this.Controls.Add(instructionsPanel);

            this.FormClosing += (s, e) => isRunning = false;
        }

        public void OpenFile(string filePath)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(OpenFile), filePath);
                return;
            }

            try
            {
                if (!File.Exists(filePath))
                {
                    UpdateStatus($"Error: File not found - {filePath}");
                    MessageBox.Show(
                        $"File not found:\n{filePath}",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // Only attempt to load common image extensions; still try to load other image types too
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".bmp" && ext != ".gif")
                {
                    // Still try to load but warn user in status
                    UpdateStatus($"Opening file (unsupported extension): {Path.GetFileName(filePath)}");
                }

                // Dispose previous image to avoid memory leak and file locks
                if (pictureBox.Image != null)
                {
                    var old = pictureBox.Image;
                    pictureBox.Image = null;
                    old.Dispose();
                }

                // Read file into memory then create image from memory to avoid locking the file
                byte[] bytes = File.ReadAllBytes(filePath);
                using (var ms = new MemoryStream(bytes))
                {
                    using (var img = Image.FromStream(ms))
                    {
                        // Create a copy (Bitmap) so it does not depend on the stream
                        pictureBox.Image = new Bitmap(img);
                    }
                }

                // With Dock = Fill and SizeMode = Zoom, the image will be scaled to fit the available window area.
                string fileName = Path.GetFileName(filePath);
                string fileSize = FormatFileSize(new FileInfo(filePath).Length);
                string timestamp = DateTime.Now.ToString("HH:mm:ss");

                int imgW = pictureBox.Image?.Width ?? 0;
                int imgH = pictureBox.Image?.Height ?? 0;

                UpdateStatus($"[{timestamp}] Opened: {fileName} ({fileSize}) - {imgW}x{imgH} (fitted to window)");

                // Bring window to front
                BringToFront();
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error opening file: {ex.Message}");
                MessageBox.Show(
                    $"Error opening file:\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void UpdateStatus(string status)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(UpdateStatus), status);
                return;
            }

            statusLabel.Text = status;
        }

        private void BringToFront()
        {
            if (WindowState == FormWindowState.Minimized)
            {
                WindowState = FormWindowState.Normal;
            }

            TopMost = true;
            TopMost = false;
            Activate();
            Focus();
        }

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double size = bytes;
            int order = 0;

            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }

            return $"{size:0.##} {sizes[order]}";
        }

        private async Task ListenForFiles()
        {
            while (isRunning)
            {
                try
                {
                    using (var server = new NamedPipeServerStream(
                        PIPE_NAME,
                        PipeDirection.In,
                        NamedPipeServerStream.MaxAllowedServerInstances,
                        PipeTransmissionMode.Message,
                        PipeOptions.Asynchronous))
                    {
                        await server.WaitForConnectionAsync();

                        using (var reader = new StreamReader(server, Encoding.UTF8))
                        {
                            string filePath = await reader.ReadToEndAsync();
                            if (!string.IsNullOrEmpty(filePath))
                            {
                                OpenFile(filePath);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (isRunning)
                    {
                        Console.WriteLine($"Error in pipe server: {ex.Message}");
                    }
                }
            }
        }

        public static bool SendFileToExistingInstance(string filePath)
        {
            try
            {
                using (var client = new NamedPipeClientStream(
                    ".",
                    PIPE_NAME,
                    PipeDirection.Out,
                    PipeOptions.None))
                {
                    client.Connect(2000); // 2 second timeout

                    using (var writer = new StreamWriter(client, Encoding.UTF8))
                    {
                        writer.AutoFlush = true;
                        writer.Write(filePath);
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    static class Program
    {
        private const string MUTEX_NAME = "Global\\PhotoFileViewerMutex";

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Get file path from command line
            string filePath = args.Length > 0 ? args[0] : string.Empty;

            bool createdNew;
            using (var mutex = new Mutex(true, MUTEX_NAME, out createdNew))
            {
                if (!createdNew)
                {
                    // Another instance is already running
                    if (string.IsNullOrEmpty(filePath))
                    {
                        MessageBox.Show(
                            "Photo File Viewer is already running.",
                            "Already Running",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return;
                    }

                    if (PhotoViewerForm.SendFileToExistingInstance(filePath))
                    {
                        // File sent successfully to existing instance
                        return;
                    }
                    else
                    {
                        MessageBox.Show(
                            "Could not connect to existing instance.\nThe file was not opened.",
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return;
                    }
                }

                // This is the first instance
                try
                {
                    Application.Run(new PhotoViewerForm(filePath));
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
        }
    }
}
