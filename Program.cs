using System;
using System.Drawing;
using System.Drawing.Drawing2D;
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

        // New control panel and aspect selector
        private Panel controlPanel;
        private ComboBox aspectComboBox;

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

            // Control panel at very top (spans full width)
            controlPanel = new Panel
            {
                // Will be placed in the layout's first row
                BackColor = Color.Gainsboro,
                Padding = new Padding(5),
                Dock = DockStyle.Fill
            };

            // Create a small flow panel to hold the aspect combo and the Flip button inline
            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            aspectComboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 80,
                Margin = new Padding(0, 6, 8, 6)
            };
            aspectComboBox.Items.AddRange(new object[] { "16:9", "9:16", "Square" });
            aspectComboBox.SelectedIndex = 0; // default to 16:9
            aspectComboBox.SelectedIndexChanged += (s, e) =>
            {
                var selected = aspectComboBox.SelectedItem as string;
                if (!string.IsNullOrEmpty(selected))
                {
                    UpdateStatus($"Aspect set: {selected}");
                    // redraw overlay when aspect changes
                    pictureBox.Invalidate();
                }
            };

            // Flip button to the right of the aspect combo
            var flipButton = new Button
            {
                Text = "Flip",
                AutoSize = true,
                Margin = new Padding(0, 4, 8, 4)
            };
            flipButton.Click += (s, e) =>
            {
                try
                {
                    if (pictureBox.Image != null)
                    {
                        // Flip horizontally
                        pictureBox.Image.RotateFlip(RotateFlipType.RotateNoneFlipX);
                        pictureBox.Refresh();
                        UpdateStatus("Image flipped horizontally");
                    }
                    else
                    {
                        UpdateStatus("No image to flip");
                    }
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Error flipping image: {ex.Message}");
                }
            };

            // Rotate button to the right of the Flip button
            var rotateButton = new Button
            {
                Text = "Rotate",
                AutoSize = true,
                Margin = new Padding(0, 4, 8, 4)
            };
            rotateButton.Click += (s, e) =>
            {
                try
                {
                    if (pictureBox.Image != null)
                    {
                        // Rotate 90 degrees clockwise
                        pictureBox.Image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        pictureBox.Refresh();
                        UpdateStatus("Image rotated 90° clockwise");
                    }
                    else
                    {
                        UpdateStatus("No image to rotate");
                    }
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Error rotating image: {ex.Message}");
                }
            };

            // Add controls into the flow panel then into the control panel
            flow.Controls.Add(aspectComboBox);
            flow.Controls.Add(flipButton);
            flow.Controls.Add(rotateButton);
            controlPanel.Controls.Add(flow);

            // Status label (will be in bottom row)
            statusLabel = new Label
            {
                // Placed in layout's last row
                Text = "",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Padding = new Padding(10, 5, 10, 5),
                BackColor = Color.LightBlue,
                Dock = DockStyle.Fill
            };

            // Panel that contains the PictureBox (middle, fills available space)
            imagePanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.Black
            };

            pictureBox = new PictureBox
            {
                Location = new Point(0, 0),
                // Will be stretched in the layout cell
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom
            };

            // Paint overlay and handle resize
            pictureBox.Paint += PictureBox_Paint;
            pictureBox.Resize += (s, e) => pictureBox.Invalidate();

            imagePanel.Controls.Add(pictureBox);

            // Panel for storage folder selection (third row)
            Panel storageFolderPanel = new Panel
            {
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(5),
                Dock = DockStyle.Fill
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
            // Fixed row height so we can vertically center controls
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));

            Label storageLabel = new Label
            {
                Text = "Storage folder:",
                AutoSize = false,
                // Let the cell sizing and Dock control vertical alignment
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 5, 0)
            };

            storageFolderTextBox = new TextBox
            {
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 3, 5, 3),
                Anchor = AnchorStyles.Left | AnchorStyles.Right
            };

            browseButton = new Button
            {
                Text = "Browse...",
                AutoSize = false,
                Dock = DockStyle.Fill,
                Height = 24,
                Margin = new Padding(0, 3, 0, 3),
                FlatStyle = FlatStyle.Standard
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

            storageFolderPanel.Controls.Add(layout);

            // Main layout to stack panels vertically: controlPanel, imagePanel, storageFolderPanel, statusLabel
            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            // Row styles: controlPanel fixed, imagePanel fills, storageFolderPanel fixed, statusLabel fixed
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // controlPanel
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // imagePanel fills remaining
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // storageFolderPanel
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F)); // statusLabel

            mainLayout.Controls.Add(controlPanel, 0, 0);
            mainLayout.Controls.Add(imagePanel, 0, 1);
            mainLayout.Controls.Add(storageFolderPanel, 0, 2);
            mainLayout.Controls.Add(statusLabel, 0, 3);

            this.Controls.Add(mainLayout);

            this.FormClosing += (s, e) => isRunning = false;
        }

        private void PictureBox_Paint(object sender, PaintEventArgs e)
        {
            try
            {
                double ratio = ParseAspectRatio(aspectComboBox.SelectedItem as string);
                if (ratio <= 0)
                    return;

                int w = pictureBox.ClientSize.Width;
                int h = pictureBox.ClientSize.Height;
                if (w <= 0 || h <= 0)
                    return;

                // Determine maximum rectangle matching ratio that fits within w x h
                double controlRatio = (double)w / h;
                int rectW, rectH;
                if (controlRatio > ratio)
                {
                    // height limits
                    rectH = h;
                    rectW = (int)Math.Round(h * ratio);
                }
                else
                {
                    // width limits
                    rectW = w;
                    rectH = (int)Math.Round(w / ratio);
                }

                int x = (w - rectW) / 2;
                int y = (h - rectH) / 2;

                using (var pen = new Pen(Color.Yellow, 2))
                {
                    pen.Alignment = PenAlignment.Center;
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    e.Graphics.DrawRectangle(pen, x, y, rectW - 1, rectH - 1);
                }
            }
            catch
            {
                // swallow any painting exceptions
            }
        }

        private double ParseAspectRatio(string s)
        {
            if (string.IsNullOrEmpty(s)) return 16.0 / 9.0;
            if (s.Equals("Square", StringComparison.OrdinalIgnoreCase)) return 1.0;
            var parts = s.Split(':');
            if (parts.Length == 2 && double.TryParse(parts[0], out double a) && double.TryParse(parts[1], out double b) && b != 0)
            {
                return a / b;
            }
            return 16.0 / 9.0;
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

                // redraw overlay when a new image is loaded
                pictureBox.Invalidate();
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
