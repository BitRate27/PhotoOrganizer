using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.IO.Pipes;
using System.Runtime.Remoting;
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

        // Keep track of currently opened file path
        private string currentFilePath;

        // Store overlay rectangle so it can be accessed later
        private int overlayX;
        private int overlayY;
        private int overlayWidth;
        private int overlayHeight;
        private Image fullImage;
        // Center point in fullImage to clip around (in source image coordinates)
        private Point fullImageClipCenter;
        // Dragging state for clip-center panning
        private bool isDraggingClip = false;
        private Point dragStartMouse;
        private Point dragStartCenter;

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
                    if (fullImage != null)
                    {
                        // Flip horizontally
                        fullImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
                        fullImageClipCenter = new Point(fullImage.Width - fullImageClipCenter.X, fullImageClipCenter.Y);
                        updatePictureBoxImage();
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
                    if (fullImage != null)
                    {
                        // Rotate 90 degrees clockwise
                        // Save old center and dimensions before rotation
                        int oldCenterX = fullImageClipCenter.X;
                        int oldCenterY = fullImageClipCenter.Y;
                        int oldWidth = fullImage.Width;
                        int oldHeight = fullImage.Height;

                        fullImage.RotateFlip(RotateFlipType.Rotate90FlipNone);

                        // After rotating 90° clockwise, a point at (x,y) in the old image
                        // maps to (y, oldWidth -1 - x) in the rotated image.
                        int newCenterX = oldCenterY;
                        int newCenterY = oldWidth - 1 - oldCenterX;

                        // Clamp to new image bounds
                        newCenterX = Math.Max(0, Math.Min(newCenterX, fullImage.Width - 1));
                        newCenterY = Math.Max(0, Math.Min(newCenterY, fullImage.Height - 1));

                        fullImageClipCenter = new Point(newCenterX, newCenterY);

                        updatePictureBoxImage();
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

            // When the picture box size changes, regenerate the clipped image so it matches new size
            pictureBox.Resize += (s, e) => { updatePictureBoxImage(); pictureBox.Invalidate(); };

            // Allow click+drag on the picture to pan the clip center
            pictureBox.MouseDown += PictureBox_MouseDown;
            pictureBox.MouseMove += PictureBox_MouseMove;
            pictureBox.MouseUp += PictureBox_MouseUp;
            pictureBox.MouseLeave += (s, e) => { if (isDraggingClip) { isDraggingClip = false; Cursor = Cursors.Default; } };

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
                ColumnCount = 4,
                RowCount = 1,
                AutoSize = false
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F)); // label width
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); // textbox fills remaining space
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // browse button
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // save button
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

            // Save button to the right of Browse
            var saveButton = new Button
            {
                Text = "Save",
                AutoSize = false,
                Dock = DockStyle.Fill,
                Height = 24,
                Margin = new Padding(6, 3, 0, 3),
                FlatStyle = FlatStyle.Standard
            };
            saveButton.Click += (s, e) =>
            {
                try
                {
                    if (string.IsNullOrEmpty(storageFolderPath) || !Directory.Exists(storageFolderPath))
                    {
                        MessageBox.Show(this, "Please select a storage folder first.", "No Storage Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    if (pictureBox.Image == null)
                    {
                        UpdateStatus("No image to save");
                        return;
                    }

                    // Determine mapping from PictureBox client coords to image pixels
                    var srcImg = pictureBox.Image as Image;
                    int imgW = srcImg.Width;
                    int imgH = srcImg.Height;

                    int pbW = pictureBox.ClientSize.Width;
                    int pbH = pictureBox.ClientSize.Height;

                    double scale = Math.Min((double)pbW / imgW, (double)pbH / imgH);
                    int dispW = (int)Math.Round(imgW * scale);
                    int dispH = (int)Math.Round(imgH * scale);
                    int offsetX = (pbW - dispW) / 2;
                    int offsetY = (pbH - dispH) / 2;

                    // Intersection of overlay rectangle with displayed image area
                    var overlayRect = new Rectangle(overlayX, overlayY, overlayWidth, overlayHeight);
                    var imageRect = new Rectangle(offsetX, offsetY, dispW, dispH);
                    var intersect = Rectangle.Intersect(overlayRect, imageRect);
                    if (intersect.IsEmpty)
                    {
                        MessageBox.Show(this, "Selected overlay does not intersect the image.", "Nothing to save", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Map to source image pixels
                    double invScale = 1.0 / scale;
                    int srcX = (int)Math.Floor((intersect.X - offsetX) * invScale);
                    int srcY = (int)Math.Floor((intersect.Y - offsetY) * invScale);
                    int srcW = (int)Math.Ceiling(intersect.Width * invScale);
                    int srcH = (int)Math.Ceiling(intersect.Height * invScale);

                    // Clamp to image bounds
                    srcX = Math.Max(0, Math.Min(srcX, imgW - 1));
                    srcY = Math.Max(0, Math.Min(srcY, imgH - 1));
                    if (srcX + srcW > imgW) srcW = imgW - srcX;
                    if (srcY + srcH > imgH) srcH = imgH - srcY;
                    if (srcW <= 0 || srcH <= 0)
                    {
                        MessageBox.Show(this, "Computed crop is empty.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Crop and save
                    using (var bmp = new Bitmap(srcW, srcH))
                    using (var g = Graphics.FromImage(bmp))
                    {
                        g.DrawImage(srcImg, new Rectangle(0, 0, srcW, srcH), new Rectangle(srcX, srcY, srcW, srcH), GraphicsUnit.Pixel);

                        string fileName = "image_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".jpg";
                        if (!string.IsNullOrEmpty(currentFilePath))
                        {
                            fileName = Path.GetFileName(currentFilePath);
                        }

                        string destPath = Path.Combine(storageFolderPath, fileName);
                        // Choose JPEG encoder to save; overwrite if exists
                        bmp.Save(destPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                        UpdateStatus($"Saved cropped image to: {destPath}");
                    }
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Error saving image: {ex.Message}");
                    MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            layout.Controls.Add(storageLabel, 0, 0);
            layout.Controls.Add(storageFolderTextBox, 1, 0);
            layout.Controls.Add(browseButton, 2, 0);
            layout.Controls.Add(saveButton, 3, 0);

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
        private void updatePictureBoxImage()
        {
            if (pictureBox.Image != null)
            {
                var old = pictureBox.Image;
                pictureBox.Image = null;
                old.Dispose();
            }
            if (fullImage != null)
            {
                int pbw = pictureBox.ClientSize.Width;
                int pbh = pictureBox.ClientSize.Height;
                if (pbw > 0 && pbh > 0)
                {
                    // Source rectangle size: try to match pictureBox size, but clamp to fullImage bounds
                    int srcW = Math.Min(fullImage.Width, pbw);
                    int srcH = Math.Min(fullImage.Height, pbh);

                    int srcX = fullImageClipCenter.X - srcW / 2;
                    int srcY = fullImageClipCenter.Y - srcH / 2;

                    // Clamp (shouldn't be necessary but safe)
                    if (srcX < 0) srcX = 0;
                    if (srcY < 0) srcY = 0;
                    if (srcX + srcW > fullImage.Width) srcW = fullImage.Width - srcX;
                    if (srcY + srcH > fullImage.Height) srcH = fullImage.Height - srcY;

                    if (pictureBox.Image != null)
                    {
                        var old = pictureBox.Image;
                        pictureBox.Image = null;
                        old.Dispose();
                    }

                    pictureBox.Image = ClipImage(fullImage, new Rectangle(srcX, srcY, srcW, srcH));
                }
            }
        }

        private void PictureBox_Paint(object sender, PaintEventArgs e)
        {
            const int overlayPadding = 15;
            try
            {
                int w = pictureBox.ClientSize.Width;
                int h = pictureBox.ClientSize.Height; 
                
                if (w <= 0 || h <= 0)
                    return;

                if (pictureBox.Image != null)
                {
                    e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    e.Graphics.CompositingQuality = CompositingQuality.HighQuality;

                    e.Graphics.DrawImage(pictureBox.Image, new Rectangle(0, 0, w, h));
                }

                double ratio = ParseAspectRatio(aspectComboBox.SelectedItem as string);
                if (ratio <= 0)
                    return;

                // Determine maximum rectangle matching ratio that fits within w x h
                double controlRatio = (double)w / h;
                int rectW, rectH;
                if (controlRatio > ratio)
                {
                    // height limits
                    rectH = h - (overlayPadding * 2);
                    rectW = (int)Math.Round((h - (overlayPadding * 2)) * ratio);
                }
                else
                {
                    // width limits
                    rectW = w - (overlayPadding * 2);
                    rectH = (int)Math.Round((w - (overlayPadding * 2)) / ratio);
                }

                int x = (w - rectW) / 2;
                int y = (h - rectH) / 2;

                // store overlay coordinates for future use
                overlayX = x;
                overlayY = y;
                overlayWidth = rectW;
                overlayHeight = rectH;

                // Draw a semi-transparent border by filling the areas outside the inner rectangle.
                // The inner rectangle remains fully transparent so the image shows through.
                var overlayColor = Color.FromArgb(140,0,0,0); // semi-transparent black
                using (var brush = new SolidBrush(overlayColor))
                {
                    // Top
                    e.Graphics.FillRectangle(brush,0,0, w, y);
                    // Bottom
                    e.Graphics.FillRectangle(brush,0, y + rectH, w, h - (y + rectH));
                    // Left
                    e.Graphics.FillRectangle(brush,0, y, x, rectH);
                    // Right
                    e.Graphics.FillRectangle(brush, x + rectW, y, w - (x + rectW), rectH);
                }
                // Optionally draw a thin border line to delineate the inner rectangle
                using (var pen = new Pen(Color.FromArgb(200,255,255,255),1))
                {
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    e.Graphics.DrawRectangle(pen, x, y, rectW -1, rectH -1);
                }
            }
            catch
            {
                // swallow any painting exceptions
            }
        }

        private void PictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            if (fullImage == null) return;

            isDraggingClip = true;
            dragStartMouse = e.Location;
            dragStartCenter = fullImageClipCenter;
            Cursor = Cursors.Hand;
            pictureBox.Capture = true;
        }

        private void PictureBox_MouseMove(object sender, MouseEventArgs e)
        {
            if (!isDraggingClip || fullImage == null) return;

            int pbw = pictureBox.ClientSize.Width;
            int pbh = pictureBox.ClientSize.Height;
            if (pbw <=0 || pbh <=0) return;

            // Source rectangle size used in Paint
            int srcW = Math.Min(fullImage.Width, pbw);
            int srcH = Math.Min(fullImage.Height, pbh);

            int dx = e.Location.X - dragStartMouse.X;
            int dy = e.Location.Y - dragStartMouse.Y;

            // Map display delta to source delta and update center (subtract so content follows cursor)
            int deltaSrcX = (int)Math.Round(dx * (double)srcW / pbw);
            int deltaSrcY = (int)Math.Round(dy * (double)srcH / pbh);

            int newCenterX = dragStartCenter.X - deltaSrcX;
            int newCenterY = dragStartCenter.Y - deltaSrcY;

            // Clamp to image bounds
            newCenterX = Math.Max(pbw/2, Math.Min(newCenterX, fullImage.Width - 1 - pbw/2));
            newCenterY = Math.Max(pbh/2, Math.Min(newCenterY, fullImage.Height - 1 - pbh/2));

            fullImageClipCenter = new Point(newCenterX, newCenterY);
            updatePictureBoxImage();
            pictureBox.Invalidate();
        }

        private void PictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            isDraggingClip = false;
            Cursor = Cursors.Default;
            pictureBox.Capture = false;
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

        // Clip the given image to the provided rectangle (inclusive). Returns a new Image (Bitmap).
        private Image ClipImage(Image source, Rectangle rect)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            // Clamp rect to source bounds
            Rectangle srcBounds = new Rectangle(0, 0, source.Width, source.Height);
            Rectangle clip = Rectangle.Intersect(srcBounds, rect);

            if (clip.IsEmpty || clip.Width <= 0 || clip.Height <= 0)
            {
                return null;
            }

            var bmp = new Bitmap(clip.Width, clip.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;
                // Draw the requested region of the source into the new bitmap
                g.DrawImage(source, new Rectangle(0, 0, clip.Width, clip.Height), clip, GraphicsUnit.Pixel);
            }

            return bmp;
        }

        // Zoom the image around the given center. Takes a source image, a center point (in source coordinates),
        // an input region size (inXSize x inYSize) to sample around the center, and a zoomFactor.
        // Produces a new Image whose size is (inXSize * zoomFactor, inYSize * zoomFactor) containing
        // the sampled region scaled to the output size.
        private Image ZoomImage(Image source, Point center, int inXSize, int inYSize, double zoomFactor)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (inXSize <= 0 || inYSize <= 0) throw new ArgumentException("inXSize and inYSize must be >0");
            if (zoomFactor <= 0) throw new ArgumentException("zoomFactor must be >0");

            // Determine source rectangle centered on `center` with requested input size.
            int srcW = Math.Min(inXSize, source.Width);
            int srcH = Math.Min(inYSize, source.Height);

            int srcX = center.X - srcW / 2;
            int srcY = center.Y - srcH / 2;

            // Clamp to source bounds
            if (srcX < 0) srcX = 0;
            if (srcY < 0) srcY = 0;
            if (srcX + srcW > source.Width) srcX = Math.Max(0, source.Width - srcW);
            if (srcY + srcH > source.Height) srcY = Math.Max(0, source.Height - srcH);

            int destW = Math.Max(1, (int)Math.Round(inXSize * zoomFactor));
            int destH = Math.Max(1, (int)Math.Round(inYSize * zoomFactor));

            var bmp = new Bitmap(destW, destH);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.DrawImage(source, new Rectangle(0, 0, destW, destH), new Rectangle(srcX, srcY, srcW, srcH), GraphicsUnit.Pixel);
            }

            return bmp;
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

                // Read file into memory then create image from memory to avoid locking the file
                byte[] bytes = File.ReadAllBytes(filePath);
                using (var ms = new MemoryStream(bytes))
                {
                    using (var img = Image.FromStream(ms))
                    {
                        // Create a copy (Bitmap) so it does not depend on the stream
                        if (fullImage != null)
                        {
                            var old = fullImage;
                            fullImage = null;
                            old.Dispose();
                        }
                        fullImage = new Bitmap(img);

                        // Default clip center is image center
                        fullImageClipCenter = new Point(fullImage.Width /2, fullImage.Height /2);

                        updatePictureBoxImage();
                    }
                }

                // With Dock = Fill and SizeMode = Zoom, the image will be scaled to fit the available window area.
                string fileName = Path.GetFileName(filePath);
                currentFilePath = filePath;
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
