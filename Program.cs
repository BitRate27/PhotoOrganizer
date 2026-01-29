using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Imaging;
using Newtonsoft.Json;

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

        // Overlay rectangle so we can reference it when saving
        private Rectangle overlayRectangle;
        private Image originalImage;
        private Image fullImage;
        // Store original image metadata (EXIF) property items
        private PropertyItem[] originalPropertyItems;

        // Center point in fullImage to clip around (in source image coordinates)
        private Point fullImageClipCenter;

        // Zoom factor increases/decreases by0.1 each time zoom is pressed
        private double zoomFactor =1.0;
        // Settings file path for persisting user preferences (JSON)
        private readonly string settingsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PhotoOrganizer", "settings.json");

        // Dragging state for clip-center panning
        private bool isDraggingClip = false;
        private Point dragStartMouse;
        private Point dragStartCenter;

        public PhotoViewerForm(string initialFile)
        {
            InitializeComponent();
            // Load persisted settings after UI is initialized
            LoadUserSettings();

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

        double maxZoomFactor(Image image, int w, int h)
        {
            // Compute zoomFactor so that the full original image fits into the picture box
            double fx = image.Width / (double)w;
            double fy = image.Height / (double)h;
            double target = Math.Max(fx, fy);
            return target;
        }

        private Rectangle justifyRectangleInImage(Image source, Point center, int w, int h)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            // Clamp rect to image bounds

            Rectangle srcArea = new Rectangle(center.X - w / 2, center.Y - h / 2, w, h);

            if (srcArea.IsEmpty || srcArea.Width <= 0 || srcArea.Height <= 0 ||
                srcArea.Width > source.Width || srcArea.Height > source.Height)
            {
                return Rectangle.Empty;
            }
            
            if (srcArea.X < 0) srcArea.X = 0;
            if (srcArea.Y < 0) srcArea.Y = 0;

            if (srcArea.X + srcArea.Width > source.Width) srcArea.X = source.Width - srcArea.Width;
            if (srcArea.Y + srcArea.Height > source.Height) srcArea.Y = source.Height - srcArea.Height;

            Rectangle imgBounds = new Rectangle(0, 0, source.Width, source.Height);
            Rectangle justified = Rectangle.Intersect(imgBounds, srcArea);

            if (justified.IsEmpty || justified.Width <= 0 || justified.Height <= 0)
            {
                return Rectangle.Empty;
            }
            return justified;

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
                    overlayRectangle = computeOverlayRectangle(selected, pictureBox.ClientSize.Width, pictureBox.ClientSize.Height);
                    updatePictureBoxImage();
                    pictureBox.Refresh();

                    // persist aspect selection
                    SaveUserSettings();
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
                        if (originalImage != null)
                        {
                            originalImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
                        }
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
                        if (originalImage != null)
                        {
                            originalImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        }

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

            // Zoom Out button to the right of Rotate button
            var zoomOutButton = new Button
            {
                Text = "Zoom Out",
                AutoSize = true,
                Margin = new Padding(0,4,8,4)
            };
            zoomOutButton.Click += (s, e) =>
            {
                try
                {
                    // Increase zoomFactor by 0.1 each time
                    if ((zoomFactor + 0.1) < maxZoomFactor(originalImage, overlayRectangle.Width, overlayRectangle.Height))
                    {
                        zoomFactor += 0.1;
                        UpdateStatus($"Zoom factor: {zoomFactor:0.0}");
                    }
                    else
                    {
                        UpdateStatus($"Maximum Zoom factor: {zoomFactor:0.0}");
                    }    

                    // Refresh the displayed image in case zoomFactor is used elsewhere
                    updatePictureBoxImage();
                    pictureBox.Refresh();
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Error adjusting zoom: {ex.Message}");
                }
            };

            // Zoom In button (decrease zoomFactor)
            var zoomInButton = new Button
            {
                Text = "Zoom In",
                AutoSize = true,
                Margin = new Padding(0,4,8,4)
            };
            zoomInButton.Click += (s, e) =>
            {
                try
                {
                    // Decrease zoomFactor by0.1 each time, but do not go below0.1
                    zoomFactor = Math.Max(0.1, zoomFactor -0.1);
                    UpdateStatus($"Zoom factor: {zoomFactor:0.0}");
                    updatePictureBoxImage();
                    pictureBox.Refresh();
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Error adjusting zoom: {ex.Message}");
                }
            };

            // Show All button to fit the original image into the picture box
            var showAllButton = new Button
            {
                Text = "Show All",
                AutoSize = true,
                Margin = new Padding(0,4,8,4)
            };
            showAllButton.Click += (s, e) =>
            {
                try
                {
                    if (originalImage == null || pictureBox == null) {
                        UpdateStatus("No image to fit");
                        return;
                    }

                    // Compute zoomFactor so that the original image fits into the picture box
                    zoomFactor = maxZoomFactor(originalImage, overlayRectangle.Width, overlayRectangle.Height);

                    // Center the clip on the fullImage so the original (which is centered in fullImage) is shown
                    if (fullImage != null)
                    {
                        fullImageClipCenter = new Point(fullImage.Width /2, fullImage.Height /2);
                    }

                    updatePictureBoxImage();
                    pictureBox.Refresh();
                    UpdateStatus($"Show All — zoomFactor set: {zoomFactor:0.00}");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Error fitting image: {ex.Message}");
                }
            };

            // Add controls into the flow panel then into the control panel
            flow.Controls.Add(aspectComboBox);
            flow.Controls.Add(flipButton);
            flow.Controls.Add(rotateButton);
            flow.Controls.Add(zoomInButton);
            flow.Controls.Add(zoomOutButton);
            flow.Controls.Add(showAllButton);
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

            // When the form is resized, recompute the overlay rectangle and regenerate the image
            this.Resize += (s, e) =>
            {
                // Recompute overlay rectangle using current aspect selection and new pictureBox size
                overlayRectangle = computeOverlayRectangle(aspectComboBox.SelectedItem as string, pictureBox.ClientSize.Width, pictureBox.ClientSize.Height);

                if (originalImage != null)
                {
                    // Compute zoomFactor limit based on new overlay size so we don't exceed source bounds
                    double max = maxZoomFactor(originalImage, overlayRectangle.Width > 0 ? overlayRectangle.Width : pictureBox.ClientSize.Width,
                    overlayRectangle.Height > 0 ? overlayRectangle.Height : pictureBox.ClientSize.Height);

                    // Ensure a sensible zoomFactor
                    zoomFactor = Math.Min(zoomFactor, max);
                }

                updatePictureBoxImage();
                pictureBox.Refresh();
            };

            // Allow click+drag on the picture to pan the clip-center
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
                        // persist storage folder
                        SaveUserSettings();
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
                    var imageRect = new Rectangle(offsetX, offsetY, dispW, dispH);
                    var intersect = Rectangle.Intersect(overlayRectangle, imageRect);
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
                        g.DrawImage(srcImg, new Rectangle(0,0, srcW, srcH), new Rectangle(srcX, srcY, srcW, srcH), GraphicsUnit.Pixel);

                        // Copy original metadata into the saved bitmap if available
                        if (originalPropertyItems != null)
                        {
                            foreach (var prop in originalPropertyItems)
                            {
                                if (prop == null) continue;
                                try
                                {
                                    // Some property items may not be valid for the new image; ignore failures
                                    bmp.SetPropertyItem(prop);
                                }
                                catch { }
                            }
                        }

                        // Add ImageEditingSoftware (ID=42043) and ImageEditor (ID=42040)
                        try
                        {
                            // Helper to create a PropertyItem instance without public ctor
                            Func<int, short, byte[], PropertyItem> makeProp = (id, type, val) =>
                            {
                                var pi = (PropertyItem)System.Runtime.Serialization.FormatterServices.GetUninitializedObject(typeof(PropertyItem));
                                pi.Id = id;
                                pi.Type = type;
                                pi.Len = val?.Length ?? 0;
                                pi.Value = val;
                                return pi;
                            };

                            string software = "PhotoOrganizer (BitRate27)";
                            byte[] softwareBytes = Encoding.ASCII.GetBytes(software + "\0");
                            var softwareProp = makeProp(42043, 2, softwareBytes); // Type2 = ASCII
                            try { bmp.SetPropertyItem(softwareProp); } catch { }

                            string editor = Environment.UserName ?? string.Empty;
                            byte[] editorBytes = Encoding.ASCII.GetBytes(editor + "\0");
                            var editorProp = makeProp(42040, 2, editorBytes);
                            try { bmp.SetPropertyItem(editorProp); } catch { }
                        }
                        catch
                        {
                            // ignore any failures creating/setting custom metadata
                        }

                        string fileName = "image_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".jpg";
                        if (!string.IsNullOrEmpty(currentFilePath))
                        {
                            fileName = Path.GetFileName(currentFilePath);
                        }

                        string destPath = Path.Combine(storageFolderPath, fileName);
                        // Choose JPEG encoder to save; overwrite if exists
                        bmp.Save(destPath, System.Drawing.Imaging.ImageFormat.Jpeg);

                   
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

            // Compute initial overlay rectangle now that controls and sizes are set
            overlayRectangle = computeOverlayRectangle(aspectComboBox.SelectedItem as string, pictureBox.ClientSize.Width, pictureBox.ClientSize.Height);
            // Ensure the image is rendered with the initial overlay
            updatePictureBoxImage();
            pictureBox.Refresh();

            this.SizeChanged += (s, e) => { SaveUserSettings(); };
            this.FormClosing += (s, e) => { SaveUserSettings(); isRunning = false; };
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
                    // Re-adjust zoomFactor to account for new pictureBox size
                    // Compute zoomFactor so that the original image fits into the picture box
                    double max = maxZoomFactor(originalImage, overlayRectangle.Width, overlayRectangle.Height);

                    // Ensure a sensible zoomFactor
                    zoomFactor = Math.Min(zoomFactor, max);

                    // Source rectangle size: try to match pictureBox size, but clamp to fullImage bounds
                    Rectangle srcArea = justifyRectangleInImage(fullImage, fullImageClipCenter,
                        (int)Math.Round((double)pbw * zoomFactor), 
                        (int)Math.Round((double)pbh * zoomFactor));
                    Rectangle dstArea = new Rectangle(0, 0, pbw, pbh);

                    pictureBox.Image = ZoomImage(fullImage, srcArea, dstArea);
                    //pictureBox.Image = ClipImage(fullImage, srcArea);
                }
            }
        }
        private Rectangle computeOverlayRectangle(string aspectRatio, int w, int h)
        {
            const int overlayPadding = 15;      
            
            double ratio = ParseAspectRatio(aspectRatio);
            if (ratio <= 0)
                return Rectangle.Empty;

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
            return new Rectangle(x, y, rectW, rectH);
        }

        private void PictureBox_Paint(object sender, PaintEventArgs e)
        {
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

                int x = overlayRectangle.X;
                int y = overlayRectangle.Y;
                int rectW = overlayRectangle.Width;
                int rectH = overlayRectangle.Height;

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

            int newCenterX = dragStartCenter.X - (int)Math.Round((double)deltaSrcX * zoomFactor);
            int newCenterY = dragStartCenter.Y - (int)Math.Round((double)deltaSrcY * zoomFactor);

            // Clamp to image bounds
            newCenterX = Math.Max(pbw/2, Math.Min(newCenterX, fullImage.Width - 1 - pbw/2));
            newCenterY = Math.Max(pbh/2, Math.Min(newCenterY, fullImage.Height - 1 - pbh/2));

            fullImageClipCenter = new Point(newCenterX, newCenterY);
            updatePictureBoxImage();
            pictureBox.Refresh();
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

        // Zoom the image around the given center. Takes and area of the source (srcArea) and make returns a new image
        // the size of the destination area (dstArea).
        private Image ZoomImage(Image source, Rectangle srcArea, Rectangle dstArea)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            if ((dstArea.Width/dstArea.Height) != (srcArea.Width/srcArea.Height))
            {
                // Aspect ratios do not match; cannot scale properly
                return null;
            }

            var bmp = new Bitmap(dstArea.Width, dstArea.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.DrawImage(source, dstArea, srcArea, GraphicsUnit.Pixel);
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

                        if (fullImage != null)
                        {
                            var old = fullImage;
                            fullImage = null;
                            old.Dispose();
                        }

                        if (originalImage != null)
                        {
                            var old = originalImage;
                            originalImage = null;
                            old.Dispose();
                        }

                        // Create a copy (Bitmap) so it does not depend on the stream
                        originalImage = new Bitmap(img);

                        // Capture metadata (PropertyItems) from the original image, if any
                        try
                        {
                            var ids = img.PropertyIdList;
                            if (ids != null && ids.Length >0)
                            {
                                originalPropertyItems = new PropertyItem[ids.Length];
                                for (int i =0; i < ids.Length; i++)
                                {
                                    try { originalPropertyItems[i] = img.GetPropertyItem(ids[i]); } catch { originalPropertyItems[i] = null; }
                                }
                            }
                            else
                            {
                                originalPropertyItems = null;
                            }
                        }
                        catch
                        {
                            originalPropertyItems = null;
                        }

                        // Create a new fullImage twice as large in each dimension, filled with black,
                        // and draw the original image centered inside it.
                        int max = Math.Max(originalImage.Width, originalImage.Height);
                        int fullW = max * 3;
                        int fullH = max * 3;
                        var big = new Bitmap(fullW, fullH);
                        using (var g = Graphics.FromImage(big))
                        {
                            g.Clear(Color.Black);
                            int offsetX = (fullW - originalImage.Width) / 2;
                            int offsetY = (fullH - originalImage.Height) / 2;
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.DrawImage(originalImage, offsetX, offsetY, originalImage.Width, originalImage.Height);
                        }

                        fullImageClipCenter = new Point(fullW / 2, fullH / 2);

                        // Set fullImage to the constructed large bitmap and dispose the temporary original
                        fullImage = big;
 
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
                pictureBox.Refresh();
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

        private void LoadUserSettings()
        {
            try
            {
                if (!File.Exists(settingsFilePath)) return;
                var json = File.ReadAllText(settingsFilePath);
                var settings = JsonConvert.DeserializeObject<UserSettings>(json);
                if (settings == null) return;

                if (!string.IsNullOrEmpty(settings.StorageFolderPath) && Directory.Exists(settings.StorageFolderPath))
                {
                    storageFolderPath = settings.StorageFolderPath;
                    try { storageFolderTextBox.Text = storageFolderPath; } catch { }
                }

                if (!string.IsNullOrEmpty(settings.Aspect))
                {
                    for (int i =0; i < aspectComboBox.Items.Count; i++)
                    {
                        if (string.Equals(aspectComboBox.Items[i].ToString(), settings.Aspect, StringComparison.OrdinalIgnoreCase))
                        {
                            aspectComboBox.SelectedIndex = i;
                            break;
                        }
                    }
                }

                if (settings.FormWidth >100 && settings.FormHeight >100)
                {
                    this.Size = new Size(settings.FormWidth, settings.FormHeight);
                }
            }
            catch
            {
                // ignore load errors
            }
        }

        private void SaveUserSettings()
        {
            try
            {
                var dir = Path.GetDirectoryName(settingsFilePath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                var settings = new UserSettings
                {
                    StorageFolderPath = storageFolderPath ?? string.Empty,
                    Aspect = aspectComboBox.SelectedItem as string ?? string.Empty,
                    FormWidth = this.Size.Width,
                    FormHeight = this.Size.Height
                };
                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(settingsFilePath, json);
            }
            catch
            {
                // ignore save errors
            }
        }

        private class UserSettings
        {
            public string StorageFolderPath { get; set; }
            public string Aspect { get; set; }
            public int FormWidth { get; set; }
            public int FormHeight { get; set; }
        }
        
        private void LoadUserSettingsLegacy()
        {
            try
            {
                if (!File.Exists(settingsFilePath)) return;
                var json = File.ReadAllText(settingsFilePath);
                string folder = ReadJsonString(json, "StorageFolderPath");
                if (!string.IsNullOrEmpty(folder) && Directory.Exists(folder))
                {
                    storageFolderPath = folder;
                    try { storageFolderTextBox.Text = storageFolderPath; } catch { }
                }

                string aspect = ReadJsonString(json, "Aspect");
                if (!string.IsNullOrEmpty(aspect))
                {
                    for (int i =0; i < aspectComboBox.Items.Count; i++)
                    {
                        if (string.Equals(aspectComboBox.Items[i].ToString(), aspect, StringComparison.OrdinalIgnoreCase))
                        {
                            aspectComboBox.SelectedIndex = i;
                            break;
                        }
                    }
                }

                int fw = ReadJsonInt(json, "FormWidth",0);
                int fh = ReadJsonInt(json, "FormHeight",0);
                if (fw >100 && fh >100)
                {
                    this.Size = new Size(fw, fh);
                }
            }
            catch
            {
                // ignore load errors
            }
        }

        private void SaveUserSettingsLegacy()
        {
            try
            {
                var dir = Path.GetDirectoryName(settingsFilePath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                string aspect = aspectComboBox.SelectedItem as string ?? string.Empty;
                string json = "{" +
                "\"StorageFolderPath\":" + JsonString(storageFolderPath) + "," +
                "\"Aspect\":" + JsonString(aspect) + "," +
                "\"FormWidth\":" + this.Size.Width + "," +
                "\"FormHeight\":" + this.Size.Height +
                "}";
                File.WriteAllText(settingsFilePath, json);
            }
            catch
            {
                // ignore save errors
            }
        }
        
        private static string JsonString(string s)
        {
            if (s == null) return "\"\"";
            var esc = s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n");
            return "\"" + esc + "\"";
        }
        
        private static string ReadJsonString(string json, string key)
        {
            if (string.IsNullOrEmpty(json) || string.IsNullOrEmpty(key)) return string.Empty;
            var token = "\"" + key + "\"";
            int idx = json.IndexOf(token, StringComparison.OrdinalIgnoreCase);
            if (idx <0) return string.Empty;
            int colon = json.IndexOf(':', idx + token.Length);
            if (colon <0) return string.Empty;
            int i = colon +1;
            while (i < json.Length && char.IsWhiteSpace(json[i])) i++;
            if (i >= json.Length) return string.Empty;
            if (json[i] != '"') return string.Empty;
            i++; // skip opening quote
            var sb = new StringBuilder();
            bool escape = false;
            for (; i < json.Length; i++)
            {
                char c = json[i];
                if (escape)
                {
                    if (c == '"') sb.Append('"');
                    else if (c == '\\') sb.Append('\\');
                    else if (c == 'n') sb.Append('\n');
                    else if (c == 'r') sb.Append('\r');
                    else sb.Append(c);
                    escape = false;
                }
                else
                {
                    if (c == '\\') { escape = true; continue; }
                    if (c == '"') break;
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }
        
        private static int ReadJsonInt(string json, string key, int defaultValue)
        {
            if (string.IsNullOrEmpty(json) || string.IsNullOrEmpty(key)) return defaultValue;
            var token = "\"" + key + "\"";
            int idx = json.IndexOf(token, StringComparison.OrdinalIgnoreCase);
            if (idx <0) return defaultValue;
            int colon = json.IndexOf(':', idx + token.Length);
            if (colon <0) return defaultValue;
            int i = colon +1;
            while (i < json.Length && (char.IsWhiteSpace(json[i]) || json[i]==',')) i++;
            int start = i;
            while (i < json.Length && (char.IsDigit(json[i]) || json[i]=='-')) i++;
            if (i > start)
            {
                if (int.TryParse(json.Substring(start, i - start), out int v)) return v;
            }
            return defaultValue;
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
