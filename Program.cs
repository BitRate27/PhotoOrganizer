using ImageMagick;
using Newtonsoft.Json;
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Pipes;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using static System.Windows.Forms.MonthCalendar;

namespace PhotoFileViewer
{
    public class PhotoViewerForm : Form
    {
        private Panel imagePanel;
        private PictureBox pictureBox;
        private Label statusLabel;
        private Label photoInfo;
        private bool isRunning = true;
        private const string PIPE_NAME = "PhotoFileViewerPipe";

        // Added controls for storage folder selection
        private TextBox storageFolderTextBox;
        private Button browseButton;
        private string storageFolderPath;

        // New control panel and aspect selector
        private Panel controlPanel;
        private ComboBox aspectComboBox;
        private ComboBox resolutionComboBox;
        // Rotation slider (degrees) -20..20, default0
        private double rotationAngle = 0.0;
        // Show grid while user is dragging rotation slider
        private bool showRotationGrid = false;

        // Keep track of currently opened file path
        private string currentFilePath;

        // Overlay rectangle so we can reference it when saving
        private Rectangle overlayRectangle;
        private Image originalImage;
        private Image fullImage;
        // Store original image metadata (EXIF) property items
        private PropertyItem[] originalPropertyItems;
        private bool gpsLocationAvailable = false;
        private TextBox locationInputTextBox;

        // Center point in fullImage to clip around (in source image coordinates)
        private Point fullImageClipCenter;

        // Zoom factor increases/decreases by0.1 each time zoom is pressed
        private double zoomFactor = 1.0;
        // Save resolution option (High/Low)
        private string saveResolution = "High";
        // Settings file path for persisting user preferences (JSON)
        private readonly string settingsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PhotoOrganizer", "settings.json");

        // Dragging state for clip-center panning
        private bool isDraggingClip = false;
        private Point dragStartMouse;
        private Point dragStartCenter;
        private const int EXIFOrientationID = 274;

        // Right-button dragging for rotation
        private bool isRightDragging = false;
        private Point rightDragStart;
        private Point rightDragCurrent;

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

        double maxZoomFactor(bool fill, int imageW, int imageH, int w, int h)
        {
            // Compute zoomFactor so that the full original image fits into the picture box
            double fx = imageW / (double)w;
            double fy = imageH / (double)h;
            double target = fill ? Math.Min(fx, fy) : Math.Max(fx, fy);
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
            this.Text = "Photo Organizer";
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
            aspectComboBox.Items.AddRange(new object[] { "16:9", "4:5", "Square" });
            aspectComboBox.SelectedIndex = 0; // default to16:9
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

            // Resolution selector placed next to aspect selector
            resolutionComboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 80,
                Margin = new Padding(0, 6, 8, 6)
            };
            resolutionComboBox.Items.AddRange(new object[] { "High", "Low", "Original" });
            resolutionComboBox.SelectedIndex = 0; // default High
            resolutionComboBox.SelectedIndexChanged += (s, e) =>
            {
                var sel = resolutionComboBox.SelectedItem as string;
                if (!string.IsNullOrEmpty(sel))
                {
                    saveResolution = sel;
                    updatePictureBoxImage();
                    pictureBox.Refresh();
                }
            };

            // Toggle button to show/hide the rotation grid manually
            var gridToggle = new CheckBox
            {
                Text = "Grid",
                Appearance = Appearance.Button,
                AutoSize = true,
                Margin = new Padding(0, 6, 8, 6),
                Checked = false
            };
            gridToggle.CheckedChanged += (s, e) =>
            {
                showRotationGrid = gridToggle.Checked;
                pictureBox?.Refresh();
            };

            // Reset button only
            var rotateResetButton = new Button
            {
                Text = "Reset",
                AutoSize = true,
                Margin = new Padding(0, 6, 4, 6)
            };

            // Helper to apply rotation around original image center and rebuild fullImage
            Action applyRotation = () =>
            {
                try
                {
                    if (originalImage == null) return;
                    double angle = rotationAngle; // degrees
                    int max = Math.Max(originalImage.Width, originalImage.Height);
                    int fullW = Math.Min(16000, max * 3);
                    int fullH = Math.Min(16000, max * 3);
                    var big = new Bitmap(fullW, fullH);
                    using (var g = Graphics.FromImage(big))
                    {
                        g.Clear(Color.Black);
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.CompositingQuality = CompositingQuality.HighQuality;
                        g.TranslateTransform(fullW / 2f, fullH / 2f);
                        g.RotateTransform((float)angle);
                        g.TranslateTransform(-originalImage.Width / 2f, -originalImage.Height / 2f);
                        g.DrawImage(originalImage, 0, 0, originalImage.Width, originalImage.Height);
                        g.ResetTransform();
                    }
                    if (fullImage != null) { try { fullImage.Dispose(); } catch { } fullImage = null; }
                    fullImage = big;
                    updatePictureBoxImage();
                    pictureBox.Refresh();
                    UpdateStatus($"Image rotated {rotationAngle:0.00}°");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Error rotating image: {ex.Message}");
                }
            };

            rotateResetButton.Click += (s, e) => { rotationAngle = 0.0; applyRotation(); };

            // Flip button
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

            // Show All button
            var showAllButton = new Button
            {
                Text = "Show All",
                AutoSize = true,
                Margin = new Padding(0, 4, 8, 4)
            };
            showAllButton.Click += (s, e) =>
            {
                try
                {
                    if (originalImage == null || pictureBox == null)
                    {
                        UpdateStatus("No image to fit");
                        return;
                    }

                    zoomFactor = maxZoomFactor(false, originalImage.Width, originalImage.Height, overlayRectangle.Width, overlayRectangle.Height);
                    if (fullImage != null)
                    {
                        fullImageClipCenter = new Point(fullImage.Width / 2, fullImage.Height / 2);
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

            // Fill button
            var fillButton = new Button
            {
                Text = "Fill",
                AutoSize = true,
                Margin = new Padding(0, 4, 8, 4)
            };
            fillButton.Click += (s, e) =>
            {
                try
                {
                    if (originalImage == null || pictureBox == null)
                    {
                        UpdateStatus("No image to fit");
                        return;
                    }

                    zoomFactor = maxZoomFactor(true, originalImage.Width, originalImage.Height, overlayRectangle.Width, overlayRectangle.Height);
                    if (fullImage != null)
                    {
                        fullImageClipCenter = new Point(fullImage.Width / 2, fullImage.Height / 2);
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
            flow.Controls.Add(resolutionComboBox);
            flow.Controls.Add(gridToggle);
            flow.Controls.Add(rotateResetButton);
            flow.Controls.Add(flipButton);
            flow.Controls.Add(showAllButton);
            flow.Controls.Add(fillButton);
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

            // Photo info label (shows original image resolution)
            photoInfo = new Label
            {
                Text = "",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI",10, FontStyle.Regular),
                Padding = new Padding(10,5,10,5),
                BackColor = Color.LightBlue,
                Dock = DockStyle.Left,
                AutoSize = true
            };

            // TextBox for user to paste lat,long when no GPS EXIF is available. Initially hidden.
            locationInputTextBox = new TextBox
            {
                Visible = false,
                Width =300,
                Dock = DockStyle.Right,
                Margin = new Padding(10,3,10,3)
            };
            locationInputTextBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    try { HandleLocationInput(locationInputTextBox.Text); } catch { }
                }
            };

            // Panel to contain photoInfo label and optional location input textbox
            var photoInfoPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.LightBlue
            };
            photoInfoPanel.Controls.Add(photoInfo);
            photoInfoPanel.Controls.Add(locationInputTextBox);

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
                SizeMode = PictureBoxSizeMode.Zoom,
                TabStop = true
            };

            // Paint overlay and handle resize
            pictureBox.Paint += PictureBox_Paint;

            // Ensure pictureBox can receive wheel events by giving it focus when the mouse enters
            pictureBox.MouseEnter += (s, e) => { try { pictureBox.Focus(); } catch { } };

            // Mouse wheel zoom (centered on cursor) — use multiplicative steps so larger zooms change more
            pictureBox.MouseWheel += (s, e) =>
            {
                try
                {
                    if (fullImage == null) return;

                    int pbW = pictureBox.ClientSize.Width;
                    int pbH = pictureBox.ClientSize.Height;
                    if (pbW <= 0 || pbH <= 0) return;

                    // Source area before zoom
                    int srcW1 = (int)Math.Round((double)pbW * zoomFactor);
                    int srcH1 = (int)Math.Round((double)pbH * zoomFactor);
                    var srcArea1 = justifyRectangleInImage(fullImage, fullImageClipCenter, srcW1, srcH1);
                    if (srcArea1.IsEmpty) return;

                    // Mouse position in pictureBox
                    var mouseX = e.X;
                    var mouseY = e.Y;

                    // Full-image point under mouse before zoom
                    double px = srcArea1.X + (mouseX / (double)pbW) * srcArea1.Width;
                    double py = srcArea1.Y + (mouseY / (double)pbH) * srcArea1.Height;

                    // Multiplicative step factor —10% per wheel step (can be tuned)
                    // If Control is pressed, use a finer step
                    double stepFactor = (ModifierKeys & Keys.Control) == Keys.Control ? 1.025 : 1.1;

                    if (e.Delta > 0)
                    {
                        double max = maxZoomFactor(false, originalImage.Width, originalImage.Height,
                        overlayRectangle.Width > 0 ? overlayRectangle.Width : pbW,
                        overlayRectangle.Height > 0 ? overlayRectangle.Height : pbH);
                        zoomFactor = Math.Min(max, zoomFactor * stepFactor);
                    }
                    else if (e.Delta < 0)
                    {
                        zoomFactor = Math.Max(0.1, zoomFactor / stepFactor);
                    }

                    // New source size after zoom
                    int srcW2 = (int)Math.Round((double)pbW * zoomFactor);
                    int srcH2 = (int)Math.Round((double)pbH * zoomFactor);

                    // Desired source origin so that px,py maps to same mouse position
                    double desiredSrcX = px - (mouseX / (double)pbW) * srcW2;
                    double desiredSrcY = py - (mouseY / (double)pbH) * srcH2;

                    double desiredCenterX = desiredSrcX + srcW2 / 2.0;
                    double desiredCenterY = desiredSrcY + srcH2 / 2.0;

                    var tentativeCenter = new Point((int)Math.Round(desiredCenterX), (int)Math.Round(desiredCenterY));
                    var srcArea2 = justifyRectangleInImage(fullImage, tentativeCenter, srcW2, srcH2);
                    if (!srcArea2.IsEmpty)
                    {
                        fullImageClipCenter = new Point(srcArea2.X + srcArea2.Width / 2, srcArea2.Y + srcArea2.Height / 2);
                    }

                    UpdateStatus($"Zoom factor: {zoomFactor:0.00}");
                    updatePictureBoxImage();
                    pictureBox.Refresh();
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Error adjusting zoom: {ex.Message}");
                }
            };

            // When the form is resized, recompute the overlay rectangle and regenerate the image
            this.Resize += (s, e) =>
            {
                // Recompute overlay rectangle using current aspect selection and new pictureBox size
                overlayRectangle = computeOverlayRectangle(aspectComboBox.SelectedItem as string, pictureBox.ClientSize.Width, pictureBox.ClientSize.Height);

                if (originalImage != null)
                {
                    // Compute zoomFactor limit based on new overlay size so we don't exceed source bounds
                    double max = maxZoomFactor(false, originalImage.Width, originalImage.Height, overlayRectangle.Width > 0 ? overlayRectangle.Width : pictureBox.ClientSize.Width,
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

                    Rectangle srcRect = GetOverlayRectangleOnFullImage(overlayRectangle, fullImageClipCenter, zoomFactor);
                    Rectangle dstRect = new Rectangle(0, 0, srcRect.Width, srcRect.Height);

                    // If user did not selected Original, save at a fixed resolution depending on aspect and resolution
                    if (!string.Equals(saveResolution, "Original", StringComparison.OrdinalIgnoreCase))
                    {
                        dstRect = GetFixedResolutionRect(aspectComboBox.SelectedItem as string,
                            saveResolution);
                    }

                    Image saveImage = ZoomImage(fullImage, srcRect, dstRect);

                    using (var bmp = new Bitmap(dstRect.Width, dstRect.Height))
                    using (var g = Graphics.FromImage(bmp))
                    {
                        g.DrawImage(saveImage, dstRect, dstRect, GraphicsUnit.Pixel);

                        // Copy original metadata into the saved bitmap if available
                        if (originalPropertyItems != null)
                        {
                            foreach (var prop in originalPropertyItems)
                            {
                                if (prop == null) continue;
                                try
                                {
                                    // Some property items may not be valid for the new image; ignore failures
                                    if (prop.Id == EXIFOrientationID)
                                    {
                                        // EXIF Orientation tag: set to1 (Horizontal)
                                        try
                                        {
                                            var newProp = (PropertyItem)System.Runtime.Serialization.FormatterServices.GetUninitializedObject(typeof(PropertyItem));
                                            newProp.Id = 274;
                                            newProp.Type = 3; // SHORT
                                            newProp.Len = 2;
                                            newProp.Value = new byte[] { 1, 0 }; // little-endian1
                                            bmp.SetPropertyItem(newProp);
                                        }
                                        catch { }
                                    }
                                    else
                                    {
                                        try { bmp.SetPropertyItem(prop); } catch { }
                                    }
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

            // Main layout
            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // controlPanel
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // imagePanel fills remaining
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // storageFolderPanel
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F)); // photoInfo
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F)); // statusLabel

            mainLayout.Controls.Add(controlPanel, 0, 0);
            mainLayout.Controls.Add(imagePanel, 0, 1);
            mainLayout.Controls.Add(storageFolderPanel, 0, 2);
            mainLayout.Controls.Add(photoInfoPanel,0,3);
            mainLayout.Controls.Add(statusLabel, 0, 4);

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
                    double max = maxZoomFactor(false, originalImage.Width, originalImage.Height, overlayRectangle.Width, overlayRectangle.Height);

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

        // Returns a Rectangle with width and height chosen from fixed presets
        // based on the aspect ratio string and resolution string.
        // Aspect may be e.g. "16:9", "9:16", "Square", "4:5" (case-insensitive).
        // Resolution is "High" or "Low" (case-insensitive).
        private Rectangle GetFixedResolutionRect(string aspect, string resolution)
        {
            if (string.IsNullOrEmpty(aspect) || string.IsNullOrEmpty(resolution))
                return Rectangle.Empty;

            string a = aspect.Trim();
            string r = resolution.Trim();

            bool high = r.Equals("High", StringComparison.OrdinalIgnoreCase);

            int w = 0, h = 0;

            if (a.Equals("16:9", StringComparison.OrdinalIgnoreCase))
            {
                if (high) { w = 3840; h = 2160; }
                else { w = 1920; h = 1080; }
            }
            else if (a.Equals("9:16", StringComparison.OrdinalIgnoreCase))
            {
                // portrait version of16:9
                if (high) { w = 2160; h = 3840; }
                else { w = 1080; h = 1920; }
            }
            else if (a.Equals("Square", StringComparison.OrdinalIgnoreCase))
            {
                if (high) { w = 2048; h = 2048; }
                else { w = 1080; h = 1080; }
            }
            else if (a.Equals("4:5", StringComparison.OrdinalIgnoreCase))
            {
                // portrait4:5 (width smaller)
                if (high) { w = 2160; h = 2700; }
                else { w = 1080; h = 1350; }
            }
            else if (a.Equals("5:4", StringComparison.OrdinalIgnoreCase))
            {
                // landscape5:4 (swap4:5)
                if (high) { w = 2700; h = 2160; }
                else { w = 1350; h = 1080; }
            }
            else
            {
                return Rectangle.Empty;
            }

            return new Rectangle(0, 0, w, h);
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

                Rectangle srcRect = GetOverlayRectangleOnFullImage(overlayRectangle, fullImageClipCenter, zoomFactor);

                bool diminishedQuality = false;
                if (!string.Equals(saveResolution, "Original", StringComparison.OrdinalIgnoreCase))
                {
                    Rectangle dstRect = GetFixedResolutionRect(aspectComboBox.SelectedItem as string,
                        saveResolution);

                    // Show red border if the source rectangle is smaller than the destination rectangle
                    if (srcRect.Width < dstRect.Width || srcRect.Height < dstRect.Height)
                    {
                        diminishedQuality = true;
                    }
                }

                // Draw a semi-transparent border by filling the areas outside the inner rectangle.
                // The inner rectangle remains fully transparent so the image shows through.
                var overlayColor = Color.FromArgb(140, 0, 0, 0); // semi-transparent black
                using (var brush = new SolidBrush(overlayColor))
                {
                    // Top
                    e.Graphics.FillRectangle(brush, 0, 0, w, y);
                    // Bottom
                    e.Graphics.FillRectangle(brush, 0, y + rectH, w, h - (y + rectH));
                    // Left
                    e.Graphics.FillRectangle(brush, 0, y, x, rectH);
                    // Right
                    e.Graphics.FillRectangle(brush, x + rectW, y, w - (x + rectW), rectH);
                }
                // If requested, draw a grid inside the overlay rectangle
                if (showRotationGrid)
                {
                    var aspect = aspectComboBox.SelectedItem as string ?? "16:9";
                    int cols = 16, rows = 9;
                    if (string.Equals(aspect, "4:5", StringComparison.OrdinalIgnoreCase)) { cols = 4; rows = 5; }
                    else if (string.Equals(aspect, "Square", StringComparison.OrdinalIgnoreCase)) { cols = 5; rows = 5; }
                    else if (string.Equals(aspect, "16:9", StringComparison.OrdinalIgnoreCase)) { cols = 16; rows = 9; }

                    if (cols > 0 && rows > 0 && rectW > 0 && rectH > 0)
                    {
                        using (var pen = new Pen(Color.FromArgb(180, 255, 255, 255), 1))
                        {
                            pen.DashStyle = DashStyle.Solid;
                            float cellW = rectW / (float)cols;
                            float cellH = rectH / (float)rows;

                            // draw vertical lines
                            for (int i = 1; i < cols; i++)
                            {
                                float gx = x + i * cellW;
                                e.Graphics.DrawLine(pen, gx, y, gx, y + rectH);
                            }

                            // draw horizontal lines
                            for (int j = 1; j < rows; j++)
                            {
                                float gy = y + j * cellH;
                                e.Graphics.DrawLine(pen, x, gy, x + rectW, gy);
                            }
                        }
                    }
                }
                // Optionally draw a thin border line to delineate the inner rectangle
                using (var pen = new Pen(diminishedQuality ? Color.Red : Color.FromArgb(200, 255, 255, 255), 1))
                {
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    e.Graphics.DrawRectangle(pen, x, y, rectW - 1, rectH - 1);
                }

                // Draw rubberband line if right-dragging
                if (isRightDragging)
                {
                    try
                    {
                        using (var rbPen = new Pen(Color.Yellow, 2))
                        {
                            rbPen.DashStyle = DashStyle.Solid;
                            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                            e.Graphics.DrawLine(rbPen, rightDragStart, rightDragCurrent);
                        }
                    }
                    catch { }
                }
            }
            catch
            {
                // swallow any painting exceptions
            }
        }

        private void PictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (fullImage == null) return;

                isDraggingClip = true;
                dragStartMouse = e.Location;
                dragStartCenter = fullImageClipCenter;
                Cursor = Cursors.Hand;
                pictureBox.Capture = true;
                return;
            }

            if (e.Button == MouseButtons.Right)
            {
                // Start rubberband rotation
                isRightDragging = true;
                rightDragStart = e.Location;
                rightDragCurrent = e.Location;
                pictureBox.Capture = true;
                Cursor = Cursors.Cross;
                pictureBox.Invalidate();
            }
        }

        private void PictureBox_MouseMove(object sender, MouseEventArgs e)
        {
            if (isRightDragging)
            {
                rightDragCurrent = e.Location;
                // Erase previous line by invalidating (causes Paint to redraw without previous line)
                pictureBox.Invalidate();
                return;
            }

            if (!isDraggingClip || fullImage == null) return;

            int pbw = pictureBox.ClientSize.Width;
            int pbh = pictureBox.ClientSize.Height;
            if (pbw <= 0 || pbh <= 0) return;

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
            newCenterX = Math.Max(0, Math.Min(newCenterX, fullImage.Width - 1));
            newCenterY = Math.Max(0, Math.Min(newCenterY, fullImage.Height - 1));

            fullImageClipCenter = new Point(newCenterX, newCenterY);
            updatePictureBoxImage();
            pictureBox.Refresh();
        }

        private void PictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDraggingClip = false;
                Cursor = Cursors.Default;
                pictureBox.Capture = false;
                return;
            }

            if (e.Button == MouseButtons.Right && isRightDragging)
            {
                // Finalize rotation based on rubberband angle
                isRightDragging = false;
                pictureBox.Capture = false;
                Cursor = Cursors.Default;

                // Compute angle in degrees from the rubberband line
                double dx = rightDragCurrent.X - rightDragStart.X;
                double dy = rightDragCurrent.Y - rightDragStart.Y;
                if (dx == 0 && dy == 0)
                {
                    // No movement; nothing to do
                    pictureBox.Invalidate();
                    return;
                }

                double angleDeg = 0.0;

                // To provide a more intuitive rotation feel, we can choose to compute the angle
                // based on the dominant direction of the drag.
                if ((dx <= 0.0) && (dy <= 0.0)) // A
                {
                    angleDeg = Math.Atan2(-dy, -dx) * (180.0 / Math.PI);
                }
                else if ((dx >= 0.0) && (dy <= 0.0)) // B
                {
                    angleDeg = Math.Atan2(dy, dx) * (180.0 / Math.PI);
                }
                else if ((dx <= 0.0) && (dy >= 0.0)) // C
                {
                    angleDeg = Math.Atan2(-dy, -dx) * (180.0 / Math.PI);
                }
                else // D
                {
                    angleDeg = Math.Atan2(dy, dx) * (180.0 / Math.PI);
                }

                // Calculate size of an unrotated image (same aspect ratio as original) that will fit completely
                // inside the originalImage rotated by -angleDeg. We compute the maximal uniform scale factor s (<=1)
                // such that a rectangle of size s*W x s*H, when rotated by -angleDeg, fits entirely within the
                // original image bounds. This assumes the candidate rectangle is centered.
                int fitW = 0, fitH = 0;
                try
                {
                    if (originalImage != null)
                    {
                        double theta = -angleDeg * Math.PI / 180.0; // rotation in radians
                        double W0 = originalImage.Width;
                        double H0 = originalImage.Height;

                        // Use absolute values of sin/cos because symmetry
                        double c = Math.Abs(Math.Cos(theta));
                        double s = Math.Abs(Math.Sin(theta));

                        // Axis-aligned bounding box of rotated rectangle (w,h) is:
                        // bboxW = w*c + h*s
                        // bboxH = w*s + h*c
                        // For w = scale*W0, h = scale*H0 we require:
                        // scale*(W0*c + H0*s) <= W0
                        // scale*(W0*s + H0*c) <= H0
                        // Solve for scale limits
                        double denomW = W0 * c + H0 * s;
                        double denomH = W0 * s + H0 * c;

                        double scaleLimitW = denomW > 1e-9 ? (W0 / denomW) : double.MaxValue;
                        double scaleLimitH = denomH > 1e-9 ? (H0 / denomH) : double.MaxValue;

                        double scale = Math.Min(1.0, Math.Min(scaleLimitW, scaleLimitH));

                        fitW = Math.Max(1, (int)Math.Round(scale * W0));
                        fitH = Math.Max(1, (int)Math.Round(scale * H0));

                        UpdateStatus($"Max unrotated fit inside rotated image: {fitW}x{fitH} (scale {scale:0.000})");

                        // Compute zoomFactor so that the original image fits into the picture box
                        zoomFactor = maxZoomFactor(true, fitW, fitH, overlayRectangle.Width, overlayRectangle.Height);
                    }
                }
                catch { }

                // Apply rotation
                try
                {
                    // Reuse the same CPU rotate logic used elsewhere
                    rotationAngle = -angleDeg;
                    int max = Math.Max(originalImage.Width, originalImage.Height);
                    int fullW = Math.Min(16000, max * 3);
                    int fullH = Math.Min(16000, max * 3);
                    var big = new Bitmap(fullW, fullH);
                    using (var g = Graphics.FromImage(big))
                    {
                        g.Clear(Color.Black);
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.CompositingQuality = CompositingQuality.HighQuality;
                        g.TranslateTransform(fullW / 2f, fullH / 2f);
                        g.RotateTransform((float)rotationAngle);
                        g.TranslateTransform(-originalImage.Width / 2f, -originalImage.Height / 2f);
                        g.DrawImage(originalImage, 0, 0, originalImage.Width, originalImage.Height);
                        g.ResetTransform();
                    }
                    if (fullImage != null) { try { fullImage.Dispose(); } catch { } fullImage = null; }
                    fullImage = big;
                    fullImageClipCenter = new Point(fullImage.Width / 2, fullImage.Height / 2);
                    updatePictureBoxImage();
                    pictureBox.Refresh();
                    UpdateStatus($"Image rotated {rotationAngle:0.00}°");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Error rotating image: {ex.Message}");
                }

                // Erase last line by invalidating
                pictureBox.Invalidate();
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

            if (srcArea.IsEmpty || dstArea.Height <= 0 || srcArea.Height <= 0)
            {
                return null;
            }   
            if ((dstArea.Width / dstArea.Height) != (srcArea.Width / srcArea.Height))
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
        Rectangle GetOverlayRectangleOnFullImage(Rectangle overlayRect, Point fullImageCenter, double zoomFactor)
        {
            int srcW = (int)Math.Round((double)overlayRect.Width * zoomFactor);
            int srcH = (int)Math.Round((double)overlayRect.Height * zoomFactor);
            int srcX = fullImageCenter.X - (srcW /2);
            int srcY = fullImageCenter.Y - (srcH /2);
            return new Rectangle(srcX, srcY, srcW, srcH);
        }

        // Try to read GPS/location from an array of PropertyItems (EXIF). Returns "lat,lon" or empty.
        private string GetImageLocation(PropertyItem[] items)
        {
            try
            {
                if (items == null || items.Length ==0) return string.Empty;
                const int GPSLatitudeId =0x0002;
                const int GPSLongitudeId =0x0004;
                const int GPSLatitudeRefId =0x0001;
                const int GPSLongitudeRefId =0x0003;

                PropertyItem latItem = null, lonItem = null, latRef = null, lonRef = null;
                foreach (var pi in items)
                {
                    if (pi == null) continue;
                    if (pi.Id == 0x8825)
                    {
                        // GPSInfo tag that may contain a pointer to the GPS IFD; not used here since we already have all PropertyItems
                        continue;
                    }
                    if (pi.Id == GPSLatitudeId) latItem = pi;
                    else if (pi.Id == GPSLongitudeId) lonItem = pi;
                    else if (pi.Id == GPSLatitudeRefId) latRef = pi;
                    else if (pi.Id == GPSLongitudeRefId) lonRef = pi;
                }

                if (latItem == null || lonItem == null) return string.Empty;

                double ParseRationalTriplet(PropertyItem pi)
                {
                    var val = pi.Value;
                    if (val == null || val.Length <24) return 0.0;
                    double[] comps = new double[3];
                    for (int i =0; i <3; i++)
                    {
                        int offset = i *8;
                        uint num = BitConverter.ToUInt32(val, offset);
                        uint den = BitConverter.ToUInt32(val, offset +4);
                        comps[i] = den ==0 ?0.0 : (double)num / den;
                    }
                    return comps[0] + comps[1] /60.0 + comps[2] /3600.0;
                }

                double lat = ParseRationalTriplet(latItem);
                double lon = ParseRationalTriplet(lonItem);

                if (latRef != null && latRef.Value != null && latRef.Value.Length >0)
                {
                    var c = (char)latRef.Value[0];
                    if (c == 'S' || c == 's') lat = -lat;
                }
                if (lonRef != null && lonRef.Value != null && lonRef.Value.Length >0)
                {
                    var c = (char)lonRef.Value[0];
                    if (c == 'W' || c == 'w') lon = -lon;
                }

                return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.######},{1:0.######}", lat, lon);
            }
            catch
            {
                return string.Empty;
            }
        }

        // Call Google Geocode API to resolve lat/lon to a formatted address. Uses environment variable GOOGLE_GEOCODE_API_KEY.
        private async Task<string> GetImageAddressAsync(PropertyItem[] items)
        {
            try
            {
                var loc = GetImageLocation(items);
                if (string.IsNullOrEmpty(loc)) return string.Empty;

                if (string.IsNullOrEmpty(apiKey.Value)) return string.Empty;

                var url = $"https://maps.googleapis.com/maps/api/geocode/json?latlng={Uri.EscapeDataString(loc)}&key={Uri.EscapeDataString(apiKey.Value)}";

                using (var http = new HttpClient())
                {
                    http.Timeout = TimeSpan.FromSeconds(5);
                    var resp = await http.GetAsync(url).ConfigureAwait(false);
                    if (!resp.IsSuccessStatusCode) return string.Empty;
                    var body = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (string.IsNullOrEmpty(body)) return string.Empty;

                    try
                    {
                        var j = JObject.Parse(body);
                        var status = (string)j["status"];
                        if (!string.Equals(status, "OK", StringComparison.OrdinalIgnoreCase)) return string.Empty;
                        var first = j["results"]?.First;
                        if (first == null) return string.Empty;
                        var formatted = (string)first["formatted_address"];
                        return formatted ?? string.Empty;
                    }
                    catch { return string.Empty; }
                }
            }
            catch { return string.Empty; }
        }

        // Handle user-pasted lat,long input, convert to EXIF PropertyItems and update state
        private void HandleLocationInput(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;
            var parts = text.Split(new char[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length <2) return;
            if (!double.TryParse(parts[0].Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double lat)) return;
            if (!double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double lon)) return;

            try
            {
                if (originalPropertyItems == null) originalPropertyItems = new PropertyItem[4];

                var latRef = (PropertyItem)System.Runtime.Serialization.FormatterServices.GetUninitializedObject(typeof(PropertyItem));
                latRef.Id =0x0001; latRef.Type =2; latRef.Len =2; latRef.Value = Encoding.ASCII.GetBytes((lat >=0 ? "N" : "S") + "\0");

                var lonRef = (PropertyItem)System.Runtime.Serialization.FormatterServices.GetUninitializedObject(typeof(PropertyItem));
                lonRef.Id =0x0003; lonRef.Type =2; lonRef.Len =2; lonRef.Value = Encoding.ASCII.GetBytes((lon >=0 ? "E" : "W") + "\0");

                Func<double, byte[]> ToRationalTriplet = (double val) =>
                {
                    val = Math.Abs(val);
                    int deg = (int)Math.Floor(val);
                    double rem = (val - deg) *60.0;
                    int min = (int)Math.Floor(rem);
                    double sec = (rem - min) *60.0;
                    byte[] bytes = new byte[24];
                    uint n1 = (uint)deg; uint d1 =1;
                    uint n2 = (uint)min; uint d2 =1;
                    uint n3 = (uint)Math.Round(sec *1000000.0); uint d3 =1000000;
                    Buffer.BlockCopy(BitConverter.GetBytes(n1),0, bytes,0,4);
                    Buffer.BlockCopy(BitConverter.GetBytes(d1),0, bytes,4,4);
                    Buffer.BlockCopy(BitConverter.GetBytes(n2),0, bytes,8,4);
                    Buffer.BlockCopy(BitConverter.GetBytes(d2),0, bytes,12,4);
                    Buffer.BlockCopy(BitConverter.GetBytes(n3),0, bytes,16,4);
                    Buffer.BlockCopy(BitConverter.GetBytes(d3),0, bytes,20,4);
                    return bytes;
                };

                var latItem = (PropertyItem)System.Runtime.Serialization.FormatterServices.GetUninitializedObject(typeof(PropertyItem));
                latItem.Id =0x0002; latItem.Type =5; latItem.Value = ToRationalTriplet(lat); latItem.Len = latItem.Value.Length;

                var lonItem = (PropertyItem)System.Runtime.Serialization.FormatterServices.GetUninitializedObject(typeof(PropertyItem));
                lonItem.Id =0x0004; lonItem.Type =5; lonItem.Value = ToRationalTriplet(lon); lonItem.Len = lonItem.Value.Length;

                Action<PropertyItem> SetOrReplace = (pi) =>
                {
                    for (int i =0; i < originalPropertyItems.Length; i++)
                    {
                        if (originalPropertyItems[i] != null && originalPropertyItems[i].Id == pi.Id) 
                        { 
                            originalPropertyItems[i] = pi; 
                            return; 
                        }
                    }
                    for (int i =0; i < originalPropertyItems.Length; i++)
                    {
                        if (originalPropertyItems[i] == null) 
                        { 
                            originalPropertyItems[i] = pi; 
                            return; 
                        }
                    }
                    var list = new System.Collections.Generic.List<PropertyItem>(originalPropertyItems); 
                    list.Add(pi); 
                    originalPropertyItems = list.ToArray();
                };

                SetOrReplace(latRef); SetOrReplace(lonRef); SetOrReplace(latItem); SetOrReplace(lonItem);

                gpsLocationAvailable = true;

                Task.Run(async () =>
                {
                    var addr = await GetImageAddressAsync(originalPropertyItems).ConfigureAwait(false);
                    if (!string.IsNullOrEmpty(addr))
                    {
                        if (this.IsHandleCreated)
                        this.BeginInvoke(new Action(() => { photoInfo.Text = $"Original resolution: {originalImage.Width} x {originalImage.Height} Location: {addr}"; locationInputTextBox.Visible = false; }));
                    }
                    else
                    {
                        if (this.IsHandleCreated)
                        this.BeginInvoke(new Action(() => { photoInfo.Text = $"Original resolution: {originalImage.Width} x {originalImage.Height} Location: {lat.ToString(System.Globalization.CultureInfo.InvariantCulture)},{lon.ToString(System.Globalization.CultureInfo.InvariantCulture)}"; locationInputTextBox.Visible = false; }));
                    }
                });
            }
            catch { }
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
                // Support HEIC/HEIF via ImageMagick: convert to JPEG in-memory so System.Drawing can read EXIF
                bool usedMagick = false;
                if (ext == ".heic" || ext == ".heif")
                {
                    try
                    {
                        using (var mag = new MagickImage(bytes))
                        {
                            // preserve profiles by writing as JPEG
                            using (var ms2 = new MemoryStream())
                            {
                                mag.Write(ms2, MagickFormat.Jpeg);
                                ms2.Position =0;
                                using (var img = Image.FromStream(ms2))
                                {
                                    usedMagick = true;
                                    // existing logic continues here
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

                                    originalImage = new Bitmap(img);

                                    // Capture metadata (PropertyItems) from the original image, if any
                                    try
                                    {
                                        var ids = img.PropertyIdList;
                                        if (ids != null && ids.Length >0)
                                        {
                                            originalPropertyItems = new PropertyItem[ids.Length];
                                            for (int i = 0; i < ids.Length; i++)
                                            {
                                                try
                                                {
                                                    if (img.GetPropertyItem(ids[i]).Id == EXIFOrientationID)
                                                    {
                                                        // Handle EXIF Orientation tag to rotate the image to correct orientation
                                                        if (img.GetPropertyItem(ids[i]).Value.Length >= 2)
                                                        {
                                                            ushort orientation = BitConverter.ToUInt16(img.GetPropertyItem(ids[i]).Value, 0);
                                                            if (orientation == 3)
                                                                originalImage.RotateFlip(RotateFlipType.Rotate180FlipNone);
                                                            else if (orientation == 6)
                                                                originalImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
                                                            else if (orientation == 8)
                                                                originalImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
                                                        }
                                                        // Don't set orientation metadata on the saved image since we already applied it;
                                                        // set to 1 (Horizontal)
                                                    }
                                                    else
                                                    {
                                                        originalPropertyItems[i] = img.GetPropertyItem(ids[i]);
                                                    }
                                                }
                                                catch { originalPropertyItems[i] = null; }
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

                                    // Determine if GPS location EXIF data exists
                                    try
                                    {
                                        gpsLocationAvailable = false;
                                        if (originalPropertyItems != null && originalPropertyItems.Length >0)
                                        {
                                            // Check for GPS tags:0x0001..0x0004 (refs and coords) or presence of0x0002/0x0004
                                            foreach (var pi in originalPropertyItems)
                                            {
                                                if (pi == null) continue;
                                                if (pi.Id ==0x0001 || pi.Id ==0x0002 || pi.Id ==0x0003 || pi.Id ==0x0004)
                                                {
                                                    gpsLocationAvailable = true;
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            gpsLocationAvailable = false;
                                        }
                                    }
                                    catch { gpsLocationAvailable = false; }

                                    // Create a new fullImage twice as large in each dimension, filled with black,
                                    // and draw the original image centered inside it.
                                    int max = Math.Max(originalImage.Width, originalImage.Height);
                                    int fullW = Math.Min(16000, max * 3);
                                    int fullH = Math.Min(16000, max * 3);
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

                                    // Set initial zoomFactor so that the original image fits within the overlay rectangle
                                    if (originalImage.Width > originalImage.Height)
                                    {
                                        aspectComboBox.SelectedItem = "16:9";
                                    }
                                    else
                                    {
                                        aspectComboBox.SelectedItem = "4:5";
                                    }
                                    zoomFactor = maxZoomFactor(true, originalImage.Width, originalImage.Height, overlayRectangle.Width, overlayRectangle.Height);

                                    updatePictureBoxImage();

                                    // Update photo info label with original image resolution and location if available
                                    try
                                    {
                                        string info = string.Empty;
                                        try { info = $"Original resolution: {originalImage.Width} x {originalImage.Height}"; } catch { info = string.Empty; }
                                        try
                                        {
                                            // If GPS data not available, indicate that
                                            if (!gpsLocationAvailable)
                                            {
                                                if (!string.IsNullOrEmpty(info)) info += " ";
                                                info += "No location info";
                                                try { locationInputTextBox.Visible = true; locationInputTextBox.Text = string.Empty; locationInputTextBox.Focus(); } catch { }
                                            }
                                            else
                                            {
                                                try { locationInputTextBox.Visible = false; } catch { }
                                                // Attempt to resolve GPS coordinates to a human-readable address using Google Geocode API
                                                try
                                                {
                                                    var addrTask = GetImageAddressAsync(originalPropertyItems);
                                                    addrTask.Wait(2000); // wait up to2s for quick UI update; non-blocking fallback below
                                                    string addr = addrTask.Status == System.Threading.Tasks.TaskStatus.RanToCompletion ? addrTask.Result : string.Empty;
                                                    if (string.IsNullOrEmpty(addr))
                                                    {
                                                        // If geocoding did not complete quickly, continue asynchronously to update label when done
                                                        addrTask.ContinueWith(t => {
                                                            try
                                                            {
                                                                var a = t.Status == System.Threading.Tasks.TaskStatus.RanToCompletion ? t.Result : string.Empty;
                                                                if (!string.IsNullOrEmpty(a))
                                                                {
                                                                    if (this.IsHandleCreated)
                                                                    this.BeginInvoke(new Action(() => { try { photoInfo.Text = info + (string.IsNullOrEmpty(info) ? "" : " ") + "Location: " + a; } catch { } }));
                                                                }
                                                            }
                                                            catch { }
                                                        }, TaskScheduler.Default);
                                                    }

                                                    if (!string.IsNullOrEmpty(addr))
                                                    {
                                                        if (!string.IsNullOrEmpty(info)) info += " ";
                                                        info += "Location: " + addr;
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                        catch { }
                                        photoInfo.Text = info;
                                    }
                                    catch { photoInfo.Text = string.Empty; }
                                }
                            }
                        }
                    }
                    catch (Exception magEx)
                    {
                        UpdateStatus($"Error reading HEIC image: {magEx.Message}");
                        return;
                    }
                }
                if (!usedMagick)
                {
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

                            originalImage = new Bitmap(img);

                            // Capture metadata (PropertyItems) from the original image, if any
                            try
                            {
                                var ids = img.PropertyIdList;
                                if (ids != null && ids.Length >0)
                                {
                                    originalPropertyItems = new PropertyItem[ids.Length];
                                    for (int i = 0; i < ids.Length; i++)
                                    {
                                        try
                                        {
                                            if (img.GetPropertyItem(ids[i]).Id == EXIFOrientationID)
                                            {
                                                // Handle EXIF Orientation tag to rotate the image to correct orientation
                                                if (img.GetPropertyItem(ids[i]).Value.Length >= 2)
                                                {
                                                    ushort orientation = BitConverter.ToUInt16(img.GetPropertyItem(ids[i]).Value, 0);
                                                    if (orientation == 3)
                                                        originalImage.RotateFlip(RotateFlipType.Rotate180FlipNone);
                                                    else if (orientation == 6)
                                                        originalImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
                                                    else if (orientation == 8)
                                                        originalImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
                                                }
                                                // Don't set orientation metadata on the saved image since we already applied it;
                                                // set to 1 (Horizontal)
                                            }
                                            else
                                            {
                                                originalPropertyItems[i] = img.GetPropertyItem(ids[i]);
                                            }
                                        }
                                        catch { originalPropertyItems[i] = null; }
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

                            // Determine if GPS location EXIF data exists
                            try
                            {
                                gpsLocationAvailable = false;
                                if (originalPropertyItems != null && originalPropertyItems.Length >0)
                                {
                                    // Check for GPS tags:0x0001..0x0004 (refs and coords) or presence of0x0002/0x0004
                                    foreach (var pi in originalPropertyItems)
                                    {
                                        if (pi == null) continue;
                                        if (pi.Id ==0x0001 || pi.Id ==0x0002 || pi.Id ==0x0003 || pi.Id ==0x0004)
                                        {
                                            gpsLocationAvailable = true;
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    gpsLocationAvailable = false;
                                }
                            }
                            catch { gpsLocationAvailable = false; }

                            // Create a new fullImage twice as large in each dimension, filled with black,
                            // and draw the original image centered inside it.
                            int max = Math.Max(originalImage.Width, originalImage.Height);
                            int fullW = Math.Min(16000, max * 3);
                            int fullH = Math.Min(16000, max * 3);
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

                            // Set initial zoomFactor so that the original image fits within the overlay rectangle
                            if (originalImage.Width > originalImage.Height)
                            {
                                aspectComboBox.SelectedItem = "16:9";
                            }
                            else
                            {
                                aspectComboBox.SelectedItem = "4:5";
                            }
                            zoomFactor = maxZoomFactor(true, originalImage.Width, originalImage.Height, overlayRectangle.Width, overlayRectangle.Height);

                            updatePictureBoxImage();

                            // Update photo info label with original image resolution and location if available
                            try
                            {
                                string info = string.Empty;
                                try { info = $"Original resolution: {originalImage.Width} x {originalImage.Height}"; } catch { info = string.Empty; }
                                try
                                {
                                    // If GPS data not available, indicate that
                                    if (!gpsLocationAvailable)
                                    {
                                        if (!string.IsNullOrEmpty(info)) info += " ";
                                        info += "No location info";
                                        try { locationInputTextBox.Visible = true; locationInputTextBox.Text = string.Empty; locationInputTextBox.Focus(); } catch { }
                                    }
                                    else
                                    {
                                        try { locationInputTextBox.Visible = false; } catch { }
                                        // Attempt to resolve GPS coordinates to a human-readable address using Google Geocode API
                                        try
                                        {
                                            var addrTask = GetImageAddressAsync(originalPropertyItems);
                                            addrTask.Wait(2000); // wait up to2s for quick UI update; non-blocking fallback below
                                            string addr = addrTask.Status == System.Threading.Tasks.TaskStatus.RanToCompletion ? addrTask.Result : string.Empty;
                                            if (string.IsNullOrEmpty(addr))
                                            {
                                                // If geocoding did not complete quickly, continue asynchronously to update label when done
                                                addrTask.ContinueWith(t => {
                                                    try
                                                    {
                                                        var a = t.Status == System.Threading.Tasks.TaskStatus.RanToCompletion ? t.Result : string.Empty;
                                                        if (!string.IsNullOrEmpty(a))
                                                        {
                                                            if (this.IsHandleCreated)
                                                            this.BeginInvoke(new Action(() => { try { photoInfo.Text = info + (string.IsNullOrEmpty(info) ? "" : " ") + "Location: " + a; } catch { } }));
                                                        }
                                                    }
                                                    catch { }
                                                }, TaskScheduler.Default);
                                            }

                                            if (!string.IsNullOrEmpty(addr))
                                            {
                                                if (!string.IsNullOrEmpty(info)) info += " ";
                                                info += "Location: " + addr;
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                catch { }
                                photoInfo.Text = info;
                            }
                            catch { photoInfo.Text = string.Empty; }
                        }
                    }
                }

                string fileName = Path.GetFileName(filePath);
                currentFilePath = filePath;
                string fileSize = FormatFileSize(new FileInfo(filePath).Length);
                string timestamp = DateTime.Now.ToString("HH:mm:ss");

                // Bring application to foreground when a new file is opened
                try { BringToFrontAndActivate(); } catch { }

                int imgW = pictureBox.Image?.Width ?? 0;
                int imgH = pictureBox.Image?.Height ?? 0;

                UpdateStatus($"[{timestamp}] Opened: {fileName} ({fileSize}) - {imgW}x{imgH} (fitted to window)");

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
                    client.Connect(2000);

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
                    for (int i = 0; i < aspectComboBox.Items.Count; i++)
                    {
                        if (string.Equals(aspectComboBox.Items[i].ToString(), settings.Aspect, StringComparison.OrdinalIgnoreCase))
                        {
                            aspectComboBox.SelectedIndex = i;
                            break;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(settings.SaveResolution))
                {
                    for (int i = 0; i < resolutionComboBox.Items.Count; i++)
                    {
                        if (string.Equals(resolutionComboBox.Items[i].ToString(), settings.SaveResolution, StringComparison.OrdinalIgnoreCase))
                        {
                            resolutionComboBox.SelectedIndex = i;
                            saveResolution = settings.SaveResolution;
                            break;
                        }
                    }
                }

                if (settings.FormWidth > 100 && settings.FormHeight > 100)
                {
                    this.Size = new Size(settings.FormWidth, settings.FormHeight);
                }
            }
            catch
            {
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
                    SaveResolution = saveResolution ?? "High",
                    FormWidth = this.Size.Width,
                    FormHeight = this.Size.Height
                };
                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(settingsFilePath, json);
            }
            catch
            {
            }
        }

        // In InitializeComponent(), after creating photoInfo label, add location input textbox (initially hidden)

        private class UserSettings
        {
            public string StorageFolderPath { get; set; }
            public string Aspect { get; set; }
            public string SaveResolution { get; set; }
            public int FormWidth { get; set; }
            public int FormHeight { get; set; }
        }

        // Add Win32 API helper methods inside PhotoViewerForm class
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_SHOW =5;
        private const int SW_RESTORE =9;

        // Call this method to bring the form to front
        private void BringToFrontAndActivate()
        {
         try
         {
         if (this.WindowState == FormWindowState.Minimized)
         {
        ShowWindow(this.Handle, SW_RESTORE);
         }
         SetForegroundWindow(this.Handle);
         this.Activate();
         }
         catch { }
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

            string filePath = args.Length > 0 ? args[0] : string.Empty;

            bool createdNew;
            using (var mutex = new Mutex(true, MUTEX_NAME, out createdNew))
            {
                if (!createdNew)
                {
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
