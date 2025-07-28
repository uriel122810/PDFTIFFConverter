using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfiumViewer;

namespace PDFTIFFConverter
{
    public partial class MainForm : Form
    {
        private List<string> selectedPdfFiles = new List<string>();
        private string outputDirectory = "";

        public MainForm()
        {
            InitializeComponents();
            this.Text = "Procesador de Imágenes - PDF a TIFF";
            this.Icon = SystemIcons.Application;
            this.WindowState = FormWindowState.Maximized;
            this.BackColor = Color.FromArgb(240, 240, 240);
        }

        private void InitializeComponents()
        {
            this.SuspendLayout();

            this.Size = new Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Main panel
            Panel mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20)
            };
            this.Controls.Add(mainPanel);

            // Title
            Label titleLabel = new Label
            {
                Text = "Convertir PDF a TIFF",
                Font = new Font("Segoe UI", 24, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 51, 51),
                AutoSize = true,
                Location = new Point(0, 20)
            };
            mainPanel.Controls.Add(titleLabel);

            // Subtitle
            Label subtitleLabel = new Label
            {
                Text = "Transforma documentos PDF en imágenes TIFF de alta calidad",
                Font = new Font("Segoe UI", 12),
                ForeColor = Color.FromArgb(102, 102, 102),
                AutoSize = true,
                Location = new Point(0, 70)
            };
            mainPanel.Controls.Add(subtitleLabel);

            // Main content panel
            Panel contentPanel = new Panel
            {
                Location = new Point(0, 120),
                Size = new Size(1160, 600),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            mainPanel.Controls.Add(contentPanel);

            // Features panel
            Panel featuresPanel = new Panel
            {
                Location = new Point(20, 20),
                Size = new Size(400, 300),
                BackColor = Color.FromArgb(220, 53, 69),
                Padding = new Padding(20)
            };
            contentPanel.Controls.Add(featuresPanel);

            Label featuresTitle = new Label
            {
                Text = "Esta herramienta te permite:",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(20, 20)
            };
            featuresPanel.Controls.Add(featuresTitle);

            string[] features = {
                "• Convertir PDF a formato TIFF",
                "• Mantener alta calidad de imagen",
                "• Combinar múltiples PDF en un archivo TIFF",
                "• Establecer DPI personalizado (300 DPI)",
                "• Procesar múltiples archivos a la vez",
                "• Convertir a escala de grises"
            };

            for (int i = 0; i < features.Length; i++)
            {
                Label featureLabel = new Label
                {
                    Text = features[i],
                    Font = new Font("Segoe UI", 11),
                    ForeColor = Color.White,
                    AutoSize = true,
                    Location = new Point(20, 60 + (i * 25))
                };
                featuresPanel.Controls.Add(featureLabel);
            }

            // File selection panel
            Panel filePanel = new Panel
            {
                Location = new Point(450, 20),
                Size = new Size(680, 300),
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(20)
            };
            contentPanel.Controls.Add(filePanel);

            Label fileLabel = new Label
            {
                Text = "Seleccionar archivos PDF:",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 51, 51),
                AutoSize = true,
                Location = new Point(20, 20)
            };
            filePanel.Controls.Add(fileLabel);

            Button selectFilesBtn = new Button
            {
                Text = "Seleccionar PDFs",
                Font = new Font("Segoe UI", 11),
                Size = new Size(150, 40),
                Location = new Point(20, 60),
                BackColor = Color.FromArgb(0, 123, 255),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false
            };
            selectFilesBtn.Click += SelectFiles_Click;
            filePanel.Controls.Add(selectFilesBtn);

            ListBox fileListBox = new ListBox
            {
                Name = "fileListBox",
                Location = new Point(20, 110),
                Size = new Size(640, 120),
                Font = new Font("Segoe UI", 10)
            };
            filePanel.Controls.Add(fileListBox);

            Button removeFileBtn = new Button
            {
                Text = "Remover seleccionado",
                Font = new Font("Segoe UI", 10),
                Size = new Size(130, 30),
                Location = new Point(530, 240),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            removeFileBtn.Click += RemoveFile_Click;
            filePanel.Controls.Add(removeFileBtn);

            // Output settings panel
            Panel outputPanel = new Panel
            {
                Location = new Point(20, 340),
                Size = new Size(1110, 120),
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(20)
            };
            contentPanel.Controls.Add(outputPanel);

            Label outputLabel = new Label
            {
                Text = "Configuración de salida:",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 51, 51),
                AutoSize = true,
                Location = new Point(20, 20)
            };
            outputPanel.Controls.Add(outputLabel);

            Button selectOutputBtn = new Button
            {
                Text = "Seleccionar carpeta de destino",
                Font = new Font("Segoe UI", 11),
                Size = new Size(200, 40),
                Location = new Point(20, 60),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            selectOutputBtn.Click += SelectOutput_Click;
            outputPanel.Controls.Add(selectOutputBtn);

            Label outputPathLabel = new Label
            {
                Name = "outputPathLabel",
                Text = "No se ha seleccionado carpeta de destino",
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.FromArgb(102, 102, 102),
                AutoSize = true,
                Location = new Point(240, 75)
            };
            outputPanel.Controls.Add(outputPathLabel);

            // Convert button
            Button convertBtn = new Button
            {
                Text = "Ir a Convertir PDF ➜",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                Size = new Size(300, 50),
                Location = new Point(410, 480),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            convertBtn.Click += Convert_Click;
            contentPanel.Controls.Add(convertBtn);

            // Progress bar
            ProgressBar progressBar = new ProgressBar
            {
                Name = "progressBar",
                Location = new Point(20, 550),
                Size = new Size(1110, 20),
                Visible = false
            };
            contentPanel.Controls.Add(progressBar);

            // Status label
            Label statusLabel = new Label
            {
                Name = "statusLabel",
                Text = "",
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.FromArgb(102, 102, 102),
                AutoSize = true,
                Location = new Point(20, 575)
            };
            contentPanel.Controls.Add(statusLabel);

            this.ResumeLayout(false);
        }

        private void SelectFiles_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "Seleccionar archivos PDF";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedPdfFiles.AddRange(openFileDialog.FileNames);
                    UpdateFileList();
                }
            }
        }

        private void RemoveFile_Click(object sender, EventArgs e)
        {
            var fileListBoxArray = this.Controls.Find("fileListBox", true);
            if (fileListBoxArray.Length > 0)
            {
                ListBox fileListBox = fileListBoxArray[0] as ListBox;
                if (fileListBox != null && fileListBox.SelectedIndex >= 0)
                {
                    selectedPdfFiles.RemoveAt(fileListBox.SelectedIndex);
                    UpdateFileList();
                }
            }
        }

        private void UpdateFileList()
        {
            var fileListBoxArray = this.Controls.Find("fileListBox", true);
            if (fileListBoxArray.Length > 0)
            {
                ListBox fileListBox = fileListBoxArray[0] as ListBox;
                if (fileListBox != null)
                {
                    fileListBox.Items.Clear();
                    foreach (string file in selectedPdfFiles)
                    {
                        fileListBox.Items.Add(Path.GetFileName(file));
                    }
                }
            }
        }

        private void SelectOutput_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Seleccionar carpeta de destino";
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    outputDirectory = folderDialog.SelectedPath;
                    var outputPathLabelArray = this.Controls.Find("outputPathLabel", true);
                    if (outputPathLabelArray.Length > 0)
                    {
                        Label outputPathLabel = outputPathLabelArray[0] as Label;
                        if (outputPathLabel != null)
                        {
                            outputPathLabel.Text = outputDirectory;
                        }
                    }
                }
            }
        }

        private async void Convert_Click(object sender, EventArgs e)
        {
            if (selectedPdfFiles.Count == 0)
            {
                MessageBox.Show("Por favor seleccione al menos un archivo PDF.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(outputDirectory))
            {
                MessageBox.Show("Por favor seleccione la carpeta de destino.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var progressBarArray = this.Controls.Find("progressBar", true);
            var statusLabelArray = this.Controls.Find("statusLabel", true);

            if (progressBarArray.Length == 0 || statusLabelArray.Length == 0)
                return;

            ProgressBar progressBar = progressBarArray[0] as ProgressBar;
            Label statusLabel = statusLabelArray[0] as Label;

            if (progressBar == null || statusLabel == null)
                return;

            progressBar.Visible = true;
            progressBar.Minimum = 0;
            progressBar.Maximum = selectedPdfFiles.Count;
            progressBar.Value = 0;

            try
            {
                List<Image> allImages = new List<Image>();

                await Task.Run(() =>
                {
                    for (int i = 0; i < selectedPdfFiles.Count; i++)
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            statusLabel.Text = $"Procesando: {Path.GetFileName(selectedPdfFiles[i])}";
                        });

                        var images = ConvertPdfToImages(selectedPdfFiles[i]);
                        allImages.AddRange(images);

                        this.Invoke((MethodInvoker)delegate
                        {
                            progressBar.Value = i + 1;
                        });
                    }

                    this.Invoke((MethodInvoker)delegate
                    {
                        statusLabel.Text = "Creando archivo TIFF...";
                    });

                    string outputPath = Path.Combine(outputDirectory,
                        $"converted_{DateTime.Now:yyyyMMdd_HHmmss}.tiff");

                    SaveAsMultipageTiff(allImages, outputPath);

                    // Clean up
                    foreach (var img in allImages)
                    {
                        img.Dispose();
                    }

                    this.Invoke((MethodInvoker)delegate
                    {
                        progressBar.Visible = false;
                        statusLabel.Text = $"Conversión completada: {outputPath}";

                        MessageBox.Show($"Conversión completada exitosamente.\nArchivo guardado en: {outputPath}",
                            "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    });
                });
            }
            catch (Exception ex)
            {
                progressBar.Visible = false;
                statusLabel.Text = "Error durante la conversión";
                MessageBox.Show($"Error durante la conversión: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<Image> ConvertPdfToImages(string pdfPath)
        {
            List<Image> images = new List<Image>();

            try
            {
                using (var document = PdfDocument.Load(pdfPath))
                {
                    for (int i = 0; i < document.PageCount; i++)
                    {
                        // Render at 300 DPI
                        using (var image = document.Render(i, 300, 300, false))
                        {
                            // Convert to grayscale and set DPI
                            var grayImage = ConvertToGrayscale(image);

                            // Establecer explícitamente los DPI
                            grayImage.SetResolution(300f, 300f);

                            images.Add(grayImage);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error procesando PDF {Path.GetFileName(pdfPath)}: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return images;
        }

        private Bitmap ConvertToGrayscale(Image original)
        {
            Bitmap grayscale = new Bitmap(original.Width, original.Height);

            // Establecer la resolución antes de dibujar
            grayscale.SetResolution(300f, 300f);

            using (Graphics g = Graphics.FromImage(grayscale))
            {
                ColorMatrix colorMatrix = new ColorMatrix(
                    new float[][]
                    {
                        new float[] {.3f, .3f, .3f, 0, 0},
                        new float[] {.59f, .59f, .59f, 0, 0},
                        new float[] {.11f, .11f, .11f, 0, 0},
                        new float[] {0, 0, 0, 1, 0},
                        new float[] {0, 0, 0, 0, 1}
                    });

                ImageAttributes attributes = new ImageAttributes();
                attributes.SetColorMatrix(colorMatrix);

                g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height),
                    0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);
            }

            return grayscale;
        }

        private void SaveAsMultipageTiff(List<Image> images, string outputPath)
        {
            if (images.Count == 0) return;

            // Asegurar que todas las imágenes tienen 300 DPI
            foreach (Image img in images)
            {
                if (img.HorizontalResolution != 300f || img.VerticalResolution != 300f)
                {
                    ((Bitmap)img).SetResolution(300f, 300f);
                }
            }

            // Configurar parámetros del encoder con compresión LZW
            EncoderParameters encoderParams = new EncoderParameters(3);
            encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
            encoderParams.Param[1] = new EncoderParameter(Encoder.Compression, (long)EncoderValue.CompressionLZW);
            // Establecer explícitamente la resolución en el encoder
            encoderParams.Param[2] = new EncoderParameter(Encoder.Quality, 100L);

            ImageCodecInfo tiffCodec = GetEncoderInfo("image/tiff");

            // Guardar la primera imagen
            images[0].Save(outputPath, tiffCodec, encoderParams);

            // Agregar las páginas subsecuentes
            for (int i = 1; i < images.Count; i++)
            {
                encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
                images[0].SaveAdd(images[i], encoderParams);
            }

            // Cerrar el archivo TIFF multipágina
            encoderParams.Param[0] = new EncoderParameter(Encoder.SaveFlag, (long)EncoderValue.Flush);
            images[0].SaveAdd(encoderParams);
        }

        private ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            return codecs.FirstOrDefault(codec => codec.MimeType == mimeType);
        }
    }
}