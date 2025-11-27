using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using Ghostscript.NET.Rasterizer; // PDF'ten PNG'ye dönüşüm için gereklidir.
using System.Drawing;
using Aspose.Cells;
using Aspose.Words;
using Aspose.Pdf;
using SaveFormat = Aspose.Cells.SaveFormat;
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Aspose.Cells.Worksheet;
using PageSetup = Aspose.Cells.PageSetup;
using System.Drawing.Imaging; // Image sınıfı için gereklidir.
using Excel = Microsoft.Office.Interop.Excel;
using ImageMagick;
using Word = Microsoft.Office.Interop.Word;
using Aspose.Pdf.Drawing;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Path = System.IO.Path;
using Action = System.Action;
using Application = System.Windows.Forms.Application;
using Color = System.Drawing.Color;
using System.Text.RegularExpressions;
using PdfSharp.Pdf.IO;

namespace excel_to_pdf_levent_beylerce
{
    public partial class Form1 : Form
    {
        // Uygulamanın ana formu için başlangıç kodu
        public Form1()
        {
            InitializeComponent();


        }


        // DragEnter olayı
        private void ListBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        // DragDrop olayı
        private void ListBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            foreach (var file in files)
            {
                string fileName = Path.GetFileName(file);
                if (!listBoxPaths.ContainsKey(fileName))
                {
                    listBox1.Items.Add(fileName);
                    listBoxPaths[fileName] = file;
                }
            }

            // Drag & Drop ile dosya geldiğinde buton text'ini değiştir
            buttonAddConvert.Text = "Convert Et";
        }



        // Ana işlemi başlatan düğme tıklama olayı
        private void button1_Click(object sender, EventArgs e)
        {

        }
        // ========= TIFF Encoder Helper =========
        private ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            return Array.Find(ImageCodecInfo.GetImageEncoders(),
                c => c.MimeType == mimeType);
        }


        private void button2_Click(object sender, EventArgs e)
        {
            PdfDocument pdf = new PdfDocument();

            foreach (var tif in Directory.GetFiles(@"\tempDir", "*.tif"))
            {
                PdfPage page = pdf.AddPage();
                page.Orientation = PdfSharp.PageOrientation.Landscape;

                XGraphics xg = XGraphics.FromPdfPage(page);
                using (PdfSharp.Drawing.XImage img = PdfSharp.Drawing.XImage.FromFile(tif))
                {
                    xg.DrawImage(img, 0, 0, page.Width, page.Height);
                }
            }

            pdf.Save("C:\\Cikti.pdf");

        }

        public void ExceliPDFyeCevir(string excelDosyaYolu, string pdfDosyaYolu)
        {

        }
        public class ExcelImageHelper
        {
            public static string[] ConvertExcelToHighResPng(string excelFile, string outputFolder, int dpi = 150)
            {
                Directory.CreateDirectory(outputFolder);

                // Geçici PDF yolu
                string tempPdf = System.IO.Path.Combine(outputFolder, "temp_" + Guid.NewGuid() + ".pdf");

                Excel.Application excelApp = null;
                Excel.Workbook wb = null;

                try
                {
                    excelApp = new Excel.Application();
                    excelApp.DisplayAlerts = false;

                    wb = excelApp.Workbooks.Open(excelFile, ReadOnly: true);

                    // Workbook -> PDF
                    wb.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, tempPdf);

                    // PDF -> PNG (Yüksek DPI)
                    var settings = new MagickReadSettings()
                    {
                        Density = new Density(dpi, dpi)
                    };

                    var pngList = new System.Collections.Generic.List<string>();

                    using (var images = new MagickImageCollection())
                    {
                        images.Read(tempPdf, settings);

                        for (int i = 0; i < images.Count; i++)
                        {
                            var img = (MagickImage)images[i];
                            img.BackgroundColor = MagickColors.White;
                            img.Alpha(AlphaOption.Remove);

                            string outPath = Path.Combine(outputFolder, $"page_{(i + 1):D8}.png");
                            img.Write(outPath);

                            pngList.Add(outPath);

                            // ProgressBar update (Form1’deki progressBar1 ve lblProgress için)
                            if (Application.OpenForms.Count > 0)
                            {
                                var form = Application.OpenForms[0] as Form1;
                                form?.Invoke((Action)(() =>
                                {
                                    int percent = (i + 1) * 100 / images.Count;

                                }));
                            }
                        }
                    }

                    return pngList.ToArray();
                }
                finally
                {
                    // COM temizliği
                    if (wb != null)
                    {
                        wb.Close(false);
                        Marshal.ReleaseComObject(wb);
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    if (File.Exists(tempPdf))
                        File.Delete(tempPdf);
                }
            }
        }


        public static string ConvertPngListToPdf(string[] pngPaths, string outputPdfPath)
        {
            PdfDocument pdf = new PdfDocument();
            pdf.Info.Title = "PNG to PDF";

            foreach (var png in pngPaths)
            {
                if (!File.Exists(png)) continue;

                // PNG boyutunu al
                using (var img = PdfSharp.Drawing.XImage.FromFile(png))
                {
                    // Yeni bir PDF sayfası oluştur
                    PdfPage page = pdf.AddPage();

                    // PNG dikey mi? yoksa yatay mı?
                    bool isPortrait = img.PixelHeight >= img.PixelWidth;

                    if (isPortrait)
                    {
                        // Dikey A4
                        page.Width = XUnit.FromMillimeter(210);
                        page.Height = XUnit.FromMillimeter(297);
                        page.Orientation = PdfSharp.PageOrientation.Portrait;
                    }
                    else
                    {
                        // Yatay A4
                        page.Width = XUnit.FromMillimeter(297);
                        page.Height = XUnit.FromMillimeter(210);
                        page.Orientation = PdfSharp.PageOrientation.Landscape;
                    }

                    using (XGraphics gfx = XGraphics.FromPdfPage(page))
                    {
                        // PNG boyutunu sayfaya sığdıracak şekilde ölçekle
                        double ratioX = page.Width / img.PixelWidth * img.HorizontalResolution / 72;
                        double ratioY = page.Height / img.PixelHeight * img.VerticalResolution / 72;
                        double ratio = Math.Min(ratioX, ratioY);

                        double width = img.PixelWidth * ratio * 72 / img.HorizontalResolution;
                        double height = img.PixelHeight * ratio * 72 / img.VerticalResolution;

                        // Sayfanın ortasına yerleştir
                        double x = (page.Width - width) / 2;
                        double y = (page.Height - height) / 2;

                        gfx.DrawImage(img, x, y, width, height);
                    }
                }

            }

            pdf.Save(outputPdfPath);
            return outputPdfPath;
        }
        public static string[] ConvertPngToSafeFormat(string[] pngPaths)
        {
            string[] safePaths = new string[pngPaths.Length];

            for (int i = 0; i < pngPaths.Length; i++)
            {
                string original = pngPaths[i];
                string safe = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(original),
                             System.IO.Path.GetFileNameWithoutExtension(original) + "_safe.png");

                using (var img = new MagickImage(original))
                {
                    // 8-bit RGB ve Alpha kaldır
                    img.Format = MagickFormat.Png8;
                    img.BackgroundColor = MagickColors.White;
                    img.Alpha(AlphaOption.Remove);

                    img.Write(safe);
                }

                safePaths[i] = safe;
            }

            return safePaths;
        }
        public static void ClearFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            try
            {
                // Tüm dosyaları sil
                foreach (var file in Directory.GetFiles(folderPath))
                {
                    File.Delete(file);
                }

                // Eğer klasör içinde alt klasörler varsa onları da sil
                foreach (var dir in Directory.GetDirectories(folderPath))
                {
                    Directory.Delete(dir, true); // true = alt klasörler ve dosyalar dahil sil
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Klasör temizlenirken hata oluştu:\n" + ex.Message);
            }
        }
        public class WordImageHelper
        {
            public static string[] ConvertWordToHighResPng(string wordFile, string outputFolder, int dpi = 150)
            {
                Directory.CreateDirectory(outputFolder);

                string tempPdf = System.IO.Path.Combine(outputFolder, "temp_" + Guid.NewGuid() + ".pdf");

                Word.Application wordApp = null;
                Word.Document doc = null;

                try
                {
                    wordApp = new Word.Application();
                    wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                    // Word belgesini aç
                    doc = wordApp.Documents.Open(wordFile, ReadOnly: true, Visible: false);

                    // Word → PDF
                    doc.ExportAsFixedFormat(
                        tempPdf,
                        Word.WdExportFormat.wdExportFormatPDF,
                        OpenAfterExport: false,
                        OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                        Range: Word.WdExportRange.wdExportAllDocument
                    );

                    // PDF → PNG
                    var settings = new MagickReadSettings()
                    {
                        Density = new Density(dpi, dpi)
                    };

                    var pngList = new List<string>();

                    using (var images = new MagickImageCollection())
                    {
                        images.Read(tempPdf, settings);

                        for (int i = 0; i < images.Count; i++)
                        {
                            var img = (MagickImage)images[i];
                            img.BackgroundColor = MagickColors.White;
                            img.Alpha(AlphaOption.Remove);

                            string outPath = Path.Combine(outputFolder, $"page_{(i + 1):D8}.png");
                            img.Write(outPath);
                            pngList.Add(outPath);

                            // ProgressBar güncellemesi (Form1)
                            if (Application.OpenForms.Count > 0)
                            {
                                var form = Application.OpenForms[0] as Form1;
                                form?.Invoke((Action)(() =>
                                {
                                    int percent = (i + 1) * 100 / images.Count;
                                }));
                            }
                        }
                    }

                    return pngList.ToArray();
                }
                finally
                {
                    // COM temizliği
                    if (doc != null)
                    {
                        doc.Close(false);
                        Marshal.ReleaseComObject(doc);
                    }

                    if (wordApp != null)
                    {
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    if (File.Exists(tempPdf))
                        File.Delete(tempPdf);
                }
            }
        }


        private string[] GetSelectionOrder(OpenFileDialog ofd)
        {
            // Explorer gerçek seçim sırasını SafeFileNames dizesindeki sıraya göre verir.
            return ofd.FileNames
                     .Select(path => new
                     {
                         Path = path,
                         Index = Array.IndexOf(ofd.SafeFileNames, Path.GetFileName(path))
                     })
                     .OrderBy(x => x.Index)
                     .Select(x => x.Path)
                     .ToArray();
        }

        public int GetExcelPageCount(string filePath)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(filePath);
            int sheetCount = wb.Sheets.Count;
            wb.Close(false);
            app.Quit();
            return sheetCount;
        }

        public int GetWordPageCount(string filePath)
        {
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Open(filePath, ReadOnly: true, Visible: false);
            int pageCount = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
            doc.Close(false);
            app.Quit();
            return pageCount;
        }


        // ListBox'ta sadece adlar, tam yolları saklamak için Dictionary
        private Dictionary<string, string> listBoxPaths = new Dictionary<string, string>();
        private int convertSessionIndex = 1; // Her convert için ayrı klasör numarası
        private async void buttonAddConvert_Click(object sender, EventArgs e)
        {
            if (buttonAddConvert.Text == "DOSYA EKLE (Excel  &&  Word)")
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Filter = "Excel - Word Dosyaları|*.xlsx;*.xls;*.docx;*.doc";
                    ofd.Multiselect = true;
                    ofd.Title = "Excel veya Word dosyaları seçin";

                    if (ofd.ShowDialog() != DialogResult.OK) return;

                    foreach (var file in ofd.FileNames)
                    {
                        string fileName = Path.GetFileName(file);
                        if (!listBoxPaths.ContainsKey(fileName))
                        {
                            listBox1.Items.Add(fileName);
                            listBoxPaths[fileName] = file;
                        }
                    }

                    buttonAddConvert.Text = "Convert Et";
                }
            }

            else if (buttonAddConvert.Text == "Convert Et")
            {
                buttonAddConvert.Enabled = false;
                progressBar1.Value = 0;
                labelLog.Text = "";
                button4.Visible = true;

                try
                {
                    string appFolder = AppDomain.CurrentDomain.BaseDirectory;
                    string sessionFolder = Path.Combine(appFolder, $"Convert_{convertSessionIndex:D4}");
                    string pngFolder = Path.Combine(sessionFolder, "pngler");
                    string pdfFolder = Path.Combine(sessionFolder, "pdfler");

                    Directory.CreateDirectory(pngFolder);
                    Directory.CreateDirectory(pdfFolder);

                    List<string> selectedFiles = new List<string>();
                    foreach (var item in listBox1.Items)
                    {
                        string fileName = item.ToString();
                        if (listBoxPaths.ContainsKey(fileName))
                            selectedFiles.Add(listBoxPaths[fileName]);
                    }

                    if (selectedFiles.Count == 0)
                    {
                        MessageBox.Show("ListBox boş. Lütfen dosya ekleyin.");
                        return;
                    }

                    // Tahmini toplam PNG sayısı = her dosya için 1
                    int totalEstimatedSteps = selectedFiles.Count * 2; // PNG ve Safe PNG için minimum 2 adım
                    progressBar1.Maximum = totalEstimatedSteps;
                    int progressCounter = 0;

                    await Task.Run(() =>
                    {
                        int fileIndex = 0;
                        List<string> createdPdfFiles = new List<string>();

                        foreach (string file in selectedFiles)
                        {
                            string ext = Path.GetExtension(file).ToLower();
                            string baseName = Path.GetFileNameWithoutExtension(file);
                            string safeBase = MakeSafeFileName(baseName);

                            string filePngFolder = Path.Combine(pngFolder, $"{fileIndex:D4}_{safeBase}");
                            Directory.CreateDirectory(filePngFolder);

                            string[] pngPaths = null;

                            if (ext == ".doc" || ext == ".docx")
                                pngPaths = WordImageHelper.ConvertWordToHighResPng(file, filePngFolder, 200);
                            else if (ext == ".xls" || ext == ".xlsx")
                                pngPaths = ExcelImageHelper.ConvertExcelToHighResPng(file, filePngFolder, 200);

                            if (pngPaths == null || pngPaths.Length == 0)
                                throw new Exception("PNG üretilmedi!");

                            // PNG logları ve progress
                            foreach (var png in pngPaths)
                            {
                                progressCounter++;
                                this.Invoke((Action)(() =>
                                {
                                    labelLog.Text = $"[{fileIndex}] {baseName} → PNG oluşturuldu: {Path.GetFileName(png)}";
                                    progressBar1.Value = Math.Min(progressCounter, progressBar1.Maximum);
                                }));
                            }

                            var sortedPng = pngPaths.OrderBy(x => Path.GetFileName(x)).ToArray();
                            var safePngs = ConvertPngToSafeFormat(sortedPng);

                            // Safe PNG logları ve progress
                            foreach (var safePng in safePngs)
                            {
                                progressCounter++;
                                this.Invoke((Action)(() =>
                                {
                                    labelLog.Text = $"[{fileIndex}] {baseName} → Safe PNG: {Path.GetFileName(safePng)}";
                                    progressBar1.Value = Math.Min(progressCounter, progressBar1.Maximum);
                                }));
                            }

                            string pdfOutput = Path.Combine(pdfFolder, $"{fileIndex:D4}_{safeBase}.pdf");
                            ConvertPngListToPdf(safePngs, pdfOutput);
                            createdPdfFiles.Add(pdfOutput);

                            this.Invoke((Action)(() =>
                            {
                                labelLog.Text = $"[{fileIndex}] {baseName} → PDF oluşturuldu: {Path.GetFileName(pdfOutput)}";
                            }));

                            fileIndex++;
                        }

                        // PDF birleştirme sorusu
                        DialogResult ask = DialogResult.None;
                        var form = Application.OpenForms[0] as Form1;

                        form.Invoke((Action)(() =>
                        {
                            ask = MessageBox.Show(
                                "Oluşturulan PDF’leri tek bir PDF’de birleştirmek ister misiniz?",
                                "PDF Birleştirme",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);
                        }));

                        if (ask == DialogResult.Yes)
                        {
                            string finalPdf = Path.Combine(pdfFolder, "RaporFinal.pdf");
                            MergePdfFiles(createdPdfFiles.ToArray(), finalPdf);

                            form.Invoke((Action)(() =>
                            {
                                labelLog.Text = "PDF'ler birleştirildi → RaporFinal.pdf";
                            }));
                        }
                    });

                    Process.Start("explorer.exe", sessionFolder);
                    MessageBox.Show("PDF işlemi tamamlandı.", "Tamam", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    convertSessionIndex++; // Sonraki convert için klasör numarasını artır
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hata oluştu:\n" + ex.Message);
                }
                finally
                {
                    buttonAddConvert.Enabled = true;
                    progressBar1.Value = 0;
                }
            }
            if (buttonAddConvert.Text == "Convert Et")
            {
                     button4.Visible =true;
               
            }
        }


        // Yukarı Buton
        private void buttonUp_Click(object sender, EventArgs e)
        {
            int index = listBox1.SelectedIndex;
            if (index > 0)
            {
                var item = listBox1.Items[index];
                listBox1.Items.RemoveAt(index);
                listBox1.Items.Insert(index - 1, item);
                listBox1.SelectedIndex = index - 1;
            }
        }





        // Dosya adlarını güvenli hale getirme
        public string MakeSafeFileName(string input)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                input = input.Replace(c, '_');
            return input;
        }

        public void MergePdfFiles(string[] files, string outputPdf)
        {
            PdfDocument output = new PdfDocument();

            foreach (string pdf in files.OrderBy(x => x))
            {
                PdfDocument input = PdfReader.Open(pdf, PdfDocumentOpenMode.Import);
                foreach (PdfPage page in input.Pages)
                {
                    output.AddPage(page);
                }
            }

            output.Save(outputPdf);
        }



        private async Task ProcessFileAsync(string file)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Updater updater = new Updater();
            //  await updater.CheckUpdate();
            listBox1.AllowDrop = true;
            listBox1.DragEnter += ListBox1_DragEnter;
            listBox1.DragDrop += ListBox1_DragDrop;

            if (buttonAddConvert.Text == "DOSYA EKLE (Excel  &&  Word)")
            {
                button4.Visible = false;
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy; // Dosya bırakılabilir
            else
                e.Effect = DragDropEffects.None;
        }

        private async void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files.Length == 0) return;

            string file = files[0]; // İlk dosya
            string ext = Path.GetExtension(file).ToLower();

            if (ext != ".xlsx" && ext != ".xls" && ext != ".docx" && ext != ".doc")
            {
                MessageBox.Show("Sadece Excel veya Word dosyası bırakabilirsiniz!");
                return;
            }

            // Burada mevcut button3_Click kodunu async olarak çağırabiliriz
            await ProcessFileAsync(file);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            int index = listBox1.SelectedIndex;
            if (index > 0)
            {
                var item = listBox1.Items[index];
                listBox1.Items.RemoveAt(index);
                listBox1.Items.Insert(index - 1, item);
                listBox1.SelectedIndex = index - 1; // seçim değişen yerde kalsın
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            int index = listBox1.SelectedIndex;
            if (index >= 0 && index < listBox1.Items.Count - 1)
            {
                var item = listBox1.Items[index];
                listBox1.Items.RemoveAt(index);
                listBox1.Items.Insert(index + 1, item);
                listBox1.SelectedIndex = index + 1; // seçim değişen yerde kalsın
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Seçilen bir şey yoksa çık
            if (listBox1.SelectedItems.Count == 0)
            {
                MessageBox.Show("Silinecek bir öğe seçilmedi.");
                return;
            }

            // Seçili öğeleri tersten sil
            for (int i = listBox1.SelectedIndices.Count - 1; i >= 0; i--)
            {
                int index = listBox1.SelectedIndices[i];
                string item = listBox1.Items[index].ToString();

                // Dictionary'den de kaldır
                if (listBoxPaths.ContainsKey(item))
                    listBoxPaths.Remove(item);

                // ListBox'tan kaldır
                listBox1.Items.RemoveAt(index);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel - Word Dosyaları|*.xlsx;*.xls;*.docx;*.doc";
                ofd.Multiselect = true;
                ofd.Title = "Excel veya Word dosyaları seçin";

                if (ofd.ShowDialog() != DialogResult.OK) return;

                foreach (var file in ofd.FileNames)
                {
                    string fileName = Path.GetFileName(file);
                    if (!listBoxPaths.ContainsKey(fileName))
                    {
                        listBox1.Items.Add(fileName);
                        listBoxPaths[fileName] = file;
                    }
                }

                buttonAddConvert.Text = "Convert Et";
            }
        }
    }
}