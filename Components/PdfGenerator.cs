using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using SixLabors.ImageSharp;
using System.Text;
using System.Diagnostics;

namespace Proj.Utils
{
    public static class PdfGenerator
    {
        // Constants for layout and formatting
        private const double Dpi = 96; // Dots per inch, for image scaling
        private const double InchesToCentimeters = 2.54; // Conversion factor
        private const double TargetImageWidthCm = 15.0; // Desired image width in cm
        private const double MarginCm = 1.0; // Page margin in cm
        private const double MinImagePageWidthCm = 16.0; // Minimum image page width
        private const double MinImagePageHeightCm = 16.0; // Minimum image page height
        private const double CharWidthCm = 0.23; // Estimated character width in cm
        private const double MinColWidthCm = 2.0; // Minimum column width
        private const int referenceLenMax = 10; // Reference length for column sizing

        private const int WordSpace = 2; // Space after word length
        private const double MaxPageWidthCm = 75.0; // Max width for table pages
        private const double DefaultPageWidthCm = 21.0; // Default page width (A4)
        private const double PageHeightCm = 34.0; // Default page height
        private const double TableTopMarginCm = 2.0; // Table top margin
        private const double TableBottomMarginCm = 2.5; // Table bottom margin
        private const double HeaderFooterPaddingCm = 6.0; // Padding for header/footer
        private const double TablePaddingCm = 3.0; // Extra spacing for table
        private const int DefaultFontSize = 10;
        private const int HeaderFontSize = 12;
        private const int ColumnFontSize = 9;
        private const string HeaderSpaceBefore = "0.3cm";
        private const string HeaderSpaceAfter = "0.3cm";
        private const string DataSpaceBefore = "0.15cm";
        private const string DataSpaceAfter = "0.15cm";
        private const double CellBorderWidth = 0.5;
        private const int SoftBreakInterval = 20; // Interval to insert soft breaks in text

        // Main method to generate the PDF
        public static byte[] Generate(Stream imageStream, List<List<string>> tableData, PdfHeaderModel headerModel, PdfFooterModel footerModel)
        {
            var stopwatch = Stopwatch.StartNew();
            Console.WriteLine("PDF generation started.");

            var headerRow = tableData[0]; // First row is table header
            var dataRows = tableData.Skip(1).ToList(); // Remaining rows are data
            var document = new Document(); // Create new PDF document

            var imageBytes = ReadFully(imageStream); // Convert image stream to byte array
            AddImageSection(document, imageBytes, headerModel, footerModel, out string imageUri); // Add image page

            AddTableSection(document, headerRow, dataRows, headerModel, footerModel); // Add table page

            var resultBytes = RenderAndSave(document); // Render PDF
            stopwatch.Stop();

            Console.WriteLine($"Total PDF generation time: {stopwatch.ElapsedMilliseconds} ms");
            return resultBytes; // Return generated PDF as byte array
        }

        // Adds a page with the image to the document
        private static void AddImageSection(Document document, byte[] imageBytes, PdfHeaderModel headerModel, PdfFooterModel footerModel, out string imageUri)
        {
            var (pageWidth, pageHeight) = CalculateImagePageSize(imageBytes, out imageUri);
            var section = document.AddSection();

            // Set page size and margins
            section.PageSetup.PageWidth = Unit.FromCentimeter(pageWidth);
            section.PageSetup.PageHeight = Unit.FromCentimeter(pageHeight);
            section.PageSetup.LeftMargin = Unit.FromCentimeter(MarginCm);
            section.PageSetup.RightMargin = Unit.FromCentimeter(MarginCm);
            section.PageSetup.TopMargin = Unit.FromCentimeter(MarginCm);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(MarginCm);

            PdfHeaderLayout.BuildHeader(section, headerModel); // Add header
            PdfFooterLayout.BuildFooter(footerModel, section); // Add footer

            var paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = Unit.FromCentimeter(1);
            paragraph.Format.Alignment = ParagraphAlignment.Center;

            var image = paragraph.AddImage(imageUri); // Add image to paragraph
            image.Width = $"{TargetImageWidthCm}cm";
            image.LockAspectRatio = true;
        }

        // Adds a page with the table to the document
        private static void AddTableSection(Document document, List<string> headerRow, List<List<string>> dataRows, PdfHeaderModel headerModel, PdfFooterModel footerModel)
        {
            var section = document.AddSection();
            int columnCount = headerRow.Count;
            var columnWidths = new double[columnCount];
            double MaxColWidthCm = GetDynamicMaxColWidthCm(columnCount);

            // Calculate individual column widths
            for (int i = 0; i < columnCount; i++)
                columnWidths[i] = CalculateColumnWidth(headerRow, dataRows, i, CharWidthCm, MinColWidthCm, MaxColWidthCm);

            double tableWidth = columnWidths.Sum();
            double pageWidth = Math.Max(Math.Min(tableWidth + TablePaddingCm, MaxPageWidthCm), DefaultPageWidthCm);
            double margin = (pageWidth - tableWidth) / 2;

            // Set page properties
            section.PageSetup.PageWidth = Unit.FromCentimeter(pageWidth);
            section.PageSetup.PageHeight = Unit.FromCentimeter(PageHeightCm);
            section.PageSetup.LeftMargin = Unit.FromCentimeter(margin);
            section.PageSetup.RightMargin = Unit.FromCentimeter(margin);
            section.PageSetup.TopMargin = Unit.FromCentimeter(TableTopMarginCm);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(TableBottomMarginCm);

            PdfHeaderLayout.BuildHeader(section, headerModel); // Add header
            PdfFooterLayout.BuildFooter(footerModel, section); // Add footer

            var table = section.AddTable(); 
            table.Borders.Width = 0;
            table.Format.Font.Size = DefaultFontSize;
            table.KeepTogether = false;

            // Create columns
            for (int i = 0; i < columnCount; i++)
            {
                var column = table.AddColumn(Unit.FromCentimeter(columnWidths[i]));
                column.Format.Font.Size = ColumnFontSize;
            }

            // Add header row
            var headerRowObj = table.AddRow();
            headerRowObj.HeadingFormat = true;
            headerRowObj.Format.Font.Bold = true;
            BuildHeaderRow(headerRowObj, headerRow);

            // Add data rows
            AddDataRows(dataRows, table, columnCount);
        }

        // Calculates optimal image page size
        private static (double PageWidth, double PageHeight) CalculateImagePageSize(byte[] imageBytes, out string imageUri)
        {
            string base64Image = Convert.ToBase64String(imageBytes);
            imageUri = "base64:" + base64Image;

            using var ms = new MemoryStream(imageBytes);
            var imgInfo = Image.Identify(ms);
            if (imgInfo == null)
                throw new InvalidOperationException("Could not identify image.");

            double originalWidthCm = imgInfo.Width / Dpi * InchesToCentimeters;
            double originalHeightCm = imgInfo.Height / Dpi * InchesToCentimeters;
            double aspectRatio = originalHeightCm / originalWidthCm;
            double newHeightCm = TargetImageWidthCm * aspectRatio;

            double pageWidth = Math.Max(MinImagePageWidthCm, TargetImageWidthCm + 2 * MarginCm);
            double pageHeight = Math.Max(MinImagePageHeightCm, newHeightCm + 2 * MarginCm + HeaderFooterPaddingCm);

            return (pageWidth, pageHeight);
        }

        // Adjusts max column width based on column count
        private static double GetDynamicMaxColWidthCm(int columnCount)
        {
            switch (columnCount)
            {
                case int n when (n >= 1 && n <= 4): //Small
                    return 11;
                case int n when (n >= 5 && n <= 12):// Medium
                    return 9;
                case int n when (n >= 13 && n <= 15):// Large
                    return 5;
                default:
                    return 3; // More than 16 columns
            }
        }

        // Builds header row of the table
        private static void BuildHeaderRow(Row headerRowObj, List<string> headerRow)
        {
            for (int i = 0; i < headerRow.Count; i++)
            {
                var cell = headerRowObj.Cells[i];
                var para = cell.AddParagraph(InsertSoftBreaks(headerRow[i], SoftBreakInterval));
                para.Format.Alignment = ParagraphAlignment.Left;
                para.Format.Font.Size = HeaderFontSize;
                para.Format.SpaceBefore = HeaderSpaceBefore;
                para.Format.SpaceAfter = HeaderSpaceAfter;
                cell.VerticalAlignment = VerticalAlignment.Center;
                cell.Borders.Bottom.Width = CellBorderWidth;
            }
        }

        // Adds all data rows to the table
        private static void AddDataRows(List<List<string>> dataRows, Table table, int columnCount)
        {
            foreach (var row in dataRows)
            {
                var dataRow = table.AddRow();
                for (int j = 0; j < columnCount; j++)
                {
                    var cell = dataRow.Cells[j];
                    var text = j < row.Count ? InsertSoftBreaks(row[j], SoftBreakInterval) : "";
                    var para = cell.AddParagraph(text);
                    para.Format.Alignment = ParagraphAlignment.Left;
                    para.Format.Font.Size = DefaultFontSize;
                    para.Format.SpaceBefore = DataSpaceBefore;
                    para.Format.SpaceAfter = DataSpaceAfter;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    cell.Borders.Bottom.Width = CellBorderWidth;
                    cell.Borders.Bottom.Color = Colors.Gray;
                }
            }
        }

        // Calculates width for each column based on header and data content
        private static double CalculateColumnWidth(List<string> headerRow, List<List<string>> dataRows, int columnIndex, double charWidthCm, double minCm, double maxCm)
        {
            var header = headerRow[columnIndex] ?? "";
            int maxWordLen = header.Split(' ', StringSplitOptions.RemoveEmptyEntries).DefaultIfEmpty("").Max(w => w.Length);
            maxWordLen += WordSpace;

            var lengths = new List<int> { header.Length };
            foreach (var row in dataRows)
            {
                if (columnIndex < row.Count && row[columnIndex] != null)
                    lengths.Add(row[columnIndex].Length);
            }

            int maxLen = lengths.Max();
            int referenceLen = referenceLenMax;
            int maxValue = Math.Max(maxWordLen, Math.Max(maxLen, referenceLen));
            return Math.Min(Math.Max(maxValue * charWidthCm, minCm), maxCm);
        }

        // Final PDF rendering and saving
        private static byte[] RenderAndSave(Document document)
        {
            var pdfRenderer = new PdfDocumentRenderer { Document = document };
            pdfRenderer.RenderDocument();
            using var ms = new MemoryStream();
            pdfRenderer.PdfDocument.Save(ms, false);
            return ms.ToArray();
        }

        // Reads entire stream into byte array
        private static byte[] ReadFully(Stream input)
        {
            using var ms = new MemoryStream();
            input.CopyTo(ms);
            return ms.ToArray();
        }

        // Adds soft breaks to long strings for better layout
        private static string InsertSoftBreaks(string input, int interval)
        {
            if (string.IsNullOrEmpty(input) || input.Length < interval)
                return input;
            var ZerowidthSpace = "\u200B";
            var sb = new StringBuilder();
            for (int i = 0; i < input.Length; i++)
            {
                if (i > 0 && i % interval == 0)
                    sb.Append(ZerowidthSpace);
                sb.Append(input[i]);
            }
            return sb.ToString();
        }
    }
}
