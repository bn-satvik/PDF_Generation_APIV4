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
        // === Constants for layout and formatting ===
        private const double Dpi = 96;
        private const double InchesToCentimeters = 2.54;
        private const double TargetImageWidthCm = 15.0;
        private const double MarginCm = 1.0;
        private const double MinImagePageWidthCm = 16.0;
        private const double MinImagePageHeightCm = 16.0;
        private const double CharWidthCm = 0.23;
        private const double MinColWidthCm = 2.0;
        private const double MaxColWidthCm = 10.0;
        private const double DefaultPageWidthCm = 21.0; // Default width, but no longer a constraint
        private const double PageHeightCm = 34.0;
        private const double TableTopMarginCm = 2.0;
        private const double TableBottomMarginCm = 2.5;

        private const double HeaderFooterPaddingCm = 6.0;
        private const double TablePaddingCm = 3.0;
        private const int DefaultFontSize = 10;
        private const int HeaderFontSize = 12;
        private const int ColumnFontSize = 9;
        private const string HeaderSpaceBefore = "0.3cm";
        private const string HeaderSpaceAfter = "0.3cm";
        private const string DataSpaceBefore = "0.15cm";
        private const string DataSpaceAfter = "0.15cm";
        private const double CellBorderWidth = 0.5;
        private const int SoftBreakInterval = 20;
        private const int LetterMargin = 2;
        
        // Main method to generate PDF
        public static byte[] Generate(Stream imageStream, List<List<string>> tableData, PdfHeaderModel headerModel, PdfFooterModel footerModel)
        {
            var stopwatch = Stopwatch.StartNew();
            Console.WriteLine("PDF generation started.");

            var headerRow = tableData[0]; // First row is header
            var dataRows = tableData.Skip(1).ToList(); // Remaining rows are data
            var document = new Document(); // Create new PDF document

            // === IMAGE SECTION ===
            var imageBytes = ReadFully(imageStream); // Read image as byte array
            string imageUri;
            var (pageWidth, pageHeight) = CalculateImagePageSize(imageBytes, out imageUri); // Calculate image page size

            var imageSection = document.AddSection(); // Add section for image
            imageSection.PageSetup.PageWidth = Unit.FromCentimeter(pageWidth);
            imageSection.PageSetup.PageHeight = Unit.FromCentimeter(pageHeight);
            imageSection.PageSetup.LeftMargin = Unit.FromCentimeter(MarginCm);
            imageSection.PageSetup.RightMargin = Unit.FromCentimeter(MarginCm);
            imageSection.PageSetup.TopMargin = Unit.FromCentimeter(MarginCm);
            imageSection.PageSetup.BottomMargin = Unit.FromCentimeter(MarginCm);

            PdfHeaderLayout.BuildHeader(imageSection, headerModel); // Add header
            PdfFooterLayout.BuildFooter(footerModel, imageSection); // Add footer

            var imageParagraph = imageSection.AddParagraph(); // Insert image
            imageParagraph.Format.SpaceBefore = Unit.FromCentimeter(1);
            imageParagraph.Format.Alignment = ParagraphAlignment.Center;
            var image = imageParagraph.AddImage(imageUri);
            image.Width = $"{TargetImageWidthCm}cm";
            image.LockAspectRatio = true;

            // === TABLE SECTION ===
            var tableSection = document.AddSection(); // New section for table
            int columnCount = headerRow.Count;
            var columnWidths = new double[columnCount];

            // Calculate dynamic width for each column based on content
            for (int i = 0; i < columnCount; i++)
                columnWidths[i] = CalculateColumnWidth(headerRow, dataRows, i, CharWidthCm, MinColWidthCm, MaxColWidthCm);

            double tableWidth = columnWidths.Sum(); // Total table width
            double pageWidthTable = tableWidth + TablePaddingCm; // No page width limit

            // Setup page size and margins
            tableSection.PageSetup.PageWidth = Unit.FromCentimeter(pageWidthTable);
            tableSection.PageSetup.LeftMargin = Unit.FromCentimeter(MarginCm);
            tableSection.PageSetup.RightMargin = Unit.FromCentimeter(MarginCm);
            tableSection.PageSetup.PageHeight = Unit.FromCentimeter(PageHeightCm);
            tableSection.PageSetup.TopMargin = Unit.FromCentimeter(TableTopMarginCm);
            tableSection.PageSetup.BottomMargin = Unit.FromCentimeter(TableBottomMarginCm);

            PdfHeaderLayout.BuildHeader(tableSection, headerModel); // Add header
            PdfFooterLayout.BuildFooter(footerModel, tableSection); // Add footer

            var table = tableSection.AddTable(); // Create table
            table.Borders.Width = 0;
            table.Format.Font.Size = DefaultFontSize;
            table.KeepTogether = false;

            // Add columns to table
            for (int i = 0; i < columnCount; i++)
            {
                var column = table.AddColumn(Unit.FromCentimeter(columnWidths[i]));
                column.Format.Font.Size = ColumnFontSize;
            }

            // Add table header row
            var headerRowObj = table.AddRow();
            headerRowObj.HeadingFormat = true;
            headerRowObj.Format.Font.Bold = true;
            BuildHeaderRow(headerRowObj, headerRow);

            // Add data rows
            AddDataRows(dataRows, table, columnCount);

            // === RENDER AND SAVE PDF ===
            var resultBytes = RenderAndSave(document); // Render PDF and return as byte array

            stopwatch.Stop();
            Console.WriteLine($"Total PDF generation time: {stopwatch.ElapsedMilliseconds} ms");

            return resultBytes;
        }

        // Calculates page size based on image size and returns Base64 image URI
        private static (double PageWidth, double PageHeight) CalculateImagePageSize(byte[] imageBytes, out string imageUri)
        {
            string base64Image = Convert.ToBase64String(imageBytes);
            imageUri = "base64:" + base64Image;

            using var ms = new MemoryStream(imageBytes);
            var imgInfo = Image.Identify(ms);
            if (imgInfo == null)
                throw new InvalidOperationException("Could not identify image.");

            // Convert pixels to centimeters
            double originalWidthCm = imgInfo.Width / Dpi * InchesToCentimeters;
            double originalHeightCm = imgInfo.Height / Dpi * InchesToCentimeters;
            double aspectRatio = originalHeightCm / originalWidthCm;

            // Scale image height proportionally
            double newHeightCm = TargetImageWidthCm * aspectRatio;

            // Determine page size including margins and header/footer space
            double pageWidth = Math.Max(MinImagePageWidthCm, TargetImageWidthCm + 2 * MarginCm);
            double pageHeight = Math.Max(MinImagePageHeightCm, newHeightCm + 2 * MarginCm + HeaderFooterPaddingCm);

            return (pageWidth, pageHeight);
        }

        // Builds the table header row
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

        // Calculates appropriate column width based on content length
        private static double CalculateColumnWidth(List<string> headerRow, List<List<string>> dataRows, int columnIndex, double charWidthCm, double minCm, double maxCm)
        {
            var header = headerRow[columnIndex] ?? "";
            int maxWordLen = header.Split(' ', StringSplitOptions.RemoveEmptyEntries).DefaultIfEmpty("").Max(w => w.Length);
            var lengths = new List<int> { header.Length };

            foreach (var row in dataRows)
            {
                if (columnIndex < row.Count && row[columnIndex] != null)
                    lengths.Add(row[columnIndex].Length);
            }
            maxWordLen= maxWordLen+ LetterMargin;
            int maxLen = lengths.Max();
            int referenceLen = 10; // Minimum width fallback
            int maxValue = Math.Max(maxWordLen, Math.Max(maxLen, referenceLen));
            return Math.Min(Math.Max(maxValue * charWidthCm, minCm), maxCm);
        }

        // Renders the MigraDoc document into a PDF and returns it as a byte array
        private static byte[] RenderAndSave(Document document)
        {
            var pdfRenderer = new PdfDocumentRenderer { Document = document };
            pdfRenderer.RenderDocument();

            using var ms = new MemoryStream();
            pdfRenderer.PdfDocument.Save(ms, false);
            return ms.ToArray();
        }

        // Reads the entire content of a stream into a byte array
        private static byte[] ReadFully(Stream input)
        {
            using var ms = new MemoryStream();
            input.CopyTo(ms);
            return ms.ToArray();
        }

        // Inserts zero-width space every N characters to allow soft breaks in long words
        private static string InsertSoftBreaks(string input, int interval)
        {
            if (string.IsNullOrEmpty(input) || input.Length < interval)
                return input;

            var sb = new StringBuilder();
            for (int i = 0; i < input.Length; i++)
            {
                if (i > 0 && i % interval == 0)
                    sb.Append('\u200B'); // Zero-width space
                sb.Append(input[i]);
            }
            return sb.ToString();
        }
    }
}
