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
        public static byte[] Generate(Stream? imageStream, List<List<string>> tableData, PdfHeaderModel headerModel, PdfFooterModel footerModel)
        {
            var stopwatch = Stopwatch.StartNew(); // Start timing the PDF generation
            Console.WriteLine("PDF generation started.");

            var document = new Document(); // Create a new PDF document

            // Check if we have an image
            bool hasImage = imageStream != null && imageStream.CanRead && imageStream.Length > 0;

            // Check if we have valid table data (at least a header + 1 row)
            bool hasTable = tableData != null && tableData.Count > 1;

            // If no image and no table, throw an error
            if (!hasImage && !hasTable)
                throw new ArgumentException("At least one of image or table data must be provided.");

            if (hasImage)
            {
                var imageBytes = ReadFully(imageStream!); // Read image stream into byte array
                AddImageSection(document, imageBytes, headerModel, footerModel, out _); // Add image to PDF
            }

            if (hasTable)
            {
                var headerRow = tableData![0]; // First row is header
                var dataRows = tableData.Skip(1).ToList(); // Remaining rows are data
                AddTableSection(document, headerRow, dataRows, headerModel, footerModel); // Add table to PDF
            }

            var resultBytes = RenderAndSave(document); // Finalize and save PDF to byte array

            stopwatch.Stop();
            Console.WriteLine($"Total PDF generation time: {stopwatch.ElapsedMilliseconds} ms");

            return resultBytes; // Return the generated PDF bytes
        }

        private static void AddImageSection(Document document, byte[] imageBytes, PdfHeaderModel headerModel, PdfFooterModel footerModel, out string imageUri)
        {
            // Calculate page size based on image dimensions; returns page width/height and image path
            var (pageWidth, pageHeight) = CalculateImagePageSize(imageBytes, out imageUri);

            // Create a new section in the PDF
            var section = document.AddSection();

            // Set page dimensions in centimeters
            section.PageSetup.PageWidth = Unit.FromCentimeter(pageWidth);
            section.PageSetup.PageHeight = Unit.FromCentimeter(pageHeight);

            // Set margins around the page
            section.PageSetup.LeftMargin = Unit.FromCentimeter(PdfConstants.MarginCm);
            section.PageSetup.RightMargin = Unit.FromCentimeter(PdfConstants.MarginCm);
            section.PageSetup.TopMargin = Unit.FromCentimeter(PdfConstants.MarginCm);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(PdfConstants.MarginCm);

            // Add header and footer using provided layout and model
            PdfHeaderLayout.BuildHeader(section, headerModel);
            PdfFooterLayout.BuildFooter(footerModel, section);

            // Add image to the center of the section
            var paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = Unit.FromCentimeter(1); // Space above the image
            paragraph.Format.Alignment = ParagraphAlignment.Center; // Center align image

            var image = paragraph.AddImage(imageUri); // Add image from file path
            image.Width = $"{PdfConstants.TargetImageWidthCm}cm"; // Set image width (height adjusts automatically)
            image.LockAspectRatio = true; // Maintain original aspect ratio
        }

        private static (double PageWidth, double PageHeight) CalculateImagePageSize(byte[] imageBytes, out string imageUri)
        {
            // Convert image bytes to base64 URI format for embedding in PDF
            string base64Image = Convert.ToBase64String(imageBytes);
            imageUri = "base64:" + base64Image;

            // Load image to get its dimensions
            using var ms = new MemoryStream(imageBytes);
            var imgInfo = Image.Identify(ms);
            if (imgInfo == null)
                throw new InvalidOperationException("Could not identify image.");

            // Convert image dimensions from pixels to centimeters
            double originalWidthCm = imgInfo.Width / PdfConstants.Dpi * PdfConstants.InchesToCentimeters;
            double originalHeightCm = imgInfo.Height / PdfConstants.Dpi * PdfConstants.InchesToCentimeters;

            // Maintain aspect ratio when resizing to target width
            double aspectRatio = originalHeightCm / originalWidthCm;
            double newHeightCm = PdfConstants.TargetImageWidthCm * aspectRatio;

            // Ensure minimum page width and height to fit image and prevent layout issues
            double pageWidth = Math.Max(PdfConstants.MinImagePageWidthCm, PdfConstants.TargetImageWidthCm + 2 * PdfConstants.MarginCm);
            double pageHeight = Math.Max(PdfConstants.MinImagePageHeightCm, newHeightCm + 2 * PdfConstants.MarginCm + PdfConstants.HeaderFooterPaddingCm);

            return (pageWidth, pageHeight); // Return calculated page dimensions
        }

        private static void AddTableSection(Document document, List<string> headerRow, List<List<string>> dataRows, PdfHeaderModel headerModel, PdfFooterModel footerModel)
        {
            // Create a new section for the table
            var section = document.AddSection();

            int columnCount = headerRow.Count;
            var columnWidths = new double[columnCount];

            // Calculate maximum allowed column width based on column count
            double maxColWidthCm = GetDynamicMaxColWidthCm(columnCount);

            // Calculate width of each column based on data and limits
            for (int i = 0; i < columnCount; i++)
                columnWidths[i] = CalculateColumnWidth(headerRow, dataRows, i, PdfConstants.CharWidthCm, PdfConstants.MinColWidthCm, maxColWidthCm);

            // Calculate total table width
            double tableWidth = columnWidths.Sum();

            // Set page width: fit table with padding, within max/min limits
            double pageWidth = Math.Max(Math.Min(tableWidth + PdfConstants.TablePaddingCm, PdfConstants.MaxPageWidthCm), PdfConstants.DefaultPageWidthCm);

            // Center the table on the page
            double margin = (pageWidth - tableWidth) / 2;

            // Set page size and margins
            section.PageSetup.PageWidth = Unit.FromCentimeter(pageWidth);
            section.PageSetup.PageHeight = Unit.FromCentimeter(PdfConstants.PageHeightCm);
            section.PageSetup.LeftMargin = Unit.FromCentimeter(margin);
            section.PageSetup.RightMargin = Unit.FromCentimeter(margin);
            section.PageSetup.TopMargin = Unit.FromCentimeter(PdfConstants.TableTopMarginCm);
            section.PageSetup.BottomMargin = Unit.FromCentimeter(PdfConstants.TableBottomMarginCm);

            // Add header and footer to the section
            PdfHeaderLayout.BuildHeader(section, headerModel);
            PdfFooterLayout.BuildFooter(footerModel, section);

            // Create the table object
            var table = section.AddTable();
            table.Borders.Width = 0;
            table.Format.Font.Size = PdfConstants.DefaultFontSize;
            table.KeepTogether = false;

            // Define each column in the table with calculated widths
            for (int i = 0; i < columnCount; i++)
            {
                var column = table.AddColumn(Unit.FromCentimeter(columnWidths[i]));
                column.Format.Font.Size = PdfConstants.ColumnFontSize;
            }

            // Add and format the header row (bold)
            var headerRowObj = table.AddRow();
            headerRowObj.HeadingFormat = true;
            headerRowObj.Format.Font.Bold = true;
            BuildHeaderRow(headerRowObj, headerRow);

            // Add all data rows to the table
            AddDataRows(dataRows, table, columnCount);
        }

        // Enum to categorize the number of columns into size buckets
        private enum ColumnSizeCategory
        {
            Small,       // 1 to 4 columns
            Medium,      // 5 to 12 columns
            Large,       // 13 to 15 columns
            ExtraLarge   // 16 or more columns
        }

        // Returns maximum column width in cm based on total column count
        private static double GetDynamicMaxColWidthCm(int columnCount)
        {
            // Determine column size category
            ColumnSizeCategory category = columnCount switch
            {
                >= 1 and <= 4 => ColumnSizeCategory.Small,
                >= 5 and <= 12 => ColumnSizeCategory.Medium,
                >= 13 and <= 15 => ColumnSizeCategory.Large,
                _ => ColumnSizeCategory.ExtraLarge,
            };
            // Return max column width based on the category
            return category switch
            {
                ColumnSizeCategory.Small => PdfConstants.MaxColWidthSmall,
                ColumnSizeCategory.Medium => PdfConstants.MaxColWidthMedium,
                ColumnSizeCategory.Large => PdfConstants.MaxColWidthLarge,
                ColumnSizeCategory.ExtraLarge => PdfConstants.MaxColWidthExtraLarge,
                _ => PdfConstants.MaxColWidthExtraLarge // Fallback
            };
        }

        // Calculate column width in cm based on content length and limits
        private static double CalculateColumnWidth(List<string> headerRow, List<List<string>> dataRows, int columnIndex, double charWidthCm, double minCm, double maxCm)
        {
            // Get header text or empty if null
            var header = headerRow[columnIndex] ?? "";

            // Find longest word in header to avoid wrapping
            int maxWordLen = header.Split(' ', StringSplitOptions.RemoveEmptyEntries).DefaultIfEmpty("").Max(w => w.Length);
            maxWordLen += PdfConstants.WordSpace; // Add space between words

            // Collect lengths of all cells in this column
            var lengths = new List<int> { header.Length };
            foreach (var row in dataRows)
            {
                if (columnIndex < row.Count && row[columnIndex] != null)
                    lengths.Add(row[columnIndex].Length);
            }

            int maxLen = lengths.Max();
            int maxValue = Math.Max(maxWordLen, Math.Max(maxLen, PdfConstants.referenceLenMax));

            // Convert char count to cm and clamp between min and max widths
            return Math.Min(Math.Max(maxValue * charWidthCm, minCm), maxCm);
        }

        // Build header row cells with formatted text and style
        private static void BuildHeaderRow(Row headerRowObj, List<string> headerRow)
        {
            for (int i = 0; i < headerRow.Count; i++)
            {
                var cell = headerRowObj.Cells[i];
                // Add paragraph with soft breaks for long text
                var para = cell.AddParagraph(InsertSoftBreaks(headerRow[i], PdfConstants.SoftBreakInterval));
                para.Format.Alignment = ParagraphAlignment.Left;
                para.Format.Font.Size = PdfConstants.HeaderFontSize;
                para.Format.SpaceBefore = PdfConstants.HeaderSpaceBefore;
                para.Format.SpaceAfter = PdfConstants.HeaderSpaceAfter;

                cell.VerticalAlignment = VerticalAlignment.Center;
                cell.Borders.Bottom.Width = PdfConstants.CellBorderWidth; // Add bottom border
            }
        }

        // Adds data rows to the table with formatting and borders
        private static void AddDataRows(List<List<string>> dataRows, Table table, int columnCount)
        {
            foreach (var row in dataRows)
            {
                var dataRow = table.AddRow();
                for (int j = 0; j < columnCount; j++)
                {
                    var cell = dataRow.Cells[j];
                    var text = j < row.Count ? InsertSoftBreaks(row[j], PdfConstants.SoftBreakInterval) : "";
                    var para = cell.AddParagraph(text);
                    para.Format.Alignment = ParagraphAlignment.Left;
                    para.Format.Font.Size = PdfConstants.DefaultFontSize;
                    para.Format.SpaceBefore = PdfConstants.DataSpaceBefore;
                    para.Format.SpaceAfter = PdfConstants.DataSpaceAfter;

                    cell.VerticalAlignment = VerticalAlignment.Center;
                    cell.Borders.Bottom.Width = PdfConstants.CellBorderWidth;
                    cell.Borders.Bottom.Color = Colors.Gray;
                }
            }
        }

        // Renders the document to a PDF and returns it as a byte array
        private static byte[] RenderAndSave(Document document)
        {
            var pdfRenderer = new PdfDocumentRenderer { Document = document };
            pdfRenderer.RenderDocument();
            using var ms = new MemoryStream();
            pdfRenderer.PdfDocument.Save(ms, false);
            return ms.ToArray();
        }

        // Reads the entire stream and returns its content as a byte array
        private static byte[] ReadFully(Stream input)
        {
            using var ms = new MemoryStream();
            input.CopyTo(ms);
            return ms.ToArray();
        }

        // Inserts zero-width spaces into strings at intervals for soft line breaks
        private static string InsertSoftBreaks(string input, int interval)
        {
            if (string.IsNullOrEmpty(input) || input.Length < interval)
                return input;

            var zeroWidthSpace = "\u200B";
            var sb = new StringBuilder();

            for (int i = 0; i < input.Length; i++)
            {
                if (i > 0 && i % interval == 0)
                    sb.Append(zeroWidthSpace);
                sb.Append(input[i]);
            }
            return sb.ToString();
        }
    }
}
