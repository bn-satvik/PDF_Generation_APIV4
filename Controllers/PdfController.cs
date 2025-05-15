using Microsoft.AspNetCore.Mvc;
using Proj.Utils;
using System.Text.Json;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.Diagnostics;
using Microsoft.Extensions.Primitives;

namespace Proj.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PdfController : ControllerBase
    {
        private readonly ILogger<PdfController> _logger;

        public PdfController(ILogger<PdfController> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// POST /api/pdf/generate
        /// Accepts multipart/form-data with optional image, optional CSV file, and required JSON metadata.
        /// Generates a PDF with the provided content and returns it as downloadable file.
        /// </summary>
        [HttpPost("generate")]
        public async Task<IActionResult> GeneratePdf()
        {
            try
            {
                var overallStopwatch = Stopwatch.StartNew();

                // Read the full form data from the request asynchronously
                var form = await Request.ReadFormAsync();

                // Extract optional files from the form data
                var image = form.Files["image"];
                var csvFile = form.Files["tableData"];

                // Extract required metadata string from form data
                StringValues metadataRaw = form["metadata"];

                // Validate mandatory metadata field
                if (string.IsNullOrWhiteSpace(metadataRaw))
                    return BadRequest("Missing metadata.");

                // Validate at least one of image or CSV file is provided
                if (image == null && csvFile == null)
                    return BadRequest("At least one of image or CSV file must be provided.");

                string metadataRawString = metadataRaw.ToString();

                // Process the PDF request with extracted inputs (image, csv, metadata)
                var result = await ProcessPdfRequestAsync(image, csvFile, metadataRawString);

                if (result == null)
                    return BadRequest("Invalid input or processing error.");

                overallStopwatch.Stop();
                Console.WriteLine($"Total request processing time: {overallStopwatch.ElapsedMilliseconds} ms\n");

                // Return the generated PDF as a downloadable file
                return File(result.Value.PdfBytes, "application/pdf", result.Value.FileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating PDF");
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        /// <summary>
        /// Processes input files and metadata to generate PDF bytes and filename.
        /// Made synchronous internally but returns Task for async compatibility.
        /// </summary>
        private Task<(byte[] PdfBytes, string FileName)?> ProcessPdfRequestAsync(IFormFile? image, IFormFile? csvFile, string metadataRaw)
        {
            try
            {
                List<List<string>>? tableData = null;

                // If CSV file provided, parse its content into a 2D string list (table)
                if (csvFile != null)
                {
                    tableData = ParseCsvFile(csvFile);
                    if (tableData == null || tableData.Count < 1)
                    {
                        Console.WriteLine("CSV must contain at least a header row.");
                        return Task.FromResult<(byte[] PdfBytes, string FileName)?>(null);
                    }
                }

                // Parse metadata JSON into dictionary
                var metadata = ParseMetadataWithLogging(metadataRaw);
                if (metadata == null)
                    return Task.FromResult<(byte[] PdfBytes, string FileName)?>(null);

                // Create header model using constants + metadata
                string logoPath = "Utils/barracuda_logo.png";
                string companyName = "Barracuda Networks";
                var headerModel = CreateHeaderModel(metadata, logoPath, companyName);

                // Create footer model (showing page numbers)
                var footerModel = new PdfFooterModel { ShowPageNumbers = true };

                byte[] pdfBytes;

                // Generate PDF based on presence of image and/or CSV data
                if (image != null && tableData != null)
                {
                    using var imageStream = image.OpenReadStream();
                    pdfBytes = PdfGenerator.Generate(imageStream, tableData, headerModel, footerModel);
                }
                else if (image != null)
                {
                    using var imageStream = image.OpenReadStream();
                    pdfBytes = PdfGenerator.Generate(imageStream, new List<List<string>>(), headerModel, footerModel);
                }
                else if (tableData != null)
                {
                    // If only CSV table data is present, generate PDF without image stream
                    pdfBytes = PdfGenerator.Generate(null, tableData, headerModel, footerModel);
                }
                else
                {
                    // If neither image nor CSV, return null
                    return Task.FromResult<(byte[] PdfBytes, string FileName)?>(null);
                }

                // Create a unique output file name based on Title and timestamp
                string fileName = CreateOutputFileName(headerModel);

                return Task.FromResult<(byte[] PdfBytes, string FileName)?>( (pdfBytes, fileName) );
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to process PDF request.");
                return Task.FromResult<(byte[] PdfBytes, string FileName)?>(null);
            }
        }

        /// <summary>
        /// Reads CSV file content and parses it into a list of rows, each row being a list of cell strings.
        /// </summary>
        private List<List<string>> ParseCsvFile(IFormFile csvFile)
        {
            var stopwatch = Stopwatch.StartNew();

            List<List<string>> tableData;
            using (var csvStream = csvFile.OpenReadStream())
            {
                tableData = ParseCsvWithCsvHelper(csvStream);
            }

            stopwatch.Stop();
            Console.WriteLine($"CSV parsing took {stopwatch.ElapsedMilliseconds} ms");
            return tableData;
        }

        /// <summary>
        /// Parses metadata JSON string and logs time taken.
        /// Returns dictionary of key-value pairs or null if invalid.
        /// </summary>
        private Dictionary<string, string>? ParseMetadataWithLogging(string metadataRaw)
        {
            var stopwatch = Stopwatch.StartNew();

            var metadata = ParseMetadata(metadataRaw);
            if (metadata == null)
            {
                Console.WriteLine("Failed to parse metadata.");
                return null;
            }

            stopwatch.Stop();
            Console.WriteLine($"Metadata parsing took {stopwatch.ElapsedMilliseconds} ms");
            return metadata;
        }

        /// <summary>
        /// Creates a filename for the generated PDF based on the title and current timestamp.
        /// </summary>
        private string CreateOutputFileName(PdfHeaderModel header)
        {
            var stopwatch = Stopwatch.StartNew();

            var fileName = $"{header.Title}_{DateTime.Now:MMM_dd_yyyy_HHmmss}.pdf";

            stopwatch.Stop();
            Console.WriteLine($"Filename creation took {stopwatch.ElapsedMilliseconds} ms");
            return fileName;
        }

        /// <summary>
        /// Deserializes JSON string into Dictionary of metadata key-value pairs.
        /// Returns null if invalid JSON.
        /// </summary>
        private Dictionary<string, string>? ParseMetadata(string metadataRaw)
        {
            try
            {
                return JsonSerializer.Deserialize<Dictionary<string, string>>(metadataRaw);
            }
            catch (JsonException ex)
            {
                _logger.LogError($"Invalid JSON format for metadata: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Builds the PDF header model object using metadata and constants.
        /// </summary>
        private PdfHeaderModel CreateHeaderModel(Dictionary<string, string> metadata, string logoPath, string companyName)
        {
            var header = new PdfHeaderModel(logoPath, companyName)
            {
                Title = metadata.TryGetValue("Title", out var title) ? title : "N/A",
                ProductName = metadata.TryGetValue("ProductName", out var productName) ? productName : "N/A",
                DateRange = metadata.TryGetValue("DateRange", out var dateRange) ? dateRange : "N/A"
            };

            return header;
        }

        /// <summary>
        /// Uses CsvHelper to read CSV from stream and converts each row to list of string values.
        /// Returns list of rows.
        /// </summary>
        private List<List<string>> ParseCsvWithCsvHelper(Stream csvStream)
        {
            var records = new List<List<string>>();

            using (var reader = new StreamReader(csvStream, Encoding.UTF8))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                Delimiter = ","
            }))
            {
                // Read each row and extract all fields into a string list
                while (csv.Read())
                {
                    var row = new List<string>();
                    for (int i = 0; csv.TryGetField(i, out string? value); i++)
                    {
                        row.Add(value ?? string.Empty);
                    }
                    records.Add(row);
                }
            }

            return records;
        }
    }
}
