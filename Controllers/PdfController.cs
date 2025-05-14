using Microsoft.AspNetCore.Mvc;
using Proj.Utils;
using System.Text.Json;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.Diagnostics;

namespace Proj.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PdfController : ControllerBase
    {
        // Logger for internal error/info tracking
        private readonly ILogger<PdfController> _logger;

        // Constructor: assigns logger
        public PdfController(ILogger<PdfController> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Handles POST request to generate a PDF.
        /// Takes an image, CSV, and metadata.
        /// Returns a downloadable PDF file.
        /// </summary>
        [HttpPost("generate")]
        public async Task<IActionResult> GeneratePdf()
        {
            try
            {
                Console.WriteLine("************************************************************");
                var overallStopwatch = Stopwatch.StartNew(); // Measures full request time

                // Read uploaded form files (image + CSV + metadata)
                var form = await Request.ReadFormAsync();
                var image = form.Files["image"];
                var csvFile = form.Files["tableData"];
                var metadataRaw = form["metadata"];

                // Validate required inputs
                if (image == null || csvFile == null || string.IsNullOrWhiteSpace(metadataRaw))
                    return BadRequest("Missing image, CSV file, or metadata.");

                // --- Parse CSV into table format ---
                var csvStopwatch = Stopwatch.StartNew();
                List<List<string>> tableData;
                using (var csvStream = csvFile.OpenReadStream())
                {
                    tableData = ParseCsvWithCsvHelper(csvStream); // Convert CSV to 2D list
                }
                csvStopwatch.Stop();
                Console.WriteLine($"CSV parsing took {csvStopwatch.ElapsedMilliseconds} ms");

                // Ensure there’s a header + data rows
                if (tableData.Count < 2)
                    return BadRequest("CSV must contain a header row and at least one data row.");

                // --- Parse metadata JSON into dictionary ---
                var metadataStopwatch = Stopwatch.StartNew();
                var metadataRawString = metadataRaw.ToString();
                if (string.IsNullOrWhiteSpace(metadataRawString))
                    return BadRequest("Metadata is missing or empty.");

                var metadata = ParseMetadata(metadataRawString); // Deserialize JSON
                if (metadata == null)
                    return BadRequest("Failed to parse metadata.");
                metadataStopwatch.Stop();
                Console.WriteLine($"Metadata parsing took {metadataStopwatch.ElapsedMilliseconds} ms");

                using var imageStream = image.OpenReadStream();

                // --- Create header & footer models ---
                string logoPath = "Utils/barracuda_logo.png"; // Static logo
                string companyName = "Barracuda Networks";     // Static company name

                // Use metadata to fill header details
                var headerModel = CreateHeaderModel(metadata, logoPath, companyName);
                var footerModel = new PdfFooterModel { ShowPageNumbers = true }; // Footer has page numbers

                // --- Generate PDF document ---
                var pdfStopwatch = Stopwatch.StartNew();
                byte[] pdfBytes = PdfGenerator.Generate(imageStream, tableData, headerModel, footerModel);
                pdfStopwatch.Stop();
                Console.WriteLine($"PDF generation took {pdfStopwatch.ElapsedMilliseconds} ms");

                // --- Generate filename from metadata + timestamp ---
                var fileNameStopwatch = Stopwatch.StartNew();
                string fileName = $"{headerModel.Title}_{DateTime.Now:MMM_dd_yyyy_HHmmss}.pdf";
                fileNameStopwatch.Stop();
                Console.WriteLine($"Filename creation took {fileNameStopwatch.ElapsedMilliseconds} ms");

                // --- Total time logging ---
                overallStopwatch.Stop();
                Console.WriteLine($"Total request processing time: {overallStopwatch.ElapsedMilliseconds} ms");

                // --- Return PDF file for download ---
                return File(pdfBytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error generating PDF");
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts metadata JSON string to Dictionary.
        /// Logs and returns null on failure.
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
        /// Builds PDF header model using metadata values.
        /// </summary>
        private PdfHeaderModel CreateHeaderModel(Dictionary<string, string> metadata, string logoPath, string companyName)
        {
            return new PdfHeaderModel
            {
                LogoPath = logoPath,
                Title = metadata.TryGetValue("Title", out var title) ? title : "N/A",
                GeneratedDate = DateTime.Now.ToString("MMM dd, yyyy"),
                CompanyName = companyName,
                ProuductName = metadata.TryGetValue("ProductName", out var productName) ? productName : "N/A",
                DateRange = metadata.TryGetValue("DateRange", out var dateRange) ? dateRange : "N/A"
            };
        }

        /// <summary>
        /// Reads CSV file and returns list of string lists (rows).
        /// Uses CsvHelper to support proper parsing.
        /// </summary>
        private List<List<string>> ParseCsvWithCsvHelper(Stream csvStream)
        {
            var records = new List<List<string>>();
            using (var reader = new StreamReader(csvStream, Encoding.UTF8))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                Delimiter = "," // Comma separated values
            }))
            {
                // Loop through CSV rows
                while (csv.Read())
                {
                    var row = new List<string>();
                    for (int i = 0; csv.TryGetField(i, out string? value); i++)
                    {
                        row.Add(value ?? string.Empty); // Ensure no null values
                    }
                    records.Add(row); // Add row to list
                }
            }
            return records;
        }
    }
}
