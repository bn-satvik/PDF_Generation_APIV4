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
        private readonly ILogger<PdfController> _logger;

        public PdfController(ILogger<PdfController> logger)
        {
            _logger = logger;
        }

        [HttpPost("generate")]
        public async Task<IActionResult> GeneratePdf()
        {
            try
            {
                Console.WriteLine("************************************************************");
                var overallStopwatch = Stopwatch.StartNew();

                // Read form data from request
                var form = await Request.ReadFormAsync();
                var image = form.Files["image"];
                var csvFile = form.Files["tableData"];
                var metadataRaw = form["metadata"];

                // Validate form inputs
                if (image == null || csvFile == null || string.IsNullOrWhiteSpace(metadataRaw))
                {
                    return BadRequest("Missing image, CSV file, or metadata.");
                }

                // Parse CSV file into table format
                var csvStopwatch = Stopwatch.StartNew();
                List<List<string>> tableData;
                using (var csvStream = csvFile.OpenReadStream())
                {
                    tableData = ParseCsvWithCsvHelper(csvStream);
                }
                csvStopwatch.Stop();
                Console.WriteLine($"CSV parsing took {csvStopwatch.ElapsedMilliseconds} ms");

                if (tableData.Count < 2)
                {
                    return BadRequest("CSV must contain a header row and at least one data row.");
                }

                // Parse metadata JSON into a dictionary
                var metadataStopwatch = Stopwatch.StartNew();
                var metadataRawString = metadataRaw.ToString();
                if (string.IsNullOrWhiteSpace(metadataRawString))
                {
                    return BadRequest("Metadata is missing or empty.");
                }

                var metadata = ParseMetadata(metadataRawString);
                if (metadata == null)
                {
                    return BadRequest("Failed to parse metadata.");
                }
                metadataStopwatch.Stop();
                Console.WriteLine($"Metadata parsing took {metadataStopwatch.ElapsedMilliseconds} ms");

                using var imageStream = image.OpenReadStream();

                // Build PDF header and footer models
                string logoPath = "Utils/barracuda_logo.png"; 
                string companyName = "Barracuda Networks";

                var headerModel = CreateHeaderModel(metadata, logoPath, companyName);
                var footerModel = new PdfFooterModel { ShowPageNumbers = true };

                // Generate PDF using utility method
                var pdfStopwatch = Stopwatch.StartNew();
                byte[] pdfBytes = PdfGenerator.Generate(imageStream, tableData, headerModel, footerModel);
                pdfStopwatch.Stop();
                Console.WriteLine($"PDF generation took {pdfStopwatch.ElapsedMilliseconds} ms");

                // Create filename from title date and time
                var fileNameStopwatch = Stopwatch.StartNew();
                string fileName = $"{headerModel.Title}_{DateTime.Now:MMM_dd_yyyy_HHmmss}.pdf";
                fileNameStopwatch.Stop();
                Console.WriteLine($"Filename creation took {fileNameStopwatch.ElapsedMilliseconds} ms");

                // Log total processing time
                overallStopwatch.Stop();
                Console.WriteLine($"Total request processing time: {overallStopwatch.ElapsedMilliseconds} ms");

                // Return PDF as downloadable file
                return File(pdfBytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error generating PDF");
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        // Converts JSON metadata string to Dictionary
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

        // Creates the header model for the PDF using metadata
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

        // Reads and parses CSV file using CsvHelper into list of rows
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
