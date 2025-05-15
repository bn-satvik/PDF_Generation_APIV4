namespace Proj.Utils
{
    public class PdfHeaderModel
    {
        // Path to the logo image file
        public string LogoPath { get; set; }

        // Title of the PDF report (e.g., "Top Attacked Users")
        public string Title { get; set; } = string.Empty;

        // Date when the report was generated (formatted)
        public string GeneratedDate { get; set; }

        // Company name to display on the header
        public string CompanyName { get; set; }

        // Name of the product 
        public string ProductName { get; set; } = string.Empty;

        // Date range covered by the report
        public string DateRange { get; set; } = string.Empty;

        // Constructor uses input values for logo path and company name
        public PdfHeaderModel(string logoPath, string companyName)
        {
            LogoPath = logoPath;
            CompanyName = companyName;
            GeneratedDate = DateTime.Now.ToString("MMM dd, yyyy");
        }
    }
}
