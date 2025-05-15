using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;

namespace Proj.Utils
{
    public static class PdfHeaderLayout
    {
        // Top margin for the PDF page header in centimeters (string format for MigraDoc)
        private const string TopMarginCm = "5cm";

        // Dimensions for the right header frame (width and height as strings for MigraDoc)
        private const string RightFrameWidth = "7cm";
        private const string RightFrameHeight = "4cm";

        // Position and margin constants for the right header frame (in centimeters as doubles)
        private const double RightFrameTopCm = 2.1;
        private const double RightFrameWidthCm = 6;
        private const double RightMarginCm = 2.5;

        // Font sizes used for text inside the right header frame
        private const int CompanyFontSize = 14;
        private const int InspectorFontSize = 12;
        private const int DateRangeFontSize = 12;

        // Space after paragraphs in the right header (in centimeters as string)
        private const string RightParagraphSpaceAfterCm = "0.45cm";

        // Dimensions for the left header frame (width, height, top and left offsets)
        private const string LeftFrameWidth = "10cm";
        private const string LeftFrameHeight = "4cm";
        private const string LeftFrameTop = "1.5cm";
        private const string LeftFrameLeft = "1.5cm";

        // Width of the company logo image in the left header frame
        private const string LogoWidth = "4cm";

        // Font size and spacing constants for the report title in the left header
        private const int TitleFontSize = 16;
        private const string TitleSpaceBefore = "0.3cm";
        private const string TitleSpaceAfter = "0.3cm";

        // Font size for the generated date text in the left header
        private const int GeneratedDateFontSize = 12;

        // Main method to build the header for the provided PDF section and header model
        public static void BuildHeader(Section section, PdfHeaderModel headerModel)
        {
            // Set the top margin of the page section header to the defined constant
            section.PageSetup.TopMargin = TopMarginCm;

            // Create and add the left portion of the header (logo, title, generated date)
            CreateLeftHeader(section, headerModel);

            // Create and add the right portion of the header (company, product, date range)
            CreateRightHeader(section, headerModel);
        }

        // Creates the right header frame and fills it with company, product, and date range info
        private static void CreateRightHeader(Section section, PdfHeaderModel headerModel)
        {
            // Add a text frame to the primary header of the section with defined width and height
            var rightFrame = section.Headers.Primary.AddTextFrame();
            rightFrame.Width = RightFrameWidth;
            rightFrame.Height = RightFrameHeight;

            // Position the frame relative to the whole page, not the section content area
            rightFrame.RelativeVertical = RelativeVertical.Page;
            rightFrame.RelativeHorizontal = RelativeHorizontal.Page;

            // Calculate the left position of the frame to align it from the right edge of the page
            var pageWidth = section.PageSetup.PageWidth;
            var frameWidth = Unit.FromCentimeter(RightFrameWidthCm);
            var rightMargin = Unit.FromCentimeter(RightMarginCm);

            rightFrame.Top = $"{RightFrameTopCm}cm"; // Set vertical position from top of page
            rightFrame.Left = (pageWidth - frameWidth - rightMargin).ToString(); // Set horizontal position

            // Add a paragraph for the company name, styled bold and aligned to the right
            var companyParagraph = rightFrame.AddParagraph(headerModel.CompanyName);
            companyParagraph.Format.Font.Bold = true;
            companyParagraph.Format.Font.Size = CompanyFontSize;
            companyParagraph.Format.Alignment = ParagraphAlignment.Right;
            companyParagraph.Format.SpaceAfter = RightParagraphSpaceAfterCm;

            // Add a paragraph for the product name, right aligned with normal font weight
            var inspectorParagraph = rightFrame.AddParagraph(headerModel.ProductName);
            inspectorParagraph.Format.Font.Size = InspectorFontSize;
            inspectorParagraph.Format.Alignment = ParagraphAlignment.Right;
            inspectorParagraph.Format.SpaceAfter = RightParagraphSpaceAfterCm;

            // Add a paragraph for the data date range text, right aligned and with specified font size
            var dateRangeParagraph = rightFrame.AddParagraph($"Data from {headerModel.DateRange}");
            dateRangeParagraph.Format.Font.Size = DateRangeFontSize;
            dateRangeParagraph.Format.Alignment = ParagraphAlignment.Right;
        }

        // Creates the left header frame containing the logo, title, and generated date text
        private static void CreateLeftHeader(Section section, PdfHeaderModel headerModel)
        {
            // Add a text frame to the primary header of the section with fixed width and height
            var leftFrame = section.Headers.Primary.AddTextFrame();
            leftFrame.Width = LeftFrameWidth;
            leftFrame.Height = LeftFrameHeight;

            // Position the frame relative to the entire page (not content area)
            leftFrame.RelativeVertical = RelativeVertical.Page;
            leftFrame.RelativeHorizontal = RelativeHorizontal.Page;

            // Set top and left offsets of the frame from the page edges
            leftFrame.Top = LeftFrameTop;
            leftFrame.Left = LeftFrameLeft;

            // Construct the full path for the logo image file
            string logoPath = Path.Combine(Directory.GetCurrentDirectory(), headerModel.LogoPath);

            // Check if the logo file exists at the given path
            if (File.Exists(logoPath))
            {
                // Add the logo image to the left frame and set its width, maintaining aspect ratio
                var logo = leftFrame.AddImage(logoPath);
                logo.Width = LogoWidth;
                logo.LockAspectRatio = true;
            }
            else
            {
                // If the logo file is missing, display placeholder text instead
                leftFrame.AddParagraph("Logo Not Found");
            }

            // Add the report title as a bold paragraph with font size and spacing specified
            var titleParagraph = leftFrame.AddParagraph(headerModel.Title);
            titleParagraph.Format.Font.Size = TitleFontSize;
            titleParagraph.Format.Font.Bold = true;
            titleParagraph.Format.SpaceBefore = TitleSpaceBefore;
            titleParagraph.Format.SpaceAfter = TitleSpaceAfter;
            titleParagraph.Format.Alignment = ParagraphAlignment.Left;

            // Add the generated date paragraph below the title, left aligned and with smaller font size
            var genLine = leftFrame.AddParagraph();
            genLine.Format.Font.Size = GeneratedDateFontSize;
            genLine.Format.Alignment = ParagraphAlignment.Left;

            // Add a bold prefix "Generated on " followed by the actual generated date text
            var boldText = genLine.AddFormattedText("Generated on ", TextFormat.Bold);
            genLine.AddText(headerModel.GeneratedDate);
        }
    }
}
