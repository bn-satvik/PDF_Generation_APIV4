using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;

namespace Proj.Utils
{
    public static class PdfHeaderLayout
    {
        // ===== Constants for layout and styling =====
        private const string TopMarginCm = "5cm";

        // Right header frame dimensions and layout
        private const string RightFrameWidth = "7cm";
        private const string RightFrameHeight = "4cm";
        private const double RightFrameTopCm = 2.1;
        private const double RightFrameWidthCm = 6;
        private const double RightMarginCm = 2.5;
        // Font sizes for right header
        private const int CompanyFontSize = 14;
        private const int InspectorFontSize = 12;
        private const int DateRangeFontSize = 12;
        private const string RightParagraphSpaceAfterCm = "0.45cm";
        // Left header frame dimensions and layout
        private const string LeftFrameWidth = "10cm";
        private const string LeftFrameHeight = "4cm";
        private const string LeftFrameTop = "1.5cm";
        private const string LeftFrameLeft = "1.5cm";
        // Logo and title styling
        private const string LogoWidth = "4cm";
        private const int TitleFontSize = 16;
        private const string TitleSpaceBefore = "0.3cm";
        private const string TitleSpaceAfter = "0.3cm";
        private const int GeneratedDateFontSize = 12;

        public static void BuildHeader(Section section, PdfHeaderModel headerModel)
        {
            section.PageSetup.TopMargin = TopMarginCm;

            // Create left and right header sections
            CreateLeftHeader(section, headerModel);
            CreateRightHeader(section, headerModel);
        }

        private static void CreateRightHeader(Section section, PdfHeaderModel headerModel)
        {
            var rightFrame = section.Headers.Primary.AddTextFrame();
            rightFrame.Width = RightFrameWidth;
            rightFrame.Height = RightFrameHeight;

            // Align frame position relative to the page
            rightFrame.RelativeVertical = RelativeVertical.Page;
            rightFrame.RelativeHorizontal = RelativeHorizontal.Page;

            // Calculate position from right edge of the page
            var pageWidth = section.PageSetup.PageWidth;
            var frameWidth = Unit.FromCentimeter(RightFrameWidthCm);
            var rightMargin = Unit.FromCentimeter(RightMarginCm);

            rightFrame.Top = $"{RightFrameTopCm}cm";
            rightFrame.Left = (pageWidth - frameWidth - rightMargin).ToString();

            // Add company name (bold, right aligned)
            var companyParagraph = rightFrame.AddParagraph(headerModel.CompanyName);
            companyParagraph.Format.Font.Bold = true;
            companyParagraph.Format.Font.Size = CompanyFontSize;
            companyParagraph.Format.Alignment = ParagraphAlignment.Right;
            companyParagraph.Format.SpaceAfter = RightParagraphSpaceAfterCm;

            // Add product name (right aligned)
            var inspectorParagraph = rightFrame.AddParagraph(headerModel.ProuductName);
            inspectorParagraph.Format.Font.Size = InspectorFontSize;
            inspectorParagraph.Format.Alignment = ParagraphAlignment.Right;
            inspectorParagraph.Format.SpaceAfter = RightParagraphSpaceAfterCm;

            // Add date range (e.g. "Data from Jan–Feb 2024")
            var dateRangeParagraph = rightFrame.AddParagraph($"Data from {headerModel.DateRange}");
            dateRangeParagraph.Format.Font.Size = DateRangeFontSize;
            dateRangeParagraph.Format.Alignment = ParagraphAlignment.Right;
        }

        private static void CreateLeftHeader(Section section, PdfHeaderModel headerModel)
        {
            var leftFrame = section.Headers.Primary.AddTextFrame();
            leftFrame.Width = LeftFrameWidth;
            leftFrame.Height = LeftFrameHeight;

            // Set frame position on the page
            leftFrame.RelativeVertical = RelativeVertical.Page;
            leftFrame.RelativeHorizontal = RelativeHorizontal.Page;
            leftFrame.Top = LeftFrameTop;
            leftFrame.Left = LeftFrameLeft;

            // Add company logo if it exists
            string logoPath = Path.Combine(Directory.GetCurrentDirectory(), headerModel.LogoPath);
            if (File.Exists(logoPath))
            {
                var logo = leftFrame.AddImage(logoPath);
                logo.Width = LogoWidth;
                logo.LockAspectRatio = true;
            }
            else
            {
                // Fallback if logo file not found
                leftFrame.AddParagraph("Logo Not Found");
            }

            // Add report title
            var titleParagraph = leftFrame.AddParagraph(headerModel.Title);
            titleParagraph.Format.Font.Size = TitleFontSize;
            titleParagraph.Format.Font.Bold = true;
            titleParagraph.Format.SpaceBefore = TitleSpaceBefore;
            titleParagraph.Format.SpaceAfter = TitleSpaceAfter;
            titleParagraph.Format.Alignment = ParagraphAlignment.Left;

            // Add generated date
            var genLine = leftFrame.AddParagraph();
            genLine.Format.Font.Size = GeneratedDateFontSize;
            genLine.Format.Alignment = ParagraphAlignment.Left;

            // "Generated on May 13, 2025"
            var boldText = genLine.AddFormattedText("Generated on ", TextFormat.Bold);
            genLine.AddText(headerModel.GeneratedDate);
        }
    }
}
