using MigraDoc.DocumentObjectModel;

namespace Proj.Utils
{
    public static class PdfFooterLayout
    {
        // ===== Constants =====
        private const int FooterFontSize = 12;
        private const ParagraphAlignment FooterAlignment = ParagraphAlignment.Right;
        private const string PageTextPrefix = "Page ";
     
        private const string PageTextConnector = " of ";

        // Builds the footer content for the PDF 
        public static void BuildFooter(PdfFooterModel footerModel, Section section)
        {
            // Get the footer section
            var footer = section.Footers.Primary;

            // Create a new paragraph in the footer
            Paragraph paragraph = footer.AddParagraph();
            paragraph.Format.Font.Size = FooterFontSize;
            paragraph.Format.Alignment = FooterAlignment;

            // Add page number info if enabled
            if (footerModel.ShowPageNumbers)
            {
                if (!string.IsNullOrEmpty(footerModel.RightText))
                {
                    paragraph.AddText(footerModel.RightText);
                  
                }

                paragraph.AddText(PageTextPrefix);
                paragraph.AddPageField();        // Adds current page number
                paragraph.AddText(PageTextConnector);
                paragraph.AddNumPagesField();    // Adds total page count
            }

            // Re-add right text if combined with other info
            if (!string.IsNullOrEmpty(footerModel.RightText))
            {
                paragraph.AddText(footerModel.RightText);
            }
        }
    }
}
