public static class PdfConstants
{

     // === Image and Page Sizing Constants ===

        public const double Dpi = 96;
        // Standard screen resolution used for converting pixels to inches for image scaling
        // → Interdependent with InchesToCentimeters and TargetImageWidthCm for image size calculation

        public const double InchesToCentimeters = 2.54;
        // Conversion factor: 1 inch = 2.54 cm; used to convert image size from inches to cm
        // → Used with Dpi and TargetImageWidthCm for image scaling calculations

        public const double TargetImageWidthCm = 15.0;
        // Desired image width (in cm) in the final PDF. This is used to compute the height proportionally.
        // → Interdependent: Works with Dpi and InchesToCentimeters for scaling
        // → Used to calculate page width and height in CalculateImagePageSize

        public const double MarginCm = 1.0;
        // Margin around the image (left, right, top, bottom)
        // → Interdependent with page size calculations (MinImagePageWidthCm/HeightCm + HeaderFooterPaddingCm)

        public const double MinImagePageWidthCm = 16.0;
        // Minimum width of the page to ensure header's left and right fields do not overlap
        // → Related to header layout design and margin settings

        public const double MinImagePageHeightCm = 16.0;
        // Minimum height of the image page to ensure enough room for image + margins + headers/footers
        // → Interdependent with MarginCm and HeaderFooterPaddingCm to ensure proper spacing


        // === Text and Column Layout ===

        public const double CharWidthCm = 0.23;
        // Approximate width of a single character in cm for font size 10
        // → Interdependent: Directly tied to DefaultFontSize (font size affects character width)
        // → Used to calculate dynamic column widths and overall table width

        public const double MinColWidthCm = 2.0;
        // Minimum width of a column to avoid cramped data presentation
        // → Used with CharWidthCm and max column widths to clamp calculated widths

        public const int referenceLenMax = 10;
        // Maximum allowed length for reference numbers (affects column size for reference field)
        // → Used in CalculateColumnWidth to ensure minimum column size for reference fields


        // === Word and Page Spacing ===

        public const int WordSpace = 2;
        // Number of characters to use as spacing between words in columns
        // → Affects minimum width calculation of columns to accommodate spacing

        public const double MaxPageWidthCm = 75.0;
        // Maximum width allowed for a page (ensures it's printable on standard A4 paper/printers)
        // → Limits final page width regardless of content width to maintain printability

        public const double DefaultPageWidthCm = 21.0;
        // Default page width, typically equivalent to standard A4 width
        // → Used as fallback if calculated page width is smaller or larger
        // → Interdependent with MaxPageWidthCm to clamp final width

        public const double PageHeightCm = 34.0;
        // Height of the page (used for both A4 and larger custom sizes)
        // → Interdependent with HeaderFooterPaddingCm, TableTopMarginCm, and TableBottomMarginCm for vertical layout


        // === Table Layout and Margins ===

        public const double TableTopMarginCm = 2.0;
        // Space from the top of the page before the table starts
        // → Interdependent with PageHeightCm and HeaderFooterPaddingCm for vertical spacing

        public const double TableBottomMarginCm = 2.5;
        // Space at the bottom of the page after the table ends
        // → Interdependent with PageHeightCm and HeaderFooterPaddingCm for vertical spacing

        public const double HeaderFooterPaddingCm = 6.0;
        // Space reserved for header and footer (added outside image/table zone)
        // → Interdependent: Reduces usable vertical space for image/table in CalculateImagePageSize and table layout

        public const double TablePaddingCm = 3.0;
        // Extra padding inside table layout, between cell borders and content
        // → Interdependent with calculated tableWidth and pageWidth for overall layout


        // === Font Sizes ===

        public const int DefaultFontSize = 10;
        // Main font size used throughout the document
        // → Interdependent with CharWidthCm (character width depends on font size)
        // → Affects readability and column width calculation

        public const int HeaderFontSize = 12;
        // Font size specifically used in headers (usually slightly larger for emphasis)
        // → Visually distinct from DefaultFontSize

        public const int ColumnFontSize = 9;
        // Font size for table column text (slightly smaller to fit more data)
        // → Used for table columns to improve fitting large data sets


        // === Spacing Between Sections ===

        public const string HeaderSpaceBefore = "0.3cm";
        // CSS-style spacing before the header content begins
        // → Works with HeaderSpaceAfter to separate header visually

        public const string HeaderSpaceAfter = "0.3cm";
        // CSS-style spacing after the header content ends
        // → Works with HeaderSpaceBefore to separate header visually

        public const string DataSpaceBefore = "0.15cm";
        // Space before each data block or row
        // → Works with DataSpaceAfter to add vertical padding between rows

        public const string DataSpaceAfter = "0.15cm";
        // Space after each data block or row
        // → Works with DataSpaceBefore to add vertical padding between rows


        // === Cell Appearance ===

        public const double CellBorderWidth = 0.5;
        // Thickness of cell borders in tables (in points or cm depending on rendering system)
        // → Consistent border width for table cells for readability


        // === Word Wrapping ===

        public const int SoftBreakInterval = 20;
        // Number of characters after which soft breaks (line breaks) are inserted into long strings
        // → Improves text wrapping in table cells for better layout


        // Constants to avoid magic numbers in column width mapping
        public const double MaxColWidthSmall = 11.0;
        // Max column width for 1-4 columns
        // → Used in GetDynamicMaxColWidthCm to clamp column width based on column count

        public const double MaxColWidthMedium = 9.0;
        // Max column width for 5-12 columns
        // → Used in GetDynamicMaxColWidthCm to clamp column width based on column count

        public const double MaxColWidthLarge = 5.0;
        // Max column width for 13-15 columns
        // → Used in GetDynamicMaxColWidthCm to clamp column width based on column count

        public const double MaxColWidthExtraLarge = 3.0;
        // Max column width for 16 or more columns
        // → Used in GetDynamicMaxColWidthCm to clamp column width based on column count
}
