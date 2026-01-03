using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace csv2xlsx
{
    public static partial class Demo
    {
        public static void CreateHyperlinkStyledSheet()
        {
            const string filePath = "HyperlinkStyled.xlsx";

            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Create stylesheet with hyperlink style
                WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = CreateStylesheet();
                stylesPart.Stylesheet.Save();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                });

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                Row row = new Row() { RowIndex = 1 };
                sheetData.Append(row);

                // Normal text
                row.Append(new Cell
                {
                    CellReference = "A1",
                    DataType = CellValues.String,
                    CellValue = new CellValue("Normal text")
                });

                // Hyperlink cell with style index = 1 (our hyperlink style)
                Cell hyperlinkCell = new Cell
                {
                    CellReference = "B1",
                    DataType = CellValues.String,
                    CellValue = new CellValue("OpenAI Website"),
                    StyleIndex = 1 // <-- apply hyperlink style
                };
                row.Append(hyperlinkCell);

                // Add hyperlink relationship
                var rel = worksheetPart.AddHyperlinkRelationship(
                    new Uri("https://www.openai.com", UriKind.Absolute),
                    true
                );

                // Add hyperlink element
                Hyperlinks hyperlinks = worksheetPart.Worksheet.GetFirstChild<Hyperlinks>();
                if (hyperlinks == null)
                {
                    hyperlinks = new Hyperlinks();
                    worksheetPart.Worksheet.InsertAfter(hyperlinks, sheetData);
                }

                hyperlinks.Append(new Hyperlink
                {
                    Reference = "B1",
                    Id = rel.Id
                });

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            Console.WriteLine("Excel file created: " + filePath);

            // ----------------------
            // Stylesheet definition
            // ----------------------
            Stylesheet CreateStylesheet()
            {
                // Fonts: index 0 = default, index 1 = hyperlink font
                Fonts fonts = new Fonts(
                    new Font(), // default
                    new Font(   // hyperlink font
                        new Color() { Rgb = "0000FF" }, // blue
                        new Underline()                 // underline
                    )
                );

                // Fills (required even if unused)
                Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }),
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 })
                );

                // Borders (required even if unused)
                Borders borders = new Borders(new Border());

                // CellFormats: index 0 = default, index 1 = hyperlink style
                CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat()  // hyperlink style
                    {
                        FontId = 1,   // use hyperlink font
                        ApplyFont = true
                    }
                );

                return new Stylesheet(fonts, fills, borders, cellFormats);
            }
        }
    }
}