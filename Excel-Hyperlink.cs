using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel.Snippets
{
    public static partial class Demo
    {
        public static string CreateHyperlinkSheet()
        {
            const string filePath = "HyperlinkDemo.xlsx";

            // Create the document
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // Add workbook + worksheet
                WorkbookPart workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets collection
                Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = doc.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);

                // Get SheetData
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Create a row
                Row row = new Row() { RowIndex = 1 };
                sheetData.Append(row);

                // Normal text cell in A1
                row.Append(
                    new Cell()
                    {
                        CellReference = "A1",
                        DataType = CellValues.String,
                        CellValue = new CellValue("Normal text")
                    }
                );

                // Hyperlink cell in B1
                string hyperlinkText = "OpenAI Website";
                string hyperlinkUrl = "https://www.openai.com";

                // Create the cell
                Cell hyperlinkCell = new Cell()
                {
                    CellReference = "B1",
                    DataType = CellValues.String,
                    CellValue = new CellValue(hyperlinkText)
                };
                row.Append(hyperlinkCell);

                // Add hyperlink relationship
                HyperlinkRelationship rel = worksheetPart.AddHyperlinkRelationship(
                    new Uri(hyperlinkUrl, UriKind.Absolute),
                    true
                );

                // Add Hyperlinks collection if missing
                Hyperlinks hyperlinks = worksheetPart.Worksheet.GetFirstChild<Hyperlinks>();
                if (hyperlinks == null)
                {
                    hyperlinks = new Hyperlinks();
                    worksheetPart.Worksheet.InsertAfter(hyperlinks, sheetData);
                }

                // Add hyperlink entry
                hyperlinks.Append(
                    new Hyperlink()
                    {
                        Reference = "B1",
                        Id = rel.Id
                    }
                );

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            Console.WriteLine("Excel file created: " + filePath);

            return filePath;
        }
    }
}