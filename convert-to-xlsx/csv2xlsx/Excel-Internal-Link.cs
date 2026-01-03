using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace csv2xlsx
{
    public static partial class Demo
    {
        public static void CreateInternallinkSheet()
        {
            const string filePath = "InternalLinkDemo.xlsx";

            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart wb = doc.AddWorkbookPart();
                wb.Workbook = new Workbook();

                // Create Sheet1
                WorksheetPart ws1 = wb.AddNewPart<WorksheetPart>();
                ws1.Worksheet = new Worksheet(new SheetData());

                // Create Sheet2
                WorksheetPart ws2 = wb.AddNewPart<WorksheetPart>();
                ws2.Worksheet = new Worksheet(new SheetData());

                // Register sheets
                Sheets sheets = wb.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet { Id = wb.GetIdOfPart(ws1), SheetId = 1, Name = "Sheet1" });
                sheets.Append(new Sheet { Id = wb.GetIdOfPart(ws2), SheetId = 2, Name = "Sheet2" });

                // Add text to Sheet2!A1
                var sd2 = ws2.Worksheet.GetFirstChild<SheetData>();
                var row2 = new Row() { RowIndex = 1 };
                sd2.Append(row2);
                row2.Append(new Cell
                {
                    CellReference = "A1",
                    DataType = CellValues.String,
                    CellValue = new CellValue("Target cell on Sheet2")
                });

                // Add hyperlink cell on Sheet1!A1
                var sd1 = ws1.Worksheet.GetFirstChild<SheetData>();
                var row1 = new Row() { RowIndex = 1 };
                sd1.Append(row1);

                row1.Append(new Cell
                {
                    CellReference = "A1",
                    DataType = CellValues.String,
                    CellValue = new CellValue("Go to Sheet2!A1")
                });

                // Add <hyperlinks> collection
                Hyperlinks links = ws1.Worksheet.GetFirstChild<Hyperlinks>();
                if (links == null)
                {
                    links = new Hyperlinks();
                    ws1.Worksheet.InsertAfter(links, sd1);
                }

                // INTERNAL hyperlink (no relationship needed)
                links.Append(new Hyperlink
                {
                    Reference = "A1",
                    Location = "Sheet2!A1"   // <-- internal target
                });

                ws1.Worksheet.Save();
                ws2.Worksheet.Save();
                wb.Workbook.Save();
            }

            Console.WriteLine("Excel file created: " + filePath);
        }
    }
}