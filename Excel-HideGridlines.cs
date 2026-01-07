using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Snippets
{
    public static partial class Demo
    {
        public static string CreateWorkbookWithoutGridlines()
        {
            string filePath = "WorkbookHideWorksheet.xlsx";

            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // Create the main parts
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                // Add SheetViews with gridlines hidden
                var sheetViews = new SheetViews();
                var sheetView = new SheetView()
                {
                    WorkbookViewId = 0U,
                    ShowGridLines = false
                };
                sheetViews.Append(sheetView);

                worksheetPart.Worksheet.Append(sheetViews);
                worksheetPart.Worksheet.Append(new SheetData()); // empty sheet

                worksheetPart.Worksheet.Save();

                // Add Sheets collection
                var sheets = new Sheets();
                var sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);

                workbookPart.Workbook.Append(sheets);
                workbookPart.Workbook.Save();
            }

            Console.WriteLine("Workbook created with gridlines hidden.");

            return filePath;
        }
    }
}
