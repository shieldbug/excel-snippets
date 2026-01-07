using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace Excel.Snippets
{
    public static partial class Demo
    {
        public static string CreateAutofitColumnSheet()
        {
            const string filePath = "AutofitColumnDemo.xlsx";

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

                Sheet sheet = new()
                {
                    Id = doc.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };

                sheets.Append(sheet);

                // Get SheetData
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Create a row
                Row row1 = new() { RowIndex = 1 };

                sheetData.Append(row1);

                row1.Append(

                    new Cell()
                    {
                        CellReference = "A1",
                        DataType = CellValues.String,
                        CellValue = new CellValue("Short")
                    }
                );

                Row row2 = new() { RowIndex = 2 };

                sheetData.Append(row2);

                row2.Append(

                    new Cell()
                    {
                        CellReference = "A2",
                        DataType = CellValues.String,
                        CellValue = new CellValue("Medium length")
                    }
                );

                Row row3 = new() { RowIndex = 3 };

                sheetData.Append(row3);

                row3.Append(

                    new Cell()
                    {
                        CellReference = "A3",
                        DataType = CellValues.String,
                        CellValue = new CellValue("This is the longest string in the column")
                    }
                );

                List<string?> ColumnValues =
                [
                    GetCell(sheetData, "A", 1).CellValue?.InnerText,
                    GetCell(sheetData, "A", 2).CellValue?.InnerText,
                    GetCell(sheetData, "A", 3).CellValue?.InnerText,
                ];

                string longest = ColumnValues.OrderByDescending(v => v != null ? v.Length : 0).First()!;

                double Width = CalculateColumnWidth(longest);

                SetColumnWidth(worksheetPart, 1, Width); // Column A = 1

                worksheetPart.Worksheet.Save();

                workbookPart.Workbook.Save();
            }

            Console.WriteLine("Excel file created: " + filePath);

            return filePath;
        }

        private static Cell GetCell(SheetData sheetData, string columnName, uint rowIndex)
        {
            // Find the row
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);

            if (row == null)
                return null;

            string cellReference = columnName + rowIndex;

            // Find the cell inside the row
            return row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellReference);
        }

        private static double CalculateColumnWidth(string longest)
        {
            /* Approximates Excel column width based on the longest string. Excel measures width
             * in "number of characters of the default font. */

            if (string.IsNullOrEmpty(longest))
                return 8.43; // Excel default width

            // Approximation for Calibri 11
            return longest.Length * 0.9 + 2;
        }

        private static void SetColumnWidth(WorksheetPart worksheetPart, uint columnIndex, double width)
        {
            var worksheet = worksheetPart.Worksheet;

            var columns = worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = new Columns();

                worksheet.InsertAt(columns, 0);
            }

            // Try to find an existing column definition
            var column = columns.Elements<Column>().FirstOrDefault(c => c.Min == columnIndex && c.Max == columnIndex);

            if (column == null)
            {
                column = new Column() { Min = columnIndex, Max = columnIndex };
                columns.Append(column);
            }

            column.Width = width;

            column.CustomWidth = true;

            worksheet.Save();
        }
    }
}