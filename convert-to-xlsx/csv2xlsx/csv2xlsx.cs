using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace csv2xlsx
{
    public static partial class Convert
    {
        public static async Task<bool> WriteOpenXmlAsync(string CsvDataPath, string? WorksheetName = null)
        {
            return await Task.Run(async () =>
            {
                if (Path.GetExtension(CsvDataPath).Equals(".csv", StringComparison.OrdinalIgnoreCase) && File.Exists(CsvDataPath))
                {
                    if (string.IsNullOrEmpty(WorksheetName))
                    {
                        /* A work sheet name is required. We use the filename if there is no user defined name available. */

                        WorksheetName = Path.GetFileNameWithoutExtension(CsvDataPath);
                    }

                    string[] CsvData = await File.ReadAllLinesAsync(CsvDataPath);

                    /* Presumably, we have write access. */

                    using SpreadsheetDocument Package = SpreadsheetDocument.Create(Path.ChangeExtension(CsvDataPath, "xlsx"), SpreadsheetDocumentType.Workbook, true);

                    Package.AddWorkbookPart();

                    if (Package != null)
                    {
                        if (Package.WorkbookPart != null)
                        {
                            Package.WorkbookPart.Workbook = new();

                            if (Package.WorkbookPart.Workbook != null)
                            {
                                Package.WorkbookPart.AddNewPart<WorksheetPart>();

                                SheetData XlsxSheetData = new();

                                foreach (var CsvRow in CsvData)
                                {
                                    Row xlRow = new();

                                    foreach (var CsvCol in CsvRow.Split(','))
                                    {
                                        /* We treat int as decimal (floating point values) */

                                        if (decimal.TryParse(CsvCol, out decimal ColDecimal))
                                        {
                                            Cell xlCell = new()
                                            {
                                                DataType = CellValues.Number,

                                                CellValue = new CellValue(ColDecimal)
                                            };

                                            xlRow.Append(xlCell);
                                        }
                                        else
                                        {
                                            Cell xlCell = new(new InlineString(new Text(CsvCol)))
                                            {
                                                DataType = CellValues.InlineString
                                            };

                                            xlRow.Append(xlCell);
                                        }
                                    }

                                    XlsxSheetData.Append(xlRow);
                                }

                                Package.WorkbookPart.WorksheetParts.First().Worksheet = new Worksheet(XlsxSheetData);

                                Package.WorkbookPart.WorksheetParts.First().Worksheet.Save();

                                Package.WorkbookPart.Workbook.AppendChild(new Sheets());

                                Package.WorkbookPart.Workbook.GetFirstChild<Sheets>()!.AppendChild(new Sheet()
                                {
                                    Id = Package.WorkbookPart.GetIdOfPart(Package.WorkbookPart.WorksheetParts.First()),

                                    SheetId = 1,

                                    Name = WorksheetName

                                });

                                Package.WorkbookPart.Workbook.Save();

                                return true;
                            }
                        }
                    }

                    return false;
                }
                else
                {
                    /* CSV file does not exist. */

                    return false;
                }
            });
        }
    }
}
