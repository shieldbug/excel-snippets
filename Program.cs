using DocumentFormat.OpenXml.Packaging;
using System.Runtime.CompilerServices;

namespace Excel.Snippets
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // await Convert.WriteOpenXmlAsync(args[0]);

            List<string> ExcelFiles = [];

            ExcelFiles.Add(Demo.CreateHyperlinkSheet());
            ExcelFiles.Add(Demo.CreateHyperlinkStyledSheet());
            ExcelFiles.Add(Demo.CreateInternalLinkSheet());
            ExcelFiles.Add(Demo.CreateInternalLinkStyledSheet());
            ExcelFiles.Add(Demo.CreateAutofitColumnSheet());
            ExcelFiles.Add(Demo.CreateWorkbookWithoutGridlines());
            ExcelFiles.Add(Demo.CreateWorkbookWithCheckbox());

            foreach (string ExcelFile in ExcelFiles)
            {
                using (var doc = SpreadsheetDocument.Open(ExcelFile, false))
                {
                    var errors = Demo.Validate(doc);

                    foreach (var error in errors)
                    {
                        Console.WriteLine($"Workbook {ExcelFile}: {error}");
                    }

                    if (!errors.Any())
                    {
                        Console.WriteLine($"Workbook {ExcelFile} is structurally valid.");
                    }
                }
            }
        }
    }
}
