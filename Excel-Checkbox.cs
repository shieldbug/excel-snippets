using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel.Snippets
{
    public static partial class Demo
    {
        public static string CreateWorkbookWithCheckbox()
        {
            const string filePath = "CheckboxRawXml.xlsx";

            using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // --- Workbook + worksheet setup ---
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);

                // Ensure at least 2 rows so B2 exists
                var row2 = new Row { RowIndex = 2 };
                sheetData.Append(new Row { RowIndex = 1 });
                sheetData.Append(row2);

                // Optional: put some text in B2 so you can see the anchor cell
                var cellB2 = new Cell
                {
                    CellReference = "B2",
                    DataType = CellValues.String,
                    CellValue = new CellValue("Checkbox cell")
                };
                row2.Append(cellB2);

                // --- VML drawing part for form control checkbox ---
                var vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>();

                // Raw VML markup for a single checkbox anchored around B2
                // This is similar to what Excel itself generates.
                const string vmlXml = @"
<xml
    xmlns:v=""urn:schemas-microsoft-com:vml""
    xmlns:o=""urn:schemas-microsoft-com:office:office""
    xmlns:x=""urn:schemas-microsoft-com:office:excel"">
  <o:shapelayout v:ext=""edit"">
    <o:idmap v:ext=""edit"" data=""1""/>
  </o:shapelayout>
  <v:shapetype id=""_x0000_t201""
      coordsize=""21600,21600""
      o:spt=""201""
      path=""m,l,21600r21600,l21600,xe"">
    <v:stroke joinstyle=""miter""/>
    <v:path gradientshapeok=""t"" o:connecttype=""rect""/>
  </v:shapetype>

  <v:shape id=""_x0000_s1025""
      type=""#_x0000_t201""
      style=""position:absolute;margin-left:70pt;margin-top:15pt;width:80pt;height:15pt;z-index:1""
      fillcolor=""window [65]""
      strokecolor=""windowText [64]"">
    <v:stroke/>
    <v:fill/>
    <x:ClientData ObjectType=""Checkbox"">
      <x:MoveWithCells/>
      <x:SizeWithCells/>
      <!-- Anchor: col1,row1,col2,row2 style; this roughly targets B2 -->
      <x:Anchor>1, 0, 1, 0, 2, 0, 2, 15</x:Anchor>
      <x:Locked>False</x:Locked>
      <x:DefaultSize>False</x:DefaultSize>
      <x:VAlign>Center</x:VAlign>
      <x:Caption>Check me</x:Caption>
      <x:Value>0</x:Value>
      <x:Checked>False</x:Checked>
      <x:Mixed>False</x:Mixed>
      <x:PrintObject>True</x:PrintObject>
    </x:ClientData>
  </v:shape>
</xml>";

                using (var stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write))
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write(vmlXml);
                }

                // --- Link worksheet to the VML drawing (legacy drawing) ---
                var legacyDrawing = new LegacyDrawing
                {
                    Id = worksheetPart.GetIdOfPart(vmlPart)
                };
                worksheetPart.Worksheet.Append(legacyDrawing);

                workbookPart.Workbook.Save();
            }

            Console.WriteLine("Excel file created: " + filePath);
            return filePath;
        }
    }
}