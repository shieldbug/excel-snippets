using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Snippets
{
    public static partial class Demo
    {
        public static IEnumerable<string> Validate(SpreadsheetDocument doc)
        {
            var errors = new List<string>();

            // 1. Validate workbook structure
            errors.AddRange(ValidateWorkbookPart(doc.WorkbookPart));

            // 2. Validate styles
            if (doc.WorkbookPart.WorkbookStylesPart != null)
                errors.AddRange(ValidateStyles(doc.WorkbookPart.WorkbookStylesPart));

            // 3. Validate shared strings
            if (doc.WorkbookPart.SharedStringTablePart != null)
                errors.AddRange(ValidateSharedStrings(doc.WorkbookPart.SharedStringTablePart));

            // 4. Validate each worksheet
            foreach (var sheetPart in doc.WorkbookPart.WorksheetParts)
                errors.AddRange(ValidateWorksheet(sheetPart));

            return errors;
        }

        // ------------------------------------------------------------
        // WORKBOOK PART VALIDATION
        // ------------------------------------------------------------
        private static IEnumerable<string> ValidateWorkbookPart(WorkbookPart workbookPart)
        {
            var errors = new List<string>();

            if (workbookPart.Workbook == null)
                errors.Add("WorkbookPart is missing <workbook> root element.");

            if (!workbookPart.Workbook.Sheets!.Any())
                errors.Add("Workbook contains no <sheet> elements.");

            // Check sheet references
            foreach (var sheet in workbookPart.Workbook.Sheets.OfType<Sheet>())
            {
                if (sheet.Id == null)
                {
                    errors.Add($"Sheet '{sheet.Name}' has no relationship Id.");
                    continue;
                }

                if (!workbookPart.Parts.Any(p => p.RelationshipId == sheet.Id))
                    errors.Add($"Sheet '{sheet.Name}' references missing part: {sheet.Id}");
            }

            return errors;
        }

        // ------------------------------------------------------------
        // WORKSHEET VALIDATION (ORDER + RELATIONSHIPS)
        // ------------------------------------------------------------
        private static IEnumerable<string> ValidateWorksheet(WorksheetPart sheetPart)
        {
            var errors = new List<string>();

            if (sheetPart.Worksheet == null)
            {
                errors.Add("WorksheetPart missing <worksheet> root.");
                return errors;
            }

            // Use your existing order validator
            errors.AddRange(WorksheetOrderValidator.ValidateOrder(sheetPart.Worksheet));

            // Check for missing SheetData
            if (sheetPart.Worksheet.GetFirstChild<SheetData>() == null)
                errors.Add("Worksheet missing <sheetData> element.");

            // Check drawing relationships
            foreach (var drawing in sheetPart.Worksheet.Descendants<Drawing>())
            {
                if (!sheetPart.Parts.Any(p => p.RelationshipId == drawing.Id))
                    errors.Add($"Worksheet has <drawing> referencing missing part: {drawing.Id}");
            }

            return errors;
        }

        // ------------------------------------------------------------
        // STYLE VALIDATION (ORDER + COUNTS)
        // ------------------------------------------------------------
        private static IEnumerable<string> ValidateStyles(WorkbookStylesPart stylesPart)
        {
            var errors = new List<string>();
            var ss = stylesPart.Stylesheet;

            if (ss == null)
            {
                errors.Add("Stylesheet missing <styleSheet> root.");
                return errors;
            }

            // Required elements
            if (ss.CellStyleFormats == null)
                errors.Add("<cellStyleFormats> missing from stylesheet.");

            if (ss.CellFormats == null)
                errors.Add("<cellFormats> missing from stylesheet.");

            // Order check
            var expectedOrder = new List<Type>
        {
            typeof(NumberingFormats),
            typeof(Fonts),
            typeof(Fills),
            typeof(Borders),
            typeof(CellStyleFormats),
            typeof(CellFormats),
            typeof(CellStyles),
            typeof(DifferentialFormats),
            typeof(TableStyles),
            typeof(Colors),
            typeof(ExtensionList)
        };

            int lastIndex = -1;
            foreach (var child in ss.ChildElements)
            {
                int index = expectedOrder.IndexOf(child.GetType());
                if (index == -1)
                {
                    errors.Add($"Unexpected element in stylesheet: {child.GetType().Name}");
                    continue;
                }

                if (index < lastIndex)
                {
                    errors.Add($"Stylesheet element {child.GetType().Name} is out of order.");
                }

                lastIndex = Math.Max(lastIndex, index);
            }

            return errors;
        }

        // ------------------------------------------------------------
        // SHARED STRINGS VALIDATION
        // ------------------------------------------------------------
        private static IEnumerable<string> ValidateSharedStrings(SharedStringTablePart sstPart)
        {
            var errors = new List<string>();
            var table = sstPart.SharedStringTable;

            if (table == null)
            {
                errors.Add("SharedStringTablePart missing <sst> root.");
                return errors;
            }

            // Check for empty entries
            int index = 0;
            foreach (var item in table.Elements<SharedStringItem>())
            {
                if (!item.Any())
                    errors.Add($"Shared string at index {index} is empty.");

                index++;
            }

            return errors;
        }
    }

    internal static class WorksheetOrderValidator
    {
        // Canonical order from the ECMA/ISO spec
        private static readonly List<Type> WorksheetOrderField =
        [
        typeof(SheetProperties),
        typeof(SheetDimension),
        typeof(SheetViews),
        typeof(SheetFormatProperties),
        typeof(Columns),
        typeof(SheetData),
        typeof(SheetCalculationProperties),
        typeof(SheetProtection),
        typeof(ProtectedRanges),
        typeof(Scenarios),
        typeof(AutoFilter),
        typeof(SortState),
        typeof(DataConsolidate),
        typeof(CustomSheetViews),
        typeof(MergeCells),
        typeof(PhoneticProperties),
        typeof(ConditionalFormatting),
        typeof(DataValidations),
        typeof(Hyperlinks),
        typeof(PrintOptions),
        typeof(PageMargins),
        typeof(PageSetup),
        typeof(HeaderFooter),
        typeof(RowBreaks),
        typeof(ColumnBreaks),
        typeof(CustomProperties),
        typeof(CellWatches),
        typeof(IgnoredErrors),
        typeof(Drawing),
        typeof(LegacyDrawing),
        typeof(LegacyDrawingHeaderFooter),
        typeof(Picture),
        typeof(OleObjects),
        typeof(Controls),
        typeof(WebPublishItems),
        typeof(TableParts),
        typeof(ExtensionList)
        ];

        public static IEnumerable<string> ValidateOrder(Worksheet worksheet)
        {
            var errors = new List<string>();

            var children = worksheet.ChildElements;
            int lastIndex = -1;

            foreach (var child in children)
            {
                var type = child.GetType();
                int index = WorksheetOrderField.IndexOf(type);

                if (index == -1)
                {
                    errors.Add($"Unknown or unexpected element: {type.Name}");
                    continue;
                }

                if (index < lastIndex)
                {
                    errors.Add(
                        $"Element {type.Name} is out of order. " +
                        $"Expected after {WorksheetOrderField[lastIndex].Name}, " +
                        $"but appears earlier."
                    );
                }

                lastIndex = Math.Max(lastIndex, index);
            }

            return errors;
        }
    }
}
