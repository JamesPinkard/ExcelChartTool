using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLTools
{
    public class Table2bGenerator
    {
        private WorkbookWriter _workbookWriter;
        private WorksheetFormatter _worksheetFormatter;

        public Table2bGenerator(WorkbookWriter workbookWriter)
        {
            _workbookWriter = workbookWriter;            
        }

        public WorksheetWriter GenerateWriter(string worksheetName, CellReference cellReference)
        {
            var writer = _workbookWriter.CreateWorksheetWriter(worksheetName, cellReference);
            writer.GetWorksheetPart().Worksheet = GenerateWorksheetFirstPart();
            return writer;
        }

        public void FormatWorksheet(WorksheetPart worksheetPart)
        {
            var formatter = new WorksheetFormatter(worksheetPart);
            formatter.FormatSheet();
            
            PrintOptions printOptions1 = GetPrintOptions();
            PageMargins pageMargins1 = GetPageMargins();
            PageSetup pageSetup1 = GetPageSetup();
            HeaderFooter headerFooter1 = GetHeaderFooter();
            IgnoredErrors ignoredErrors1 = GetIgnoredErrors();
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter1 = GetLegacyHeaderFooter();

            var worksheet1 = worksheetPart.Worksheet;
            worksheet1.Append(printOptions1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(headerFooter1);
            worksheet1.Append(ignoredErrors1);
            worksheet1.Append(legacyDrawingHeaderFooter1);
        }

        private Worksheet GenerateWorksheetFirstPart()
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties1 = GetSheetProperties();
            SheetDimension sheetDimension1 = GetSheetDimension();
            SheetViews sheetViews1 = GetSheetViews();
            SheetFormatProperties sheetFormatProperties1 = GetSheetFormatProperties();
            Columns columns1 = GetColumns();
            SheetData sheetData1 = new SheetData();

            worksheet1.Append(sheetProperties1);
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            return worksheet1;
        }

        // Creates an Worksheet instance and adds its children.
        private Worksheet GenerateWorksheet()
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties1 = GetSheetProperties();
            SheetDimension sheetDimension1 = GetSheetDimension();
            SheetViews sheetViews1 = GetSheetViews();
            SheetFormatProperties sheetFormatProperties1 = GetSheetFormatProperties();
            Columns columns1 = GetColumns();
            SheetData sheetData1 = new SheetData();

            MergeCells mergeCells1 = GetMergeCells();

            PrintOptions printOptions1 = GetPrintOptions();
            PageMargins pageMargins1 = GetPageMargins();
            PageSetup pageSetup1 = GetPageSetup();
            HeaderFooter headerFooter1 = GetHeaderFooter();
            IgnoredErrors ignoredErrors1 = GetIgnoredErrors();
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter1 = GetLegacyHeaderFooter();

            worksheet1.Append(sheetProperties1);
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells1);
            worksheet1.Append(printOptions1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(headerFooter1);
            worksheet1.Append(ignoredErrors1);
            worksheet1.Append(legacyDrawingHeaderFooter1);
            return worksheet1;
        }

        private static LegacyDrawingHeaderFooter GetLegacyHeaderFooter()
        {
            return new LegacyDrawingHeaderFooter() { Id = "rId2" };
        }

        private static IgnoredErrors GetIgnoredErrors()
        {
            IgnoredErrors ignoredErrors1 = new IgnoredErrors();
            IgnoredError ignoredError1 = new IgnoredError() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "F31" }, Formula = true };

            ignoredErrors1.Append(ignoredError1);
            return ignoredErrors1;
        }

        private static HeaderFooter GetHeaderFooter()
        {
            HeaderFooter headerFooter1 = new HeaderFooter();
            OddHeader oddHeader1 = new OddHeader();
            oddHeader1.Text = "&C&\"-,Bold\"&12TABLE 2B\nSummary of Flow Meter Readings&11\n&10 &11 4&Xth&X Quarter 2015 Remediation Status Report\nMountain View Nitrate Plume Restoration Project";
            OddFooter oddFooter1 = new OddFooter();
            oddFooter1.Text = "&L&G&RPage &P of &N";

            headerFooter1.Append(oddHeader1);
            headerFooter1.Append(oddFooter1);
            return headerFooter1;
        }

        private static PageSetup GetPageSetup()
        {
            return new PageSetup() { FitToHeight = (UInt32Value)14U, Orientation = OrientationValues.Portrait, Id = "rId1" };
        }

        private static PageMargins GetPageMargins()
        {
            return new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 1.2D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
        }

        private static PrintOptions GetPrintOptions()
        {
            return new PrintOptions() { HorizontalCentered = true };
        }

        private static MergeCells GetMergeCells()
        {
            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)8U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "H100:H112" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "H3:H17" };
            MergeCell mergeCell3 = new MergeCell() { Reference = "H31:H45" };
            MergeCell mergeCell4 = new MergeCell() { Reference = "H59:H73" };
            MergeCell mergeCell5 = new MergeCell() { Reference = "H85:H99" };
            MergeCell mergeCell6 = new MergeCell() { Reference = "H18:H30" };
            MergeCell mergeCell7 = new MergeCell() { Reference = "H46:H58" };
            MergeCell mergeCell8 = new MergeCell() { Reference = "H74:H84" };

            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            mergeCells1.Append(mergeCell3);
            mergeCells1.Append(mergeCell4);
            mergeCells1.Append(mergeCell5);
            mergeCells1.Append(mergeCell6);
            mergeCells1.Append(mergeCell7);
            mergeCells1.Append(mergeCell8);
            return mergeCells1;
        }

        private static Columns GetColumns()
        {
            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 3.68359375D, Style = (UInt32Value)93U, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 12D, Style = (UInt32Value)93U, BestFit = true, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 8.89453125D, Style = (UInt32Value)93U };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 11.1015625D, Style = (UInt32Value)93U, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 16D, Style = (UInt32Value)93U, BestFit = true, CustomWidth = true };
            Column column6 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 15.68359375D, Style = (UInt32Value)93U, CustomWidth = true };
            Column column7 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 9.41796875D, Style = (UInt32Value)93U, CustomWidth = true };
            Column column8 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 10.1015625D, Style = (UInt32Value)93U, CustomWidth = true };
            Column column9 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 8.89453125D, Style = (UInt32Value)93U };
            Column column10 = new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 9.68359375D, Style = (UInt32Value)93U, CustomWidth = true };
            Column column11 = new Column() { Min = (UInt32Value)11U, Max = (UInt32Value)11U, Width = 11.41796875D, Style = (UInt32Value)93U, CustomWidth = true };
            Column column12 = new Column() { Min = (UInt32Value)12U, Max = (UInt32Value)12U, Width = 8.89453125D, Style = (UInt32Value)93U, CustomWidth = true };
            Column column13 = new Column() { Min = (UInt32Value)13U, Max = (UInt32Value)16384U, Width = 8.89453125D, Style = (UInt32Value)93U };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);
            columns1.Append(column6);
            columns1.Append(column7);
            columns1.Append(column8);
            columns1.Append(column9);
            columns1.Append(column10);
            columns1.Append(column11);
            columns1.Append(column12);
            columns1.Append(column13);
            return columns1;
        }

        private static SheetFormatProperties GetSheetFormatProperties()
        {
            return new SheetFormatProperties() { DefaultColumnWidth = 8.89453125D, DefaultRowHeight = 14.4D, DyDescent = 0.55000000000000004D };
        }

        private static SheetViews GetSheetViews()
        {
            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Pane pane1 = new Pane() { VerticalSplit = 2D, TopLeftCell = "A57", ActivePane = PaneValues.BottomLeft, State = PaneStateValues.Frozen };
            Selection selection1 = new Selection() { Pane = PaneValues.BottomLeft, ActiveCell = "B117", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B117" } };

            sheetView1.Append(pane1);
            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            return sheetViews1;
        }

        private static SheetDimension GetSheetDimension()
        {
            return new SheetDimension() { Reference = "B2:L120" };
        }

        private static SheetProperties GetSheetProperties()
        {
            SheetProperties sheetProperties1 = new SheetProperties();
            PageSetupProperties pageSetupProperties1 = new PageSetupProperties() { FitToPage = true };

            sheetProperties1.Append(pageSetupProperties1);
            return sheetProperties1;
        }
    }
}
