﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class ReportGenerator
    {
        public void GenerateReport(string docName)
        {           
            
            string sumChartSheetName = @"SumVol";
            string ratesChartSheetName = @"WeeklyFlowRates";
            string newDocumentName = @".\O&M_Copy.xlsx";
            string rawRPWSheetName = @"RAW Data_all";

            CopySheet(docName, newDocumentName);

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(newDocumentName, true))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;

                // create cumulative volume chartsheet part
                var sheets = workbookPart.Workbook.Sheets;
                var sheetGenerator = new SheetGenerator(sheets);
                var openXmlPartGenerator = new OpenXmlPartGenerator(workbookPart);
                var cumulativeVolumeChartGenerator = new CumulativeVolumeChartGenerator();
                var cumalativeChartsheetMaker = new ChartsheetMaker(sheetGenerator, openXmlPartGenerator, cumulativeVolumeChartGenerator);
                cumalativeChartsheetMaker.MakeChartSheet("SumVol", "rId6");            
                                
                // create flow rate chartsheet part
                FlowRateChartGenerator flowRateChartGenerator = new FlowRateChartGenerator();                
                var flowRateChartsheetMaker = new ChartsheetMaker(sheetGenerator, openXmlPartGenerator, flowRateChartGenerator);
                flowRateChartsheetMaker.MakeChartSheet("WeeklyFlowRates", "rId7");

                // create worksheet row table
                var stylePartGenerator = new WorkbookStylesPartGenerator(workbookPart);
                stylePartGenerator.CreateWorkbookStylesPart();
                var workbookHandler = new WorkbookHandler(workbookPart);
                var worksheet = workbookHandler.GetWorksheet(rawRPWSheetName);
                var rowTable = new WorksheetRowTable(worksheet);

                // create field parser
                var sharedStringList = workbookHandler.GetSharedStringList();
                var parserFacory = new FieldParserFactory(sharedStringList);
                var parser = parserFacory.MakeParser();

                // parse extraction well fields
                var fieldProcessor = new FieldProcessor(rowTable, parser);
                var rawFields = fieldProcessor.ProcessFields();
                var readingErrorLogger = new ReadingErrorLogger();
                readingErrorLogger.Log(rawFields);
                var totalizerAdjuster = new WellTotalizerAdjuster();
                totalizerAdjuster.AdjustReading(rawFields, "Effluent", DateTime.Parse("4/1/2017"), 516600);
                var extractionWellModifier = new ExtractionWellFieldModifier();
                var fields = extractionWellModifier.Modify(rawFields);

                // Process influent data;
                var stationTableParser = new StationTableParser();
                var influentquarterParser = new QuarterTableParser(new ThirdQuarterState());
                var influentRecordParser = new QuarterRecordParser(influentquarterParser, "Influent");
                var influentRecordQuery = new StationTableRecordQuery(stationTableParser, influentRecordParser);

                var influentStationNameFilter = new StationNameFieldFilter(new List<string>() { "RPW-06", "RPW-07" });
                //var influentFieldFilter = new ReplacementFieldFilter(influentStationNameFilter, @"RPW-6/7");
                var influentFieldFilter = new ReplacementFieldFilter(influentStationNameFilter, @"Influent");

                var influentRecordProcessor = new RecordProcessor(fields, influentRecordQuery, influentFieldFilter);
                var influentRecords = influentRecordProcessor.ProcessRecords();

                // Process effluent data;
                var quarterParser = new QuarterTableParser(new ThirdQuarterState());
                var effluentRecordParser = new QuarterRecordParser(quarterParser, "Effluent");
                var recordQuery = new StationTableRecordQuery(stationTableParser, effluentRecordParser);
                
                //var effluentFieldFilter = new StationNameFieldFilter("RPW-03");
                var effluentFieldFilter = new StationNameFieldFilter("Effluent");

                var effluentrecordProcessor = new RecordProcessor(fields, recordQuery, effluentFieldFilter);
                var records = effluentrecordProcessor.ProcessRecords();


                // ATTEMPT TO WRITE RECORDS
                WorkbookWriter workbookWriter = new WorkbookWriter(spreadsheetDocument.WorkbookPart);

                Table2bGenerator tableGenerator = new Table2bGenerator(workbookWriter);
                var worksheetWriter = tableGenerator.GenerateWriter("records", new CellReference(2, 2));
                var rangeProcessor = new RangeProcessor(worksheetWriter);
                var influentSheetRange = rangeProcessor.AddRecords(influentRecords);
                var sheetRange = rangeProcessor.AddRecords(records);
                rangeProcessor.WriteRecords();
                tableGenerator.FormatWorksheet(worksheetWriter.GetWorksheetPart());


                var chartLibrary = new ChartLibrary(spreadsheetDocument);
                var scatterChartMediator = chartLibrary.GetScatterChartMediator(sumChartSheetName);
                var effluentScatterSeriesFormatter = GetExtractionOrEffluentSeries(scatterChartMediator, "Extraction", "Effluent");

                //  set Cumulative Volume Series
                var xFormula = sheetRange.GetColumnFormula(4);
                var volumeCellFormula = sheetRange.GetColumnFormula(6);
                effluentScatterSeriesFormatter.SetSeriesFormula(xFormula, volumeCellFormula);

                var barChartMediator = chartLibrary.GetBarChartMediator(ratesChartSheetName);
                var effluentSeriesFormatter = GetExtractionOrEffluentSeries(barChartMediator, "Extraction", "Effluent");
                var weekRateFormula = sheetRange.GetColumnFormula(7);
                effluentSeriesFormatter.SetSeriesFormula(xFormula, weekRateFormula);


                // set Pump Rate Bar Chart
                var influentScatterSeriesFormatter = GetExtractionOrEffluentSeries(scatterChartMediator, "Injection", "Influent");
                var influentFormula = influentSheetRange.GetColumnFormula(4);
                var influentVolumeFormula = influentSheetRange.GetColumnFormula(6);
                influentScatterSeriesFormatter.SetSeriesFormula(influentFormula, influentVolumeFormula);

                var influentSeriesFormatter = GetExtractionOrEffluentSeries(barChartMediator, "Injection", "Influent");
                var influentWeekRateFormula = influentSheetRange.GetColumnFormula(7);
                influentSeriesFormatter.SetSeriesFormula(influentFormula, influentWeekRateFormula);

                var stationTableParserForReport = new StationTableParser();
                var stationReport = new QuarterlyReport(fields, stationTableParserForReport);
                var reportRecords = stationReport.ProcessReport();

                Table2aGenerator table2a = new Table2aGenerator(workbookWriter);
                var stationReportWriter = table2a.GenerateWriter("stationReport", new CellReference(2, 2));
                stationReportWriter.WriteRecords(reportRecords);

                var reportWorksheetPart = stationReportWriter.GetWorksheetPart();
                table2a.FormatWorksheet(reportWorksheetPart);                
                //var reportFormatter = new WorksheetFormatter(reportWorksheetPart);
                //reportFormatter.FormatSheet();
                
            }

            System.Diagnostics.Process.Start(newDocumentName);
        }

        public void FormatTableBColumns(WorksheetPart worksheetPart)
        {
            var sheetFormatProperties = GenerateSheetFormatProperties();
            var columns = GenerateColumns();

            var worksheet = worksheetPart.Worksheet;
            worksheet.SheetFormatProperties = sheetFormatProperties;
        }

        public void FormatTableBHeaders(WorksheetPart worksheetPart)
        {
            var pageMargins = GeneratePageMargins();
            var headerFooter = GenerateHeaderFooter();

            var worksheet = worksheetPart.Worksheet;
            worksheet.Append(pageMargins);
            worksheet.Append(headerFooter);
        }


        // Creates an SheetFormatProperties instance and adds its children.
        public SheetFormatProperties GenerateSheetFormatProperties()
        {
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultColumnWidth = 8.88671875D, DefaultRowHeight = 14.4D, DyDescent = 0.3D };
            return sheetFormatProperties1;
        }
        
        // Creates an Columns instance and adds its children.
        public Columns GenerateColumns()
        {
            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 3.7109375D, Style = (UInt32Value)298U, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 12D, Style = (UInt32Value)298U, BestFit = true, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 8.85546875D, Style = (UInt32Value)298U };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 11.140625D, Style = (UInt32Value)298U, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 18D, Style = (UInt32Value)269U, BestFit = true, CustomWidth = true };
            Column column6 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 15.7109375D, Style = (UInt32Value)269U, CustomWidth = true };
            Column column7 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 9.42578125D, Style = (UInt32Value)298U, CustomWidth = true };
            Column column8 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 10.140625D, Style = (UInt32Value)298U, CustomWidth = true };
            Column column9 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 8.85546875D, Style = (UInt32Value)298U };
            Column column10 = new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 9.7109375D, Style = (UInt32Value)298U, CustomWidth = true };
            Column column11 = new Column() { Min = (UInt32Value)11U, Max = (UInt32Value)11U, Width = 11.42578125D, Style = (UInt32Value)298U, CustomWidth = true };
            Column column12 = new Column() { Min = (UInt32Value)12U, Max = (UInt32Value)12U, Width = 8.85546875D, Style = (UInt32Value)298U, CustomWidth = true };
            Column column13 = new Column() { Min = (UInt32Value)13U, Max = (UInt32Value)16384U, Width = 8.85546875D, Style = (UInt32Value)298U };

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

        // Creates an PageMargins instance and adds its children.
        public PageMargins GeneratePageMargins()
        {
            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 1.2D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            return pageMargins1;
        }

        // Creates an HeaderFooter instance and adds its children.
        public HeaderFooter GenerateHeaderFooter()
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

        private ISeriesFormatter GetExtractionOrEffluentSeries(IChartMediator chartMediator, string originalName, string newName)
        {
            ISeriesFormatter seriesFormatter;
            if (chartMediator.HasSeries(originalName))
            {
                seriesFormatter = chartMediator.GetSeriesFormatter(originalName);
                seriesFormatter.SetSeriesTitle(newName);
            }
            else
            {
                seriesFormatter = chartMediator.GetSeriesFormatter(newName);
            }
            return seriesFormatter;
        }

        private static void CopyWorkbook(string docName, string newDocumentName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(docName, false))
            using (SpreadsheetDocument newDocument = SpreadsheetDocument.Create(newDocumentName, SpreadsheetDocumentType.Workbook))
            {
                foreach (var part in spreadsheetDocument.Parts)
                {
                    newDocument.AddPart(part.OpenXmlPart, part.RelationshipId);
                }
            }
        }

        private static void CopySheet(string docName, string newDocumentName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(docName, false))
            using (SpreadsheetDocument newDocument = SpreadsheetDocument.Create(newDocumentName, SpreadsheetDocumentType.Workbook))
            {
                var newWorkbook = newDocument.AddWorkbookPart();

                newWorkbook.AddNewPart<WorkbookStylesPart>("rId4");
                var stringTablePart = spreadsheetDocument.WorkbookPart.SharedStringTablePart;
                newWorkbook.AddPart(stringTablePart, "rId5");                
                
                
                newWorkbook.Workbook = new Workbook();
                var bookviews = GenerateBookViews();
                newWorkbook.Workbook.Append(bookviews);
                newWorkbook.Workbook.Append(new Sheets());

                IEnumerable<Sheet> sheetsWithName = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>()
                    .Where(s => s.Name == "RAW Data_all");
                Sheet foundSheet = sheetsWithName.First();
                string sheetId = foundSheet.Id;
                OpenXmlPart rawData = spreadsheetDocument.WorkbookPart.GetPartById(sheetId);

                var rawDataPart = newWorkbook.AddPart(rawData, "rId1");
                
                var sheets = newWorkbook.Workbook.Sheets;                
                Sheet rawDataSheet = new Sheet() { Name = "RAW Data_all", SheetId = 1, Id = newWorkbook.GetIdOfPart(rawDataPart) };
                sheets.Append(rawDataSheet);                               
            }
        }

        // Creates an BookViews instance and adds its children.
        static public BookViews GenerateBookViews()
        {
            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 0, YWindow = 0, WindowWidth = (UInt32Value)20310U, WindowHeight = (UInt32Value)9540U };

            bookViews1.Append(workbookView1);
            return bookViews1;
        }
    }
}
