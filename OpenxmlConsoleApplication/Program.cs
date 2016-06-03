using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

using DrawingChart = DocumentFormat.OpenXml.Drawing.Charts.Chart;
using DrawingValues = DocumentFormat.OpenXml.Drawing.Charts.Values;

using OpenXMLTools;

namespace OpenxmlConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            var reportGenerator = new ReportGenerator();
            reportGenerator.GenerateReport();
        }

        private static void TestReportGeneration()
        {
            string docName = @".\O&M_Master Spreadsheet_Q1 2016.xlsx";
            string sumChartSheetName = @"SumVol";
            string ratesChartSheetName = @"WeeklyFlowRates";
            string newDocumentName = @".\O&M_Copy.xlsx";
            string rawRPWSheetName = @"RAW Data_all";

            CopyWorkbook(docName, newDocumentName);

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(newDocumentName, true))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                var workbookHandler = new WorkbookHandler(workbookPart);
                var worksheet = workbookHandler.GetWorksheet(rawRPWSheetName);
                var rowTable = new WorksheetRowTable(worksheet);

                var sharedStringList = workbookHandler.GetSharedStringList();
                var parserFacory = new FieldParserFactory(sharedStringList);
                var parser = parserFacory.MakeParser();

                var fieldProcessor = new FieldProcessor(rowTable, parser);
                var fields = fieldProcessor.ProcessFields();

                var stationTableParser = new StationTableParser();

                // Process effluent data;
                var quarterParser = new QuarterTableParser(new ThirdQuarterState());
                var effluentRecordParser = new QuarterRecordParser(quarterParser, "Effluent");
                var recordQuery = new StationTableRecordQuery(stationTableParser, effluentRecordParser);
                var effluentFieldFilter = new StationNameFieldFilter("RPW-03");
                var effluentrecordProcessor = new RecordProcessor(fields, recordQuery, effluentFieldFilter);
                var records = effluentrecordProcessor.ProcessRecords();

                // Process influent data;
                var influentquarterParser = new QuarterTableParser(new ThirdQuarterState());
                var influentRecordParser = new QuarterRecordParser(influentquarterParser, "Influent");
                var influentRecordQuery = new StationTableRecordQuery(stationTableParser, influentRecordParser);
                var influentFieldFilter = new StationNameFieldFilter(new List<string>() { "RPW-06", "RPW-07" });
                var influentRecordProcessor = new RecordProcessor(fields, influentRecordQuery, influentFieldFilter);
                var influentRecords = influentRecordProcessor.ProcessRecords();

                // ATTEMPT TO WRITE RECORDS
                WorkbookWriter workbookWriter = new WorkbookWriter(spreadsheetDocument.WorkbookPart);
                var worksheetWriter = workbookWriter.CreateWorksheetWriter("records", new CellReference(2, 2));
                var rangeProcessor = new RangeProcessor(worksheetWriter);
                var influentSheetRange = rangeProcessor.AddRecords(influentRecords);
                var sheetRange = rangeProcessor.AddRecords(records);
                rangeProcessor.WriteRecords();

                var cumulativeWorksheetPart = worksheetWriter.GetWorksheetPart();
                var cumulativeFormatter = new WorksheetFormatter(cumulativeWorksheetPart);
                cumulativeFormatter.FormatSheet();

                //var values = worksheetQuery.GetStationValues();
                //var valueWriter = new RecordWriter(@"rpw_output.csv");
                //valueWriter.Write(values);

                var chartLibrary = new ChartLibrary(spreadsheetDocument);
                var scatterChartMediator = chartLibrary.GetScatterChartMediator(sumChartSheetName);
                var effluentScatterSeriesFormatter = GetExtractionOrEffluentSeries(scatterChartMediator, "Extraction", "Effluent");

                //  set Cumulative Volume Series
                var xFormula = sheetRange.GetColumnFormula(4);
                var volumeCellFormula = sheetRange.GetColumnFormula(6);
                effluentScatterSeriesFormatter.SetSeriesFormula(xFormula, volumeCellFormula);

                var barChartMediator = chartLibrary.GetBarChartMediator(ratesChartSheetName);
                var effluentSeriesFormatter = GetExtractionOrEffluentSeries(barChartMediator, "Extraction", "Effluent");
                var weekRateFormula = sheetRange.GetColumnFormula(5);
                effluentSeriesFormatter.SetSeriesFormula(xFormula, weekRateFormula);


                // set Pump Rate Bar Chart
                var influentScatterSeriesFormatter = GetExtractionOrEffluentSeries(scatterChartMediator, "Injection", "Influent");
                var influentFormula = influentSheetRange.GetColumnFormula(4);
                var influentVolumeFormula = influentSheetRange.GetColumnFormula(6);
                influentScatterSeriesFormatter.SetSeriesFormula(influentFormula, influentVolumeFormula);

                var influentSeriesFormatter = GetExtractionOrEffluentSeries(barChartMediator, "Injection", "Influent");
                var influentWeekRateFormula = influentSheetRange.GetColumnFormula(5);
                influentSeriesFormatter.SetSeriesFormula(influentFormula, influentWeekRateFormula);

                var stationTableParserForReport = new StationTableParser();
                var stationReport = new QuarterlyReport(fields, stationTableParserForReport);
                var reportRecords = stationReport.ProcessReport();
                var stationReportWriter = workbookWriter.CreateWorksheetWriter("stationReport", new CellReference(2, 2));
                stationReportWriter.WriteRecords(reportRecords);

                var reportWorksheetPart = stationReportWriter.GetWorksheetPart();
                var reportFormatter = new WorksheetFormatter(reportWorksheetPart);
                reportFormatter.FormatSheet();
            }

            System.Diagnostics.Process.Start(newDocumentName);
        }

        // Sheet names
        //string injectionRateSheetName = @"Weekly Inj_Rates";
        //string extractionRateSheetName = @"Weekly Ext_Rates";
        //string remotePumpingRateSheetName = @"WeeklyRPWs";
        //string rawInjSheetName = @"Sorted_Inj";
        //string rawExtSheetName = @"Sorted_Ext";
        private static ISeriesFormatter GetExtractionOrEffluentSeries(IChartMediator chartMediator, string originalName, string newName)
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

        private static void PrintStationRates(List<Tuple<string, int, double>> stationRates)
        {
            using (StreamWriter writer = new StreamWriter("output.csv"))
            {

                foreach (var field in stationRates)
                {
                    Console.WriteLine("{0} {1} {2}", field.Item1, field.Item2, field.Item3);
                    writer.WriteLine("{0}, Week {1}, {2}, {3}", field.Item1, field.Item2, MountainViewField.GetSundayOfWeek(field.Item2), field.Item3);
                }

            }
        }

        private static void PrintWeekRates(Dictionary<int, double> weeklyRates)
        {
            using (StreamWriter writer = new StreamWriter("output.txt"))
            {

                foreach (var field in weeklyRates)
                {
                    Console.WriteLine("{0} {1}", field.Key, field.Value);
                    writer.WriteLine("{0} {1}", field.Key, field.Value);
                }

            }
        }

        private static void PrintFields(IEnumerable<MountainViewField> rpwFields)
        {
            using (StreamWriter writer = new StreamWriter("output.txt"))
            {

                foreach (var field in rpwFields)
                {
                    Console.WriteLine("{0} {1}", field.GetWeek(), field.MeasureTime);
                    writer.WriteLine("{0} {1}", field.GetWeek(), field.MeasureTime);
                }

            }
        }

        private static void PrintTime(List<SharedStringItem> sharedStringList, IEnumerable<Row> rpwRows)
        {
            foreach (var row in rpwRows)
            {
                Console.WriteLine(row.RowIndex.Value);

                Cell cell = row.ChildElements.ElementAt(1) as Cell;

                if (cell.DataType == "s")
                {
                    row.CloneNode(true);
                    var sharedIndex = int.Parse(cell.CellValue.Text);
                    Console.Write("{0} ", sharedStringList[sharedIndex].Text.Text);
                }
                else
                {
                    Console.Write("{0} ", cell.CellValue.Text);
                }

                Console.Write("\r\n");
            }
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
    }
}
