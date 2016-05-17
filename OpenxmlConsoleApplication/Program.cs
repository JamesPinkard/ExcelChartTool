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
                var recordParser = new WeekRecordParser();
                var recordQuery = new StationTableRecordQuery(recordParser);
                var fieldFilter = new StationNameFieldFilter("RPW-03");
                var recordProcessor = new RecordProcessor(fields, recordQuery, fieldFilter);                                
                var records = recordProcessor.ProcessRecords();

                // ATTEMPT TO WRITE RECORDS
                WorkbookWriter workbookWriter = new WorkbookWriter(spreadsheetDocument.WorkbookPart);
                var worksheetWriter = workbookWriter.CreateWorksheetWriter("records");
                worksheetWriter.WriteRecords(records);


                //var values = worksheetQuery.GetStationValues();
                //var valueWriter = new RecordWriter(@"rpw_output.csv");
                //valueWriter.Write(values);

                //  set Cumulative Volume Series
                var helper = new SpreadsheetHelper(spreadsheetDocument);
                var sheetPart = helper.GetSheetPart(sumChartSheetName) as ChartsheetPart;
                var drawingsPart = sheetPart.GetPartsOfType<DrawingsPart>().First();
                var chartPart = drawingsPart.GetPartsOfType<ChartPart>().First();
                var chartObject = chartPart.ChartSpace.GetFirstChild<DrawingChart>();
                var scatterChart = chartObject.PlotArea.GetFirstChild<ScatterChart>();
                var series = scatterChart.GetFirstChild<ScatterChartSeries>();

                var formatter = new ScatterChartSeriesFormatter(series);
                formatter.SetValues("records", 2, 4, records.Count());

                // set Pump Rate Bar Chart
                var barSheetPart = helper.GetSheetPart(ratesChartSheetName) as ChartsheetPart;
                var barDrawingsPart = barSheetPart.GetPartsOfType<DrawingsPart>().First();
                var barChartPart = barDrawingsPart.GetPartsOfType<ChartPart>().First();
                var barChartObject = barChartPart.ChartSpace.GetFirstChild<DrawingChart>();
                var barChart = barChartObject.PlotArea.GetFirstChild<BarChart>();
                var barSeries = barChart.GetFirstChild<BarChartSeries>();

                var barFormatter = new BarChartSeriesFormatter(barSeries);
                barFormatter.SetValues("records", 2, 3, records.Count());
            }

            System.Diagnostics.Process.Start(newDocumentName);

        }

        // Sheet names
        //string injectionRateSheetName = @"Weekly Inj_Rates";
        //string extractionRateSheetName = @"Weekly Ext_Rates";
        //string remotePumpingRateSheetName = @"WeeklyRPWs";
        //string rawInjSheetName = @"Sorted_Inj";
        //string rawExtSheetName = @"Sorted_Ext";

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
