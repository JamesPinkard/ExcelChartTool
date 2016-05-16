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
    class SpreadsheetChecker
    {
        private static void CheckSheetsExist(string injectionRateSheet, string extractionRateSheet, string remotePumpingRateSheet, SpreadsheetHelper helper)
        {
            bool injSheetExists = helper.VerifySheet(injectionRateSheet);
            bool extSheetExists = helper.VerifySheet(extractionRateSheet);
            bool rpwSheetExists = helper.VerifySheet(remotePumpingRateSheet);

            Console.WriteLine("Injection Sheet Should Exist: {0}", injSheetExists);
            Console.WriteLine("Extraction Sheet Should Exist: {0}", extSheetExists);
            Console.WriteLine("Remote Pumping Sheet Should Exist: {0}", rpwSheetExists);
        }

        private static void CheckChartValues(string sumChartSheetName, SpreadsheetDocument spreadsheetDocument)
        {
            SpreadsheetHelper helper = new SpreadsheetHelper(spreadsheetDocument);

            ChartsheetPart sumChartSheetPart = helper.GetSheetPart(sumChartSheetName) as ChartsheetPart;

            IEnumerable<DrawingsPart> drawings = sumChartSheetPart.GetPartsOfType<DrawingsPart>();
            DrawingsPart drawPart = drawings.First();
            IEnumerable<ChartPart> allChartParts = drawPart.GetPartsOfType<ChartPart>();
            ChartPart chartPart = allChartParts.First();
            var myChart = chartPart.ChartSpace.GetFirstChild<DrawingChart>();
            var scatterChart = myChart.PlotArea.GetFirstChild<ScatterChart>();
            var series = scatterChart.GetFirstChild<ScatterChartSeries>();
            var xValues = series.GetFirstChild<XValues>();
            var yValues = series.GetFirstChild<YValues>();

            var xRefs = xValues.Elements<NumberReference>();
            var yRefs = yValues.Elements<NumberReference>();
            //string[] formulaText = new string[] { xRefs.First().Formula.Text, yRefs.First().Formula.Text };
            string[] formulaText = new string[] { sumChartSheetPart.Chartsheet.OuterXml };

            File.WriteAllLines(@".\output.txt", formulaText);
            Console.WriteLine(sumChartSheetPart.Chartsheet.OuterXml);
        }
    }
}
