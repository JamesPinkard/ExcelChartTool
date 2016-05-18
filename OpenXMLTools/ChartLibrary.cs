using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class ChartLibrary
    {
        public ChartLibrary(SpreadsheetDocument spreadsheetDocument)
        {
            _spreadsheetDocument = spreadsheetDocument;
        }

        public ScatterChartMediator GetScatterChartMediator(string sheetName)
        {
            Chart chartObject = GetChartObject(sheetName);
            var scatterChart = chartObject.PlotArea.GetFirstChild<ScatterChart>();
            return new ScatterChartMediator(scatterChart);
        }

        public BarChartMediator GetBarChartMediator(string sheetName)
        {
            Chart chartObject = GetChartObject(sheetName);
            var barChart = chartObject.PlotArea.GetFirstChild<BarChart>();
            return new BarChartMediator(barChart);
        }

        private Chart GetChartObject(string sheetName)
        {
            var sheetPart = GetSheetPart(sheetName) as ChartsheetPart;
            var drawingsPart = sheetPart.GetPartsOfType<DrawingsPart>().First();
            var chartPart = drawingsPart.GetPartsOfType<ChartPart>().First();
            var chartObject = chartPart.ChartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
            return chartObject;
        }

        private OpenXmlPart GetSheetPart(string worksheetName)
        {
            IEnumerable<Sheet> sheetsWithName = _spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Name == worksheetName);
            Sheet foundSheet = sheetsWithName.First();
            string sheetId = foundSheet.Id;
            OpenXmlPart openXmlPart = _spreadsheetDocument.WorkbookPart.GetPartById(sheetId);

            return openXmlPart;
        }

        SpreadsheetDocument _spreadsheetDocument;
    }
}
