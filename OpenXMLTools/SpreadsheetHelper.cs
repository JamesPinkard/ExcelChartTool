using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

using DrawingChart = DocumentFormat.OpenXml.Drawing.Charts.Chart;
using DrawingValues = DocumentFormat.OpenXml.Drawing.Charts.Values;


namespace OpenXMLTools
{
    public class SpreadsheetHelper
    {
        public SpreadsheetHelper(SpreadsheetDocument spreadsheetDocument)
        {
            this._spreadsheetDocument = spreadsheetDocument;
        }

        public bool VerifySheet(string worksheetName)
        {
            IEnumerable<Sheet> sheets = _spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Name == worksheetName);

            if (sheets.Count() == 0)
            {
                return false;
            }
            return true;
        }

        public Sheet GetSheet(string worksheetName)
        {
            IEnumerable<Sheet> sheetsWithName = _spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Name == worksheetName);
            return sheetsWithName.First();
        }

        public Worksheet GetWorksheet(string worksheetName)
        {
            Sheet sheet = GetSheet(worksheetName);
            WorksheetPart worksheetPart = (WorksheetPart)_spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id);
            return worksheetPart.Worksheet;
            
        }

        public OpenXmlPart GetSheetPart(string worksheetName)
        {
            IEnumerable<Sheet> sheetsWithName = _spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Name == worksheetName);
            Sheet foundSheet = sheetsWithName.First();
            string sheetId = foundSheet.Id;
            OpenXmlPart openXmlPart = _spreadsheetDocument.WorkbookPart.GetPartById(sheetId);

            return openXmlPart;
        }

        public WorksheetPart AddWorksheet()
        {
            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = _spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            AssociateSheet(newWorksheetPart);

            return newWorksheetPart;
        }

        public WorksheetPart AddWorksheet(string worksheetName)
        {
            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = _spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheet = AssociateSheet(newWorksheetPart);
            sheet.Name = worksheetName;

            return newWorksheetPart;
        }

        public void BuildChart()
        {
            ChartsheetPart newChartSheetPart = AddChartsheet();


            var drawPart = newChartSheetPart.AddNewPart<DrawingsPart>();
            newChartSheetPart.Chartsheet.Drawing = new Drawing() { Id = newChartSheetPart.GetIdOfPart(drawPart) };

            var chartPart = drawPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            var chart = chartPart.ChartSpace.AppendChild<DrawingChart>(new DrawingChart());
        }

        public ChartsheetPart AddChartsheet()
        {
            // Add a blank ChartsheetPart.
            ChartsheetPart newChartSheetPart = _spreadsheetDocument.WorkbookPart.AddNewPart<ChartsheetPart>();
            newChartSheetPart.Chartsheet = new Chartsheet();

            AssociateSheet(newChartSheetPart);

            return newChartSheetPart;
        }

        private Sheet AssociateSheet(OpenXmlPart newWorksheetPart)
        {
            Sheets sheets = _spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = _spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Give the new worksheet a name.
            string sheetName = "Sheet_" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);

            return sheet;
        }

        private SpreadsheetDocument _spreadsheetDocument;
    }
}
