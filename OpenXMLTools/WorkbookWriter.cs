using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;


namespace OpenXMLTools
{
    public class WorkbookWriter
    {
        public WorkbookWriter(WorkbookPart workbookPart)
        {
            _workbookPart = workbookPart;
        }

        public WorksheetWriter CreateWorksheetWriter(string worksheetName)
        {
            WorksheetPart worksheetPart = AddWorksheet( worksheetName);
            return new WorksheetWriter(worksheetPart, _workbookPart);
        }


        private WorksheetPart AddWorksheet(string worksheetName)
        {
            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            var root = newWorksheetPart.RootElement;

            var sheet = AssociateSheet(newWorksheetPart);
            sheet.Name = worksheetName;
            newWorksheetPart.Worksheet.Save();
            return newWorksheetPart;
        }

        private Sheet AssociateSheet(OpenXmlPart newWorksheetPart)
        {
            Sheets sheets = _workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = _workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }            

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId };
            sheets.Append(sheet);

            return sheet;
        }

        private WorkbookPart _workbookPart;
    }
}
