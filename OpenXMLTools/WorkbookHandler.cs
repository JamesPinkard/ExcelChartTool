using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class WorkbookHandler
    {
        public WorkbookHandler(WorkbookPart workbookPart)
        {
            _workbookPart = workbookPart;
        }

        public bool VerifySheet(string worksheetName)
        {
            IEnumerable<Sheet> sheets = _workbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Name == worksheetName);

            if (sheets.Count() == 0)
            {
                return false;
            }
            return true;
        }
        
        public Worksheet GetWorksheet(string worksheetName)
        {
            Sheet sheet = GetSheet(worksheetName);
            WorksheetPart worksheetPart = (WorksheetPart)_workbookPart.GetPartById(sheet.Id);
            return worksheetPart.Worksheet;

        }

        public List<SharedStringItem> GetSharedStringList()
        {
            var sharedStringTable = _workbookPart.SharedStringTablePart.SharedStringTable;
            var sharedStringList = sharedStringTable.Elements<SharedStringItem>().ToList();
            return sharedStringList;
        }

        private Sheet GetSheet(string worksheetName)
        {
            IEnumerable<Sheet> sheetsWithName = _workbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Name == worksheetName);
            return sheetsWithName.First();
        }

        private WorkbookPart _workbookPart;
    }
}
