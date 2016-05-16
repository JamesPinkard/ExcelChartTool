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

namespace OpenXMLTools
{
    public class RecordProvider
    {
        public RecordProvider(SpreadsheetDocument spreadsheetDocument)
        {
            _spreadsheetDocument = spreadsheetDocument;
            var sharedStringTable = _spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable;
            _sharedStringList = sharedStringTable.Elements<SharedStringItem>().ToList();
            _helper = new SpreadsheetHelper(_spreadsheetDocument);
        }

        public WorksheetQuery MakeWorksheetQuery(string sheetName)
        {            
            FieldParserFactory parserFactory = new FieldParserFactory(_sharedStringList);
            IFieldParser parser = parserFactory.MakeParser();

            Worksheet worksheet = Helper.GetWorksheet(sheetName);
            IEnumerable<Row> rows = worksheet.Descendants<Row>();

            IEnumerable<MountainViewField> fields = parser.Parse(rows);
            WorksheetQuery query = MakeQuery(fields);
            return query;
        }

        private WorksheetQuery MakeQuery(IEnumerable<MountainViewField> fields)
        {
            return new WorksheetQuery(fields);
        }

        public SpreadsheetHelper Helper
        {
            get
            {
                return _helper;
            }
        }
                
        private readonly SpreadsheetDocument _spreadsheetDocument;
        private readonly List<SharedStringItem> _sharedStringList;
        private readonly SpreadsheetHelper _helper;
    }
}
