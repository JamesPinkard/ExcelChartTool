using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class WorksheetRowTable : IRowTable
    {
        public WorksheetRowTable(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public IEnumerable<Row> GetRows()
        {
            return _worksheet.Descendants<Row>();
        }

        private Worksheet _worksheet;
    }
}
