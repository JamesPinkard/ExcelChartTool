using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class SheetGenerator
    {
        private readonly Sheets _sheets;

        public SheetGenerator(Sheets sheets)
        {
            if (sheets == null)
                throw new ArgumentException("Sheets");

            _sheets = sheets;
        }

        public Sheet CreateSheet(StringValue name, StringValue id)
        {
            var sheetId = _sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            var createdSheet = new Sheet() { Name = name, SheetId = sheetId, Id = id };
            _sheets.Append(createdSheet);
            return createdSheet;
        }
    }
}
