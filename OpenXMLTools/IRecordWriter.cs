using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public interface IRecordWriter
    {
        void WriteRecords(IEnumerable<IRecord> records);
        WorksheetPart GetWorksheetPart();
        CellReference GetStartingCell();
        string GetSheetName();
    }
}