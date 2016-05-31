using System.Collections.Generic;

namespace OpenXMLTools
{
    public interface IRecordWriter
    {
        void WriteRecords(IEnumerable<IRecord> records);
        CellReference GetStartingCell();
        string GetSheetName();
    }
}