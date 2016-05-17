using System.Collections.Generic;

namespace OpenXMLTools
{
    public interface IRecordParser
    {
        IEnumerable<IRecord> Parse(IEnumerable<StationTable> stationTables, IEnumerable<int> weeks);
    }
}