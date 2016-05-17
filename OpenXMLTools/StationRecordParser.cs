using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class StationRecordParser : IRecordParser
    {
        public IEnumerable<IRecord> Parse(IEnumerable<StationTable> stationTables, IEnumerable<int> weeks)
        {
            List<IRecord> stationRecords = new List<IRecord>();
            foreach (var week in weeks)
            {
                foreach (var station in stationTables)
                {
                    if (station.Contains(week))
                    {
                        var record = station.GetRecordForWeek(week);
                        stationRecords.Add(record);
                    }
                }
            }
            return stationRecords;
        }
    }
}
