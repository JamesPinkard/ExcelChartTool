using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class WeekRecordParser : IRecordParser
    {
        public IEnumerable<IRecord> Parse(IEnumerable<StationTable> stationTables, IEnumerable<int> weeks)
        {
            List<RecordByWeek> weekRecords = new List<RecordByWeek>();
            double cumalativeVolume = 0;

            foreach (var week in weeks)
            {
                List<double> ratesOfTheWeek = new List<double>();
                List<double> volumeOfTheWeek = new List<double>();
                foreach (var station in stationTables)
                {
                    if (station.Contains(week))
                    {
                        var weekTable = station.GetTableForWeek(week);
                        var rate = weekTable.GetWeeklyFlowRate();
                        var volume = weekTable.GetNetVolume();

                        if (rate > 0) ratesOfTheWeek.Add(rate);
                        if (volume > 0) volumeOfTheWeek.Add(volume);
                    }
                    else
                    {
                        ratesOfTheWeek.Add(0);
                        volumeOfTheWeek.Add(0);
                    }
                }
                cumalativeVolume += volumeOfTheWeek.Sum();
                weekRecords.Add(new RecordByWeek(week, ratesOfTheWeek.Sum(), cumalativeVolume));
            }

            return weekRecords;
        }
    }
}
