using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class QuarterRecordParser : IRecordParser
    {
        public QuarterRecordParser(QuarterTableParser quarterTableParser, string name)
        {
            _quarterTableParser = quarterTableParser;
            _name = name;
        }

        // Same as WeekRecordParser
        public IEnumerable<IRecord> Parse(IEnumerable<StationTable> stationTables, IEnumerable<int> weeks)
        {
            List<CumulativeRecord> weekRecords = new List<CumulativeRecord>();
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
                        var rate = weekTable.GetAverageWeeklyFlowRate();
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
                var weekRecord = new WeekCumulativeRecord(week, ratesOfTheWeek.Sum(), cumalativeVolume, volumeOfTheWeek.Sum());
                var namedRecord = new NamedCumulativeRecord(weekRecord, _name);
                weekRecords.Add(namedRecord);
            }

            List<QuarterTable> stationQuarterTables = new List<QuarterTable>();
            foreach (var station in stationTables)
            {
                stationQuarterTables.AddRange(_quarterTableParser.Parse(station.GetStationFields()));
            }
                    

            var stationFields = stationTables.SelectMany(s => s.GetStationFields())
                .OrderBy(f => f.MeasureTime);
            var quarterTables = _quarterTableParser.Parse(stationFields);

            foreach (var qtable in quarterTables)
            {
                var firstFieldOfQuarter = qtable.GetFields().First();
                
                var indexOfFirstWeekInQuarter = firstFieldOfQuarter.GetWeek();
                var firstFieldOfWeek = stationFields.Where(f => f.GetWeek() == indexOfFirstWeekInQuarter).First();

                CumulativeRecord firstMeasurementOfQuarter;
                int index;

                if (firstFieldOfQuarter.MeasureTime.Month == firstFieldOfWeek.MeasureTime.Month)
                {
                    // Set index and first measurement of week
                    var firstWeekOfQuarter = weekRecords.Where(r => r.Week == indexOfFirstWeekInQuarter);
                    firstMeasurementOfQuarter = firstWeekOfQuarter.First();
                    index = weekRecords.IndexOf(firstMeasurementOfQuarter);
                }
                else
                {
                    // Set index and first measurement of quarter
                    var firstWeekOfQuarter = weekRecords.Where(r => r.Week == indexOfFirstWeekInQuarter);
                    index = weekRecords.IndexOf(firstWeekOfQuarter.First()) + 1;
                    firstMeasurementOfQuarter = weekRecords[index];
                }

                var quarterFlows = stationQuarterTables.Where(t => qtable.GetFields().Contains(t.GetFields().First()))
                    .Select(r => r.GetAverageWeeklyFlowRate());

                var quarterCumulative = quarterFlows.Sum();                    

                weekRecords[index] = new QuarterCumulativeRecord(firstMeasurementOfQuarter, quarterCumulative);
            }

            return weekRecords;
        }

        private QuarterTableParser _quarterTableParser;
        private string _name;
    }
}
