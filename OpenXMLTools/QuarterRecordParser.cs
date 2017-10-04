using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class QuarterRecordParser : IRecordParser
    {
        public QuarterRecordParser(QuarterTableParser quarterTableParser, string stationName)
        {
            _quarterTableParser = quarterTableParser;
            _stationName = stationName;
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
                var namedRecord = new NamedCumulativeRecord(weekRecord, _stationName);
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
                var tablesForYearlyQuarter = stationQuarterTables.Where(t => qtable.GetFields().Contains(t.GetFields().First()));
                var tableGroupings = tablesForYearlyQuarter.GroupBy(t => t.GetFields().First().StationName);
                var stationNames = tableGroupings.Select(g => g.Key);

                //var influentNames = new List<string>() { "RPW-06", "RPW-07", "RPW-6/7" };
                var influentNames = new List<string>() { "RPW-06", "RPW-07", "Influent" };

                var containsAllInfluentWells = influentNames.Intersect(stationNames).Count() == influentNames.Count();
                double quarterCumulative = 0;
                
                // Handles the transition quarter where RPW-6/7 is used instead of RPW-06 and RPW-7
                if (containsAllInfluentWells)
                {
                    foreach (var tableGroup in tableGroupings)
                    {
                        var avgFlowRate = tableGroup.First().GetAverageWeeklyFlowRate();
                        //if (tableGroup.Key == "RPW-6/7") { quarterCumulative += avgFlowRate * (5.0 / 13.0); }
                        if (tableGroup.Key == "Influent") { quarterCumulative += avgFlowRate * (2.0 / 13.0); }
                        else { quarterCumulative += avgFlowRate * (11.0 / 13.0); }
                    }
                }
                else
                {
                    var quarterFlows = tablesForYearlyQuarter.Select(r => r.GetAverageWeeklyFlowRate());
                    quarterCumulative = quarterFlows.Sum();                    
                }

                weekRecords[index] = new QuarterCumulativeRecord(firstMeasurementOfQuarter, quarterCumulative);
            }

            return weekRecords;
        }

        private QuarterTableParser _quarterTableParser;
        private string _stationName;
    }
}
