using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class WorksheetQuery
    {
        public WorksheetQuery(IEnumerable<MountainViewField> fields)
        {
            this._fields = fields;
        }

        public IEnumerable<StationTable> GroupMeasurementsIntoWeeks()
        {
            List<StationTable> stationMeasurements = new List<StationTable>();

            IEnumerable<IGrouping<string, MountainViewField>> query = _fields.GroupBy(field => field.StationName);

            foreach (IGrouping<string, MountainViewField> wellFields in query)
            {
                StationTable weekQuery = new StationTable(wellFields.Key);

                foreach (var field in wellFields)
                {
                    weekQuery.AddField(field);
                }

                stationMeasurements.Add(weekQuery);
            }

            return stationMeasurements;
        }

        public IEnumerable<int> GetUniqueWeekIndices()
        {
            HashSet<int> weekIndexes = new HashSet<int>();

            foreach (var f in _fields)
            {
                weekIndexes.Add(f.GetWeek());
            }

            return weekIndexes;
        }

        public List<RecordByWeek> GetRecordsByWeek()
        {
            var queryByStation = GroupMeasurementsIntoWeeks();
            var weekIndexes = GetUniqueWeekIndices();

            List<RecordByWeek> weeklyRates = new List<RecordByWeek>();
            double cumalativeVolume = 0;

            foreach (var week in weekIndexes)
            {
                List<double> ratesOfTheWeek = new List<double>();
                List<double> volumeOfTheWeek = new List<double>();
                foreach (var station in queryByStation)
                {                    
                    if (station.Contains(week))
                    {
                        var rate = station.GetWeeklyRate(week);
                        var volume = station.GetNetVolume(week);

                        if (rate > 0) ratesOfTheWeek.Add(rate);
                        if (volume > 0 ) volumeOfTheWeek.Add(volume);
                    }
                }
                cumalativeVolume += volumeOfTheWeek.Sum();
                weeklyRates.Add(new RecordByWeek(week, ratesOfTheWeek.Sum(), cumalativeVolume));                
            }

            return weeklyRates;
        }

        public List<RecordByStation> GetRecordsByStation()
        {
            var weeks = GetUniqueWeekIndices();
            var stationGrouping = GroupMeasurementsIntoWeeks();
            List<RecordByStation> stationRecords = new List<RecordByStation>();

            foreach (var week in weeks)
            {
                foreach (var station in stationGrouping)
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

        private readonly IEnumerable<MountainViewField> _fields;
    }
}
