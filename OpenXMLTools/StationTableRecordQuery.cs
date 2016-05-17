using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class StationTableRecordQuery :IRecordQuery
    {
        public StationTableRecordQuery(IRecordParser recordParser)
        {
            _recordParser = recordParser;
        }

        private IEnumerable<int> GetUniqueWeekIndices(IEnumerable<MountainViewField> fields)
        {
            HashSet<int> weekIndexes = new HashSet<int>();

            foreach (var f in fields)
            {
                weekIndexes.Add(f.GetWeek());
            }

            return weekIndexes;
        }

        private IEnumerable<StationTable> CompileStationTables(IEnumerable<MountainViewField> fields)
        {
            List<StationTable> stationMeasurements = new List<StationTable>();

            IEnumerable<IGrouping<string, MountainViewField>> query = fields.GroupBy(field => field.StationName);

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

        public IEnumerable<IRecord> Query(IEnumerable<MountainViewField> fields)
        {
            var stationTables = CompileStationTables(fields);
            var weekIndexes = GetUniqueWeekIndices(fields);
            var records = _recordParser.Parse(stationTables, weekIndexes);
            return records;
        }

        private IRecordParser _recordParser;
    }
}
