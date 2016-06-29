using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class StationReport
    {
        public StationReport(IEnumerable<MountainViewField> fields, StationTableParser stationTableParser)
        {
            _stationTableParser = stationTableParser;
            _fields = fields;
            _uniqueWeekFieldQuery = new UniqueWeekFieldQuery();
        }

        public IEnumerable<IRecord> ProcessReport()
        {
            var report = new List<IRecord>();
            var tables = _stationTableParser.CompileStationTables(_fields);            

            foreach (var stationTable in tables)
            {
                var stationFields = stationTable.GetStationFields();
                var weeks = _uniqueWeekFieldQuery.GetUniqueWeekIndices(stationFields);
                var firstField = stationFields.First();
                var parser = new MeasurementRecordParser(firstField);

                foreach (var week in weeks)
                {
                    var weekTable = stationTable.GetTableForWeek(week);
                    var measurementRecords = parser.ProcessMeasurementRecord(weekTable);
                    report.AddRange(measurementRecords);
                }               
            }

            return report;
        }

        StationTableParser _stationTableParser;
        IEnumerable<MountainViewField> _fields;
        UniqueWeekFieldQuery _uniqueWeekFieldQuery;
    }
}
