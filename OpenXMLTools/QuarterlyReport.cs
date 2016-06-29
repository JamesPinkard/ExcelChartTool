using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class QuarterlyReport
    {
        public QuarterlyReport(IEnumerable<MountainViewField> fields, StationTableParser stationTableParser)
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
                var quarterlyParser = new QuarterTableParser(new ThirdQuarterState());
                var quarterTables = quarterlyParser.Parse(stationFields);
                var parser = new MeasurementRecordParser(stationFields.First());                    
                foreach (var quarter in quarterTables)
                {
                    var records = parser.ProcessMeasurementRecord(quarter);
                    report.AddRange(records);                   
                }
            }

            return report;
        }

        StationTableParser _stationTableParser;
        IEnumerable<MountainViewField> _fields;
        UniqueWeekFieldQuery _uniqueWeekFieldQuery;
    }
}
