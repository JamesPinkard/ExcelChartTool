using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class StationTableRecordQuery :IRecordQuery
    {
        public StationTableRecordQuery(StationTableParser stationTableParser, IRecordParser recordParser)
        {
            _recordParser = recordParser;
            _stationTableParser = stationTableParser;
        }               

        public IEnumerable<IRecord> Query(IEnumerable<MountainViewField> fields, IEnumerable<int> weeks)
        {
            var stationTables = _stationTableParser.CompileStationTables(fields);            
            var records = _recordParser.Parse(stationTables, weeks);
            return records;
        }
        
        private IRecordParser _recordParser;
        private StationTableParser _stationTableParser;
    }
}
