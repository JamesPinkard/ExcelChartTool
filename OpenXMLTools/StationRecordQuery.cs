using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class StationRecordQuery
    {
        public StationRecordQuery(RecordProvider recordProvider)
        {
            _recordProvider = recordProvider;
        }

        public List<RecordByStation> GetStationRecords(string sheetName)
        {
            var worksheetQuery = _recordProvider.MakeWorksheetQuery(sheetName);
            return worksheetQuery.GetRecordsByStation();
        }

        public List<RecordByWeek> GetWeeklyRates(string sheetName)
        {
            var worksheetQuery = _recordProvider.MakeWorksheetQuery(sheetName);
            return worksheetQuery.GetRecordsByWeek();
        }

        private readonly RecordProvider _recordProvider;
    }
}
