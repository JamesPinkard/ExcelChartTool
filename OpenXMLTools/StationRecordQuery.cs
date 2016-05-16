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

        public List<StationRecord> GetStationRecords(string sheetName)
        {
            var worksheetQuery = _recordProvider.MakeWorksheetQuery(sheetName);
            return worksheetQuery.GetStationValues();
        }

        public List<WeekRateRecord> GetWeeklyRates(string sheetName)
        {
            var worksheetQuery = _recordProvider.MakeWorksheetQuery(sheetName);
            return worksheetQuery.GetWeeklyRates();
        }

        private readonly RecordProvider _recordProvider;
    }
}
