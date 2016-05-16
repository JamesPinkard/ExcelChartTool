using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;


namespace OpenXMLTools
{
    public class WorksheetQuery
    {
        public WorksheetQuery(WeeklyQueryFactory queryFactory)
        {                       
            _factory = queryFactory;
        }

        public List<WeekRateRecord> GetWeeklyRates()
        {            
            var weeklyRates = _factory.GetWeeklyRates();

            return weeklyRates;
        }

        public List<StationRecord> GetStationValues()
        {
            return _factory.GetStationValues();
        }

        public IEnumerable<StationTable> RatesByStation()
        {
            return _factory.GroupMeasurementsIntoWeeks();
        }                      
        
        private WeeklyQueryFactory _factory;
    }
}
