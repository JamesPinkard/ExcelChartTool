using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class StationTable
    {
        public string StationName
        {
            get
            {
                return _stationName;
            }
        }

        public StationTable(string stationName)
        {
            _stationName = stationName;
            _fieldWeekDictionary = new FieldWeekDictionary();
        }

        public void AddField(MountainViewField field)
        {
            if (field.StationName != StationName)
            {
                return;
            }

            _fieldWeekDictionary.AddField(field);
        }

        public bool Contains(int weekIndex)
        {
            return _fieldWeekDictionary.Contains(weekIndex);
        }

        public WeekTable GetTableForWeek(int weekIndex)
        {
            return _fieldWeekDictionary.GetTableForWeek(weekIndex);
        }

        public RecordByStation GetRecordForWeek(int weekIndex)
        {
            if (Contains(weekIndex))
            {
                var weekTable = GetTableForWeek(weekIndex);
                return new RecordByStation(StationName, weekIndex, weekTable.GetAverageWeeklyFlowRate());
            }
            else return null;
        }

        public IEnumerable<MountainViewField> GetFieldsForWeek(int weekIndex)
        {
            return _fieldWeekDictionary.GetFieldsForWeek(weekIndex);
        }

        public IEnumerable<MountainViewField> GetStationFields()
        {
            return _fieldWeekDictionary.GetFields();
        }

        readonly FieldWeekDictionary _fieldWeekDictionary;
        readonly string _stationName;
        
    }
}
