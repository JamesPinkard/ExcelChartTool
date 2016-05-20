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
            this._stationName = stationName;
            this._verifier = new FieldVerifier(_weeklyMeasurements);
        }

        public void AddField(MountainViewField field)
        {
            if (field.StationName != this.StationName)
            {
                return;
            }

            int weekOfMeasurement = field.GetWeek();
            if (_weeklyMeasurements.ContainsKey(weekOfMeasurement))
            {
                _weeklyMeasurements[weekOfMeasurement].Add(field);
            }
            else
            {
                _weeklyMeasurements[weekOfMeasurement] = new List<MountainViewField>();
                _weeklyMeasurements[weekOfMeasurement].Add(field);
            }
        }

        public bool Contains(int weekIndex)
        {
            return _weeklyMeasurements.ContainsKey(weekIndex);
        }

        public WeekTable GetTableForWeek(int weekIndex)
        {
            SetWeek(weekIndex);
            var fields = GetFieldsForWeek(weekIndex);
            return new WeekTable(weekIndex, fields, _firstField, _lastField);
        }

        public RecordByStation GetRecordForWeek(int weekIndex)
        {
            if (Contains(weekIndex))
            {
                var weekTable = GetTableForWeek(weekIndex);
                return new RecordByStation(StationName, weekIndex, weekTable.GetWeeklyFlowRate());
            }
            else return null;
        }

        public IEnumerable<MountainViewField> GetFieldsForWeek(int weekIndex)
        {
            return _weeklyMeasurements[weekIndex];
        }

        public IEnumerable<MountainViewField> GetStationFields()
        {
            List<MountainViewField> stationFields = new List<MountainViewField>();
            foreach (List<MountainViewField> weekFields in _weeklyMeasurements.Values)
            {
                stationFields.AddRange(weekFields);
            }
            stationFields.OrderBy(f => f.MeasureTime);
            return stationFields;
        }

        private void SetWeek(int weekIndex)
        {
            var fields = _weeklyMeasurements[weekIndex];
            var orderedFields = fields.OrderBy(f => f.MeasureTime);
            _firstField = orderedFields.First();
            _lastField = orderedFields.Last();

            if (_firstField == _lastField)
            {

                int alternativeIndex = _verifier.GetAlternativeIndex(weekIndex);
                var alternativeField = _weeklyMeasurements[alternativeIndex].OrderBy(f => f.MeasureTime);

                if (weekIndex > alternativeIndex)
                {
                    _firstField = orderedFields.First();
                    _lastField = alternativeField.First();
                }
                else
                {
                    _firstField = alternativeField.Last();
                    _lastField = orderedFields.Last();
                }
            }
        }

        private bool InvalidIndex(int weekIndex, int alternativeIndex)
        {
            return !Contains(alternativeIndex) && alternativeIndex != weekIndex;
        }

        private MountainViewField _firstField;
        private MountainViewField _lastField;
        readonly Dictionary<int, List<MountainViewField>> _weeklyMeasurements = new Dictionary<int, List<MountainViewField>>();
        readonly string _stationName;
        private FieldVerifier _verifier;
    }
}
