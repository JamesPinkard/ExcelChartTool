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

        public double GetWeeklyRate(int weekIndex)
        {
            SetWeek(weekIndex);

            var duration = _lastField.MeasureTime - _firstField.MeasureTime;
            var flow = _lastField.TotalizerReading - _firstField.TotalizerReading;
            var rate = flow / duration.TotalMinutes;
            return rate;
        }

        public double GetNetVolume(int weekIndex)
        {
            SetWeek(weekIndex);            
            var totalFlow = _lastField.TotalizerReading - _firstField.TotalizerReading;
           
            return totalFlow;
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

        public RecordByStation GetRecordForWeek(int weekIndex)
        {
            if (Contains(weekIndex))
            {
                return new RecordByStation(StationName, weekIndex, GetWeeklyRate(weekIndex));
            }
            else return null;
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
