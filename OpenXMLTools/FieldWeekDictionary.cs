using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class FieldWeekDictionary
    {
        public FieldWeekDictionary()
        {
            _verifier = new FieldVerifier(_weeklyMeasurements);
        }

        public FieldWeekDictionary(IEnumerable<MountainViewField> fields )
        {
            _verifier = new FieldVerifier(_weeklyMeasurements);
            foreach (var f in fields)
            {
                AddField(f);
            }
            
        }

        public void AddField(MountainViewField field)
        {
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

        public IEnumerable<MountainViewField> GetFieldsForWeek(int weekIndex)
        {
            return _weeklyMeasurements[weekIndex];
        }

        public IEnumerable<List<MountainViewField>> Values()
        {
            return _weeklyMeasurements.Values;
        }

        public IEnumerable<MountainViewField> GetFields()
        {
            List<MountainViewField> AllFields = new List<MountainViewField>();
            foreach (List<MountainViewField> weekFields in Values())
            {
                AllFields.AddRange(weekFields);
            }
            AllFields.OrderBy(f => f.MeasureTime);
            return AllFields;
        }

        private void SetWeek(int weekIndex)
        {

            var fields = _weeklyMeasurements[weekIndex];
            var orderedFields = fields.OrderBy(f => f.MeasureTime);

            int alternativeIndex = _verifier.GetAlternativeIndex(weekIndex);
            var alternativeField = _weeklyMeasurements[alternativeIndex].OrderBy(f => f.MeasureTime);

            if (weekIndex == 1)
            {
                _firstField = orderedFields.First();
                _lastField = orderedFields.Last();
            }
            else
            {
                if (weekIndex > alternativeIndex)
                {
                    _lastField = orderedFields.Last();
                    _firstField = alternativeField.Last();
                }
                else
                {
                    _lastField = alternativeField.First();
                    _firstField = orderedFields.First();
                }
            }
        }

        private MountainViewField _firstField;
        private MountainViewField _lastField;
        readonly Dictionary<int, List<MountainViewField>> _weeklyMeasurements = new Dictionary<int, List<MountainViewField>>();
        private FieldVerifier _verifier;
    }
}
