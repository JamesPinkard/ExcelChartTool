using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class QuarterTable
    {
        public QuarterTable(FieldWeekDictionary fields, MountainViewField previousQuarterField)
        {
            _fields = fields;
            _previousQuarterField = previousQuarterField;
        }

        public double GetAverageWeeklyFlowRate()
        {
            var quarterFields = GetFields();
            var firstField = _previousQuarterField;
            var lastField = quarterFields.Last();


            var duration = (lastField.MeasureTime - firstField.MeasureTime).TotalMinutes;
            var flow = lastField.TotalizerReading - firstField.TotalizerReading;
            var rate = flow / duration;

            return rate;
        }

        public WeekTable GetTableForWeek(int weekIndex)
        {
            WeekTable result;
            if (weekIndex == GetFields().First().GetWeek())
            {
                var weekTable = _fields.GetTableForWeek(weekIndex);
                var weekFields = weekTable.GetFieldsForWeek();
                var firstField = _previousQuarterField;
                var lastField = weekFields.Last();
                result = new WeekTable(weekTable.GetWeek(), weekFields, firstField, lastField);
            }
            else
            {
                result = _fields.GetTableForWeek(weekIndex);
            }
            return result;
        }

        public IEnumerable<MountainViewField> GetFields()
        {
            return _fields.GetFields();
        }

        private FieldWeekDictionary _fields;
        private MountainViewField _previousQuarterField;
    }
}