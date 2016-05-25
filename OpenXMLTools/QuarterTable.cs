using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class QuarterTable
    {
        public QuarterTable(FieldWeekDictionary fields)
        {
            _fields = fields;
        }

        public double GetAverageWeeklyFlowRate()
        {
            var quarterFields = GetFields();
            var firstField = quarterFields.First();
            var lastField = quarterFields.Last();


            var duration = (lastField.MeasureTime - firstField.MeasureTime).TotalMinutes;
            var flow = lastField.TotalizerReading - firstField.TotalizerReading;
            var rate = flow / duration;

            return rate;
        }

        public WeekTable GetTableForWeek(int weekIndex)
        {
            return _fields.GetTableForWeek(weekIndex);
        }

        public IEnumerable<MountainViewField> GetFields()
        {
            return _fields.GetFields();
        }

        private FieldWeekDictionary _fields;
    }
}