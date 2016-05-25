using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class WeekTable
    {
        public WeekTable(int weekIndex, IEnumerable<MountainViewField> fields, MountainViewField firstField, MountainViewField lastField)
        {
            _weekIndex = weekIndex;
            _fields = fields;
            _firstField = firstField;
            _lastField = lastField;
        }

        public int GetWeek()
        {
            return _weekIndex;
        }

        public double GetCumulativeTime()
        {
            var duration = _lastField.MeasureTime - _firstField.MeasureTime;
            return duration.TotalMinutes;
        }

        public double GetAverageWeeklyFlowRate()
        {
            var duration = GetCumulativeTime();
            if (duration == 0)
            {
                return 0;
            }

            var flow = _lastField.TotalizerReading - _firstField.TotalizerReading;
            var rate = flow / duration;            
            return rate;
        }

        public double GetNetVolume()
        {            
            var totalFlow = _lastField.TotalizerReading - _firstField.TotalizerReading;

            return totalFlow;
        }

        public IEnumerable<MountainViewField> GetFieldsForWeek()
        {
            return _fields;
        }
        private int _weekIndex;
        private IEnumerable<MountainViewField> _fields;
        private MountainViewField _firstField;
        private MountainViewField _lastField;
    }
}
