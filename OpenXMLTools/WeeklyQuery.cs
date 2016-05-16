using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class WeeklyQuery
    {
        public string StationName
        {
            get
            {
                return _stationName;
            }
        }

        public WeeklyQuery(string stationName)
        {
            this._stationName = stationName;
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
            var fields = _weeklyMeasurements[weekIndex];
            var orderedFields = fields.OrderBy(f => f.MeasureTime);

            MountainViewField firstField;
            MountainViewField lastField;

            firstField = orderedFields.First();
            lastField = orderedFields.Last();
            
            if (firstField == lastField)
            {
                int alternativeIndex = weekIndex - 1;
                bool hasNoPriorIndex = false;
                while (!Contains(alternativeIndex))
                {
                    alternativeIndex--;
                    if (alternativeIndex < 0)
                    {
                        hasNoPriorIndex = true;
                        break;
                    }
                }
                if (hasNoPriorIndex)
                {
                    while(!Contains(alternativeIndex) && alternativeIndex != weekIndex)
                    {
                        alternativeIndex++;
                        if (alternativeIndex > _weeklyMeasurements.Count())
                        {                            
                            return 0;
                        }
                    }
                }

                var alternativeField = _weeklyMeasurements[alternativeIndex].OrderBy(f => f.MeasureTime);

                if (hasNoPriorIndex)
                {
                    firstField = orderedFields.First();
                    lastField = alternativeField.First();
                }
                else
                {
                    firstField = alternativeField.Last();
                    lastField = orderedFields.Last();
                }
            }            

            var duration = lastField.MeasureTime - firstField.MeasureTime;
            var flow = lastField.TotalizerReading - firstField.TotalizerReading;

            var rate = flow / duration.TotalMinutes;

            return rate;
        }


        readonly Dictionary<int, List<MountainViewField>> _weeklyMeasurements = new Dictionary<int, List<MountainViewField>>();
        readonly string _stationName;
    }
}
