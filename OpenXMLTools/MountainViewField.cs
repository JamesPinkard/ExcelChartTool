using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class MountainViewField
    {
        public string StationName { get; set; }
        public DateTime MeasureTime { get; set; }
        public int TotalizerReading { get; set; }
        
        public bool ValidDate()
        {
            return MeasureTime >= _startDay;
        }

        public bool ValidReading()
        {
            return TotalizerReading > 0;
        }

        public int GetWeek()
        {
            var duration = MeasureTime - _startDay;
            return (duration.Days / 7) +1;
        }

        public static DateTime GetSundayOfWeek(int week)
        {
            var copyDate = _startDay;
            return copyDate.AddDays(TimeSpan.FromDays(((week - 1) * 7)).Days).Date;
        }
        
        private static readonly DateTime _startDay = new DateTime(2015, 6, 14);

        public override string ToString()
        {
            return string.Format("Reading({0}, {1}, {2})", StationName, MeasureTime, TotalizerReading);
        }
    }
}
