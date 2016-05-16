using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Globalization;

namespace OpenXMLTools
{
    public class StationRecord : IRecord
    {
        public string StationName { get; private set; }
        public int Week { get; private set; }
        public double FlowRateForWeek { get; private set; }

        public StationRecord(string stationName, int week, double flowRate)
        {
            StationName = stationName;
            Week = week;
            FlowRateForWeek = flowRate;
        }

        public DateTime GetSundayOfWeek()
        {
            var copyDate = _startDay;
            return copyDate.AddDays(TimeSpan.FromDays(((Week - 1) * 7)).Days).Date;
        }

        public override string ToString()
        {
            return String.Format("{0}, Week {1}, {2}, {3}", StationName, Week, GetSundayOfWeek(), FlowRateForWeek);
        }

        public IEnumerable<Cell> ToCells()
        {
            return new Cell[]
            {
                new Cell { InlineString = new InlineString() {Text = new Text(StationName) }, DataType = CellValues.InlineString},
                new Cell { CellValue = new CellValue(Week.ToString()), DataType = CellValues.Number},
                new Cell { CellValue = new CellValue(GetSundayOfWeek().ToOADate().ToString()), StyleIndex = 5},
                new Cell { CellValue = new CellValue(FlowRateForWeek.ToString()), DataType = CellValues.Number }
            };
        }

        public IEnumerable<Cell> GetHeaderCells()
        {
            return new Cell[]
            {
                new Cell { CellValue = new CellValue("StationName"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Week Number"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Sunday of week"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Total Flow Rate"), DataType = CellValues.String}
            };
        }

        public static Row HeaderRow { get; set; }
        private static readonly DateTime _startDay = new DateTime(2015, 6, 14);        
    }
}
