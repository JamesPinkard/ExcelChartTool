using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class CumulativeRecord : IRecord
    {
        public int Week { get; private set; }
        public double CumulativeFlowRate { get; private set; }

        public IEnumerable<Cell> GetHeaderCells()
        {
            return new Cell[]
            {
                new Cell { CellValue = new CellValue("Week Number"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Sunday of week"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Cumulative Volume"), DataType = CellValues.String}
            };
        }

        public DateTime GetSundayOfWeek()
        {
            var copyDate = _startDay;
            return copyDate.AddDays(TimeSpan.FromDays(((Week - 1) * 7)).Days).Date;
        }

        public IEnumerable<Cell> ToCells()
        {
            return new Cell[]
            {
                new Cell { CellValue = new CellValue(Week.ToString()), DataType = CellValues.Number},
                new Cell { CellValue = new CellValue(GetSundayOfWeek().ToOADate().ToString()), StyleIndex = 5},
                new Cell { CellValue = new CellValue(CumulativeFlowRate.ToString()), DataType = CellValues.Number }
            };
        }

        private static readonly DateTime _startDay = new DateTime(2015, 6, 14);
    }
}