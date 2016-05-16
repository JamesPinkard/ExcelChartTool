using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class RecordByWeek : IRecord
    {
        public int Week { get; private set; }
        public double TotalFlowRate { get; private set; }
        public double CumalativeVolume { get; private set; }

        public RecordByWeek(int week, double totalRate, double cumalativeVolume)
        {
            Week = week;
            TotalFlowRate = totalRate;
            CumalativeVolume = cumalativeVolume;
        }

        public IEnumerable<Cell> GetHeaderCells()
        {
            return new Cell[]
            {                
                new Cell { CellValue = new CellValue("Week Number"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Sunday of week"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Total Flow Rate"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Cumalative Volume"), DataType = CellValues.String}
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
                new Cell { CellValue = new CellValue(TotalFlowRate.ToString()), DataType = CellValues.Number },
                new Cell { CellValue = new CellValue(CumalativeVolume.ToString()), DataType = CellValues.Number }
            };
        }

        private static readonly DateTime _startDay = new DateTime(2015, 6, 14);

    }
}
