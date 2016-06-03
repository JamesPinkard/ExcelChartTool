using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace OpenXMLTools
{
    public class WeekCumulativeRecord : CumulativeRecord
    {
        public override int Week { get; }
        public double WeeklyFlowRate { get; private set; }
        public double CumalativeVolume { get; private set; }
        public double NetVolume { get; private set; }

        public WeekCumulativeRecord(int week, double weeklyFlowRate, double cumalativeVolume, double netVolume)
        {
            Week = week;
            WeeklyFlowRate = weeklyFlowRate;
            CumalativeVolume = cumalativeVolume;
            NetVolume = netVolume;
        }

        public override List<Cell> GetHeader()
        {
            return new List<Cell>()
            {                
                new Cell { CellValue = new CellValue("Week Number"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Sunday of week"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Net Volume"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Cumalative Volume"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Weekly Flow Rate"), StyleIndex = (UInt32Value)297U, DataType = CellValues.String}
            };
        }

        public DateTime GetSundayOfWeek()
        {
            var copyDate = _startDay;
            return copyDate.AddDays(TimeSpan.FromDays(((Week - 1) * 7)).Days).Date;
        }

        public override List<Cell> GetFilledCells()
        {
            return new List<Cell>()
            {                
                new Cell { CellValue = new CellValue(Week.ToString()), StyleIndex = (UInt32Value)294U, DataType = CellValues.Number},
                new Cell { CellValue = new CellValue(GetSundayOfWeek().ToOADate().ToString()), StyleIndex = (UInt32Value)316U },
                new Cell { CellValue = new CellValue(NetVolume.ToString()), StyleIndex = (UInt32Value)324U, DataType = CellValues.Number },
                new Cell { CellValue = new CellValue(CumalativeVolume.ToString()), StyleIndex = (UInt32Value)336U, DataType = CellValues.Number },
                new Cell { CellValue = new CellValue(WeeklyFlowRate.ToString()), StyleIndex = (UInt32Value)323U, DataType = CellValues.Number }
            };
        }

        private static readonly DateTime _startDay = new DateTime(2015, 6, 14);

    }
}
