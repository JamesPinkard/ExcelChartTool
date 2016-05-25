using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class IndividualMeasurementRecord : MeasurementRecord
    {
        public IndividualMeasurementRecord(string stationName, DateTime measureTime, int totalizerReading, double cumulativeTime, int cumulativeFlow)
        {
            StationName = stationName;
            MeasureTime = measureTime;
            TotalizerReading = totalizerReading;
            CumulativeTime = cumulativeTime;
            CumulativeFlow = cumulativeFlow;
        }
        public string StationName { get; set; }
        public DateTime MeasureTime { get; set; }
        public int TotalizerReading { get; set; }
        public double CumulativeTime { get; set; }
        public int CumulativeFlow { get; set; }

        public override List<Cell> GetFilledCells()
        {
            return new List<Cell>()
            {
                new Cell { CellValue = new CellValue(StationName), DataType = CellValues.String},
                new Cell { CellValue = new CellValue(MeasureTime.ToOADate().ToString()), StyleIndex = 5},
                new Cell { CellValue = new CellValue(TotalizerReading.ToString()), DataType = CellValues.Number},
                new Cell { CellValue = new CellValue(CumulativeTime.ToString()), DataType = CellValues.Number },
                new Cell { CellValue = new CellValue(CumulativeFlow.ToString()), DataType = CellValues.Number}
            };
        }

        public override List<Cell> GetHeader()
        {
            return new List<Cell>()
            {
                new Cell { CellValue = new CellValue("Station"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Date and Time"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Total Flow\n (gal)"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Cumulative Time\n (min)"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Cumulative Flow\n (gal)"), DataType = CellValues.String}
            };
        }
    }
}
