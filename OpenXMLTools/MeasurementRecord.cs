using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public abstract class MeasurementRecord : IRecord
    {
        public IEnumerable<Cell> GetHeaderCells()
        {
            return new Cell[]
            {
                new Cell { CellValue = new CellValue("Station"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Date and Time"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Total Flow\n (gal)"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Cumulative Time\n (min)"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Cumulative Flow\n (gal)"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Week #\n"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Weekly Cumulative\n Time"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Weekly Cumulative\n Flow"), DataType = CellValues.String},
                new Cell { CellValue = new CellValue("Weekly Cumulative\n Flow Rate\n (GPM)"), DataType = CellValues.String}
            };
        }

        public IEnumerable<Cell> ToCells()
        {
            var filledCells = GetFilledCells();

            int rowLength = 9;

            if (filledCells.Count < 9)
            {
            int padding = rowLength - filledCells.Count();

                for (int i = 0; i < padding; i++)
                {
                    filledCells.Add(new Cell());
                }
            }

            return filledCells;
        }


        public abstract List<Cell> GetFilledCells();
    }
}
