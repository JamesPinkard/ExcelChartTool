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
            var header = GetHeader();
            return header;
        }

        public IEnumerable<Cell> ToCells()
        {
            var filledCells = GetFilledCells();

            int rowLength = 10;

            if (filledCells.Count < rowLength)
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
        public abstract List<Cell> GetHeader();

        
    }
}
