using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class NamedCumulativeRecord : CumulativeRecord
    {
        public NamedCumulativeRecord(CumulativeRecord cumulativeRecord, string name)
        {
            _cumulativeRecord = cumulativeRecord;
            _name = name;
        }
        public override int Week
        {
            get
            {
                return _cumulativeRecord.Week;
            }
        }

        public override List<Cell> GetFilledCells()
        {
            List<Cell> cellList = new List<Cell>();
            cellList.Add(new Cell { CellValue = new CellValue(_name.ToString()), DataType = CellValues.String });
            cellList.AddRange(_cumulativeRecord.GetFilledCells());
            
            return cellList;
        }

        public override List<Cell> GetHeader()
        {
            List<Cell> cellList = new List<Cell>();
            cellList.Add(new Cell { CellValue = new CellValue("Station Name"), DataType = CellValues.String });
            cellList.AddRange(_cumulativeRecord.GetHeader());
            return cellList;
        }

        private CumulativeRecord _cumulativeRecord;
        private string _name;
    }
}
