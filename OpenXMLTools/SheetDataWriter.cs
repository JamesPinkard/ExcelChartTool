using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class SheetDataWriter : IRecordWriter
    {
        public string GetSheetName()
        {
            throw new NotImplementedException();
        }

        public CellReference GetStartingCell()
        {
            throw new NotImplementedException();
        }

        public void WriteRecords(IEnumerable<IRecord> records)
        {
            throw new NotImplementedException();
        }
    }
}
