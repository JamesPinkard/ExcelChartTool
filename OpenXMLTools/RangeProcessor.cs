using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class RangeProcessor
    {
        public RangeProcessor(WorksheetWriter worksheetWriter)
        {
            _worksheetWriter = worksheetWriter;
        }

        public WorksheetRange AddRecords(IEnumerable<IRecord> records)
        {
            if (records.Count() == 0)
            {
                throw new ArgumentException("records must not be empty");
            }
           _records.AddRange(records);

            int topOfRange = _bottomOfRange + 1;
            _bottomOfRange += records.Count();

            return new WorksheetRange(_worksheetWriter.GetSheetName(), topOfRange, _bottomOfRange);
        }

        public void WriteRecords()
        {
            _worksheetWriter.WriteRecords(_records);
        }

        // because of header bottom of range sta
        int _bottomOfRange = 1;        
        List<IRecord> _records = new List<IRecord>();
        WorksheetWriter _worksheetWriter;
    }
}
