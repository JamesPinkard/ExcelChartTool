using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class RangeProcessor
    {
        public RangeProcessor(IRecordWriter recordWriter)
        {
            _recordWriter = recordWriter;
            _bottomOfRange = _recordWriter.GetStartingCell().RowIndex;
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

            return new WorksheetRange(_recordWriter.GetSheetName(), topOfRange, _bottomOfRange);
        }

        public void WriteRecords()
        {
            _recordWriter.WriteRecords(_records);
        }

        // because of header bottom of range sta
        int _bottomOfRange;        
        List<IRecord> _records = new List<IRecord>();
        IRecordWriter _recordWriter;
    }
}
