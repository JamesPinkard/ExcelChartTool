using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ClosedXML;
using ClosedXML.Excel;

namespace OpenXMLTools
{
    public class RecoredWriter : IDisposable
    {
        public RecoredWriter(IXLWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public void WriteRecords(IEnumerable<StationRecord> records)
        {
            WriteHeader();
            foreach (var row in records)
            {                
                WriteRow(row);
            }
        }

        private void WriteHeader()
        {
            _rowIndex++;

            using (var row = _worksheet.Row(_rowIndex))
            {
                
                var stationCell = GetNextCell(row);
                stationCell.Value = "StationName";

                var weekCell = GetNextCell(row);
                weekCell.Value = "Week Number";

                var dateCell = GetNextCell(row);
                dateCell.Value = "Sunday of week";

                var rateCell = GetNextCell(row);
                rateCell.Value = "Total Flow Rate";

            }
             ResetCellIndex();

        }

        private void WriteRow(StationRecord record)
        {
            _rowIndex++;

            using (var row = _worksheet.Row(_rowIndex))
            {

                var stationCell = GetNextCell(row);
                stationCell.Value = record.StationName;

                var weekCell = GetNextCell(row);
                weekCell.Value = record.Week;

                var dateCell = GetNextCell(row);
                dateCell.Value = record.GetSundayOfWeek();
                dateCell.Style.DateFormat.Format = "mm-dd-yyyy";

                var rateCell = GetNextCell(row);
                rateCell.Value = record.FlowRateForWeek;
                rateCell.Style.NumberFormat.Format = "##0.000";
            }

            ResetCellIndex();
        }

        private void ResetCellIndex()
        {
            _columnIndex = 0;
        }

        private IXLCell GetNextCell(IXLRow row)
        {
            _columnIndex++;
            return row.Cell(_columnIndex);
        }

        public void Dispose()
        {
            _worksheet.Dispose();
        }

        private int _rowIndex = 0;
        private int _columnIndex = 0;
        private IXLWorksheet _worksheet;
    }
}
