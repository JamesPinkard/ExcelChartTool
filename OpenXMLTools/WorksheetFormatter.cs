using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class WorksheetFormatter
    {
        public WorksheetFormatter(WorksheetPart worksheetPart)
        {
            _worksheetPart = worksheetPart;
        }

        public void FormatSheet()
        {
            var worksheet = _worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var rows = sheetData.Descendants<Row>();
            _mergeCellRanges = new List<CellRange>();

            _columnIndex = 0;

            IEnumerable<OpenXmlElement> firstRowCells = rows.First().ChildElements;
            foreach (Cell cell in firstRowCells)
            {
                SetNewCellRange(cell);
                _columnIndex++;
            }

            resetColumnIndex();

            foreach (var row in rows.Skip(1))
            {
                var rowCells = (IEnumerable<OpenXmlElement>)row.ChildElements;
                foreach (Cell cell in rowCells)
                {
                    var cellRange = _mergeDictionary[_columnIndex];
                    if (cell.CellValue == null)
                    {
                        cellRange.ExtendRange(1, 0);
                    }
                    else if (cell.CellValue != null && _mergeDictionary[_columnIndex].IsValidRange())
                    {
                        _mergeCellRanges.Add(cellRange);
                        SetNewCellRange(cell);
                    }
                    else
                    {
                        cellRange.MoveRange(1, 0);
                    }

                    _columnIndex++;
                }
                resetColumnIndex();
            }

            var lastRowCells = (IEnumerable<OpenXmlElement>)rows.Last().ChildElements;
            foreach (Cell cell in lastRowCells)
            {
                if (cell.CellValue == null)
                {
                    var cellRange = _mergeDictionary[_columnIndex];
                    _mergeCellRanges.Add(cellRange);
                    SetNewCellRange(cell);
                }
                _columnIndex++;
            }

            var mergeCells = new MergeCells() { Count = (uint)_mergeCellRanges.Count() };
            foreach (var range in _mergeCellRanges)
            {
                var mergeCell = new MergeCell() { Reference = range.GetRangeReference() };
                mergeCells.Append(mergeCell);
            }

            //var oldMergeList = _worksheetPart.Worksheet.GetFirstChild<MergeCells>();
            //_worksheetPart.Worksheet.ReplaceChild<MergeCells>(mergeCells, oldMergeList);
            _worksheetPart.Worksheet.Append(mergeCells);

        }

        private void SetNewCellRange(Cell cell)
        {
            _mergeDictionary[_columnIndex] = new CellRange(cell.CellReference.Value, cell.CellReference.Value);
        }

        private void resetColumnIndex()
        {
            _columnIndex = 0;
        }

        int _columnIndex = 0;
        private WorksheetPart _worksheetPart;
        private Dictionary<int, CellRange> _mergeDictionary = new Dictionary<int, CellRange>();
        private List<CellRange> _mergeCellRanges;
    }
}
