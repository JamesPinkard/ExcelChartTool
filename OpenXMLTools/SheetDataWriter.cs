using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class SheetDataWriter : IRecordWriter
    {
        public SheetDataWriter(WorksheetPart worksheetPart, WorkbookPart workbookPart):this(worksheetPart, workbookPart, new CellReference(1, 1))
        {

        }

        public SheetDataWriter(WorksheetPart worksheetPart, WorkbookPart workbookPart, CellReference startingCell)
        {
            _worksheetPart = worksheetPart;
            _columnNames = GetColumnNames();
            _workbookPart = workbookPart;
            _startingCell = startingCell;
            _rowIndex = (uint)startingCell.RowIndex - 1;
            _columnIndex = startingCell.ColumnIndex - 1;
        }

        public CellReference GetStartingCell()
        {
            return _startingCell;
        }

        public string GetSheetName()
        {
            string sheetId = _workbookPart.GetIdOfPart(_worksheetPart);
            Sheet sheet = _workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Id.Value.Equals(sheetId)).First();
            return sheet.Name;
        }

        public WorksheetPart GetWorksheetPart()
        {
            return _worksheetPart;
        }

        public void WriteRecords(IEnumerable<IRecord> records)
        {
            SheetData sheetData = new SheetData();

            Row headerRow = WriteRow(records.First().GetHeaderCells());
            sheetData.Append(headerRow);

            foreach (var r in records)
            {
                Row row = WriteRow(r.ToCells());
                sheetData.Append(row);
            }

            _worksheetPart.Worksheet.Append(sheetData);
        }

        private Row WriteRow(IEnumerable<Cell> cells)
        {
            Row row = GetNextRow();
            foreach (Cell c in cells)
            {
                c.CellReference = GetNextCellReference();
                row.Append(c);
            }            
            ResetCellIndex();
            return row;
        }

        private void ResetCellIndex()
        {
            _columnIndex = _startingCell.ColumnIndex - 1;
        }

        private Row GetNextRow()
        {
            _rowIndex++;
            return new Row() { RowIndex = _rowIndex };
        }

        private StringValue GetNextCellReference()
        {
            string cellReference = _columnNames[_columnIndex] + _rowIndex;
            _columnIndex++;
            return new StringValue(cellReference);
        }

        private List<String> GetColumnNames()
        {
            string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return alphabet.ToList().Select(c => c.ToString()).ToList();
        }

        private uint _rowIndex;
        private int _columnIndex;
        private readonly List<string> _columnNames;
        private WorksheetPart _worksheetPart;
        private WorkbookPart _workbookPart;
        private CellReference _startingCell;
    }
}
