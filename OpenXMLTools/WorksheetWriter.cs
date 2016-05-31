using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class WorksheetWriter : IRecordWriter
    {
        public WorksheetWriter(WorksheetPart worksheetPart, WorkbookPart workbookPart):this(worksheetPart, workbookPart, new CellReference(1, 1))
        {

        }

        public WorksheetWriter(WorksheetPart worksheetPart, WorkbookPart workbookPart, CellReference startingCell)
        {
            _worksheetPart = worksheetPart;             
            _columnNames = GetColumnNames();
            _workbookPart = workbookPart;
            _startingCell = startingCell;
            _rowIndex = (uint)startingCell.RowIndex -1;
            _columnIndex = startingCell.ColumnIndex -1;
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
            WorksheetPart replacementPart = _workbookPart.AddNewPart<WorksheetPart>();
            _reader = OpenXmlReader.Create(_worksheetPart);
            _writer = OpenXmlWriter.Create(replacementPart);
            
            while (_reader.Read())
            {
                if (_reader.ElementType == typeof(SheetData))
                {
                    if (_reader.IsEndElement)
                    {
                        continue;
                    }
                    _writer.WriteStartElement(new SheetData());

                    WriteRow(records.First().GetHeaderCells());
                    foreach (var sr in records)
                    {
                        WriteRow(sr.ToCells());
                    }
                    _writer.WriteEndElement();         
                }
                else
                {
                    if (_reader.IsStartElement)
                    {
                        _writer.WriteStartElement(_reader);
                    }
                    else if (_reader.IsEndElement)
                    {
                        _writer.WriteEndElement();
                    }
                }
            }

            _reader.Close();
            _writer.Close();

            string originalSheetId = _workbookPart.GetIdOfPart(_worksheetPart);
            string replacementSheetId = _workbookPart.GetIdOfPart(replacementPart);

            Sheet sheet = _workbookPart.Workbook.Descendants<Sheet>()
                .Where(s => s.Id.Value.Equals(originalSheetId)).First();
            sheet.Id.Value = replacementSheetId;
            _workbookPart.DeletePart(_worksheetPart);
            _worksheetPart = replacementPart;
        }

        private void WriteRow(IEnumerable<Cell> cells)
        {            
            _writer.WriteStartElement(GetNextRow());
            foreach (Cell c in cells)
            {
                c.CellReference = GetNextCellReference();
                _writer.WriteElement(c);                
            }
            _writer.WriteEndElement();
            ResetCellIndex();
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
        private OpenXmlWriter _writer;
        private WorkbookPart _workbookPart;
        private OpenXmlReader _reader;
        private CellReference _startingCell;
    }
}
