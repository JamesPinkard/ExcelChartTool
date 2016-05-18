using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

using DrawingChart = DocumentFormat.OpenXml.Drawing.Charts.Chart;
using DrawingValues = DocumentFormat.OpenXml.Drawing.Charts.Values;

using OpenXMLTools;

namespace OpenXMLTools
{
    public class WorksheetWriter
    {
        public WorksheetWriter(WorksheetPart worksheetPart, WorkbookPart workbookPart)
        {
            _worksheetPart = worksheetPart;             
            _columnNames = GetColumnNames();
            _workbookPart = workbookPart;
        }

        public string GetSheetName()
        {
            string sheetId = _workbookPart.GetIdOfPart(_worksheetPart);
            Sheet sheet = _workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Id.Value.Equals(sheetId)).First();
            return sheet.Name;
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

        public void WriteRow(IEnumerable<Cell> cells)
        {            
            _writer.WriteStartElement(GetNextRow());
            foreach (Cell c in cells)
            {                
                _writer.WriteElement(c);                
            }
            _writer.WriteEndElement();
            ResetCellIndex();
        }

        private void ResetCellIndex()
        {
            _columnIndex = 0;
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

        private uint _rowIndex = 0;
        private int _columnIndex = 0;
        private readonly List<string> _columnNames;
        private WorksheetPart _worksheetPart;
        private OpenXmlWriter _writer;
        private WorkbookPart _workbookPart;
        private OpenXmlReader _reader;
    }
}
