using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    class ThreeColumnRowParser : IRowParser
    {

        public ThreeColumnRowParser(CellParser cellParser)
        {
            _cellParser = cellParser;             
        }

        public bool ValidateElementCount(Row row)
        {
            return row.ChildElements.Count >= 3;
        }

        public string ParseStationFromRow(Row row)
        {
            Cell stationCell = (Cell)row.ChildElements.ElementAt(0);
            string station = _cellParser.ParseStation(stationCell);
            return station;
        }

        public DateTime ParseTimeFromRow(Row row)
        {
            Cell measureTimeCell = (Cell)row.ChildElements.ElementAt(1);
            DateTime measureTime = _cellParser.ParseTime(measureTimeCell);
            return measureTime;
        }

        public int ParseMeasurementFromRow(Row row)
        {
            Cell totalizerReadingCell = (Cell)row.ChildElements.ElementAt(2);
            int totalizerReading = _cellParser.ParseInt(totalizerReadingCell);
            return totalizerReading;
        }

        CellParser _cellParser;
    }
}
