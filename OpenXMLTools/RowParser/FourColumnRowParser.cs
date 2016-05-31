using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    class FourColumnRowParser : IRowParser
    {
        public FourColumnRowParser(CellParser cellParser)
        {
            _cellParser = cellParser;
        }

        public bool ValidateElementCount(Row row)
        {
            return row.ChildElements.Count >= 4;
        }
        
        public string ParseStationFromRow(Row row)
        {
            Cell stationCell = (Cell)row.ChildElements.ElementAt(0);
            string station = _cellParser.ParseStation(stationCell);
            return station;
        }
        
        public DateTime ParseTimeFromRow(Row row)
        {
            Cell dateCell = (Cell)row.ChildElements.ElementAt(1);
            DateTime date = _cellParser.ParseTime(dateCell);

            Cell timeCell = (Cell)row.ChildElements.ElementAt(2);
            if (isValid(timeCell))
            {
                DateTime time = _cellParser.ParseTime(timeCell);                
                date = date.AddHours(time.Hour);
                date = date.AddMinutes(time.Minute);
            }

            return date;
        }

        public int ParseMeasurementFromRow(Row row)
        {
            Cell totalizerReadingCell = (Cell)row.ChildElements.ElementAt(3);
            if (isValid(totalizerReadingCell))
            {
                int totalizerReading = _cellParser.ParseInt(totalizerReadingCell);
                return totalizerReading;
            }
            else return 0;
        }

        private static bool isValid(Cell cell)
        {
            return cell.CellValue != null;
        }



        CellParser _cellParser;
    }
}
