using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class CellParser
    {
        public CellParser(List<SharedStringItem> sharedStringList)
        {
            _sharedStringList = sharedStringList;
        }

        public int ParseInt(Cell totalizerReadingCell)
        {
            if (totalizerReadingCell.CellValue.Text == null)
            {
                return 0;
            }
            var totalizerString = totalizerReadingCell.CellValue.Text;
            return int.Parse(totalizerString);
        }

        public DateTime ParseTime(Cell measureTimeCell)
        {
            var timeString = measureTimeCell.CellValue.Text;
            var timeDouble = Double.Parse(timeString);
            return DateTime.FromOADate(timeDouble);
        }

        public string ParseStation(Cell stationCell)
        {
            if (stationCell.DataType == "s")
            {

                var sharedIndex = int.Parse(stationCell.CellValue.Text);
                return _sharedStringList[sharedIndex].Text.Text;

            }
            else
            {
                return stationCell.CellValue.Text;
            }

        }

        private List<SharedStringItem> _sharedStringList;
    }
}
