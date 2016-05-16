using System;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public interface IRowParser
    {
        bool ValidateElementCount(Row row);
        string ParseStationFromRow(Row row);
        DateTime ParseTimeFromRow(Row row);
        int ParseMeasurementFromRow(Row row);
    }
}
