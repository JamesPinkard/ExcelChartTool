using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class FieldMapper : IFieldMapper
    {
        public FieldMapper(IRowParser rowParser)
        {
            _rowParser = rowParser;
        }

        public bool TryParse(Row row, out MountainViewField field)
        {
            field = new MountainViewField();
            if (!_rowParser.ValidateElementCount(row))
            {
                return false;
            }

            field.StationName = _rowParser.ParseStationFromRow(row);
            field.MeasureTime = _rowParser.ParseTimeFromRow(row);
            field.TotalizerReading = _rowParser.ParseMeasurementFromRow(row);

            if (!field.ValidDate() || !field.ValidReading())
            {
                return false;
            }
            return true;
        }

        IRowParser _rowParser;
    }
}
