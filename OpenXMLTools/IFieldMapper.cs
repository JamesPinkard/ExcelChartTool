using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public interface IFieldMapper
    {
        bool TryParse(Row row, out MountainViewField field);
    }
}
