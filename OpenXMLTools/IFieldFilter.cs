using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace OpenXMLTools
{
    public interface IFieldFilter
    {
        IEnumerable<MountainViewField> Filter(IEnumerable<MountainViewField> fields);
    }
}