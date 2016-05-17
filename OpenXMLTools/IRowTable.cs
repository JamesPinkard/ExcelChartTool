using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace OpenXMLTools
{
    public interface IRowTable
    {
        IEnumerable<Row> GetRows();
    }
}