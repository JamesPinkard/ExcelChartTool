using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public interface IRecord
    {
        IEnumerable<Cell> ToCells();
        IEnumerable<Cell> GetHeaderCells(); 
    }
}