using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class WorksheetRange
    {
        public int Top { get; set; }
        public int Bottom { get; set; }
        public string SheetName { get; set; }

        public WorksheetRange(string sheetName, int top, int bottom)
        {
            Top = top;
            Bottom = bottom;
            SheetName = sheetName;
        }

        public string GetColumnFormula(int column)
        {
            var columns = GetColumnNames();
            var columnIndex = column - 1;

            var columnName = columns[columnIndex];
            var formula = string.Format("('{0}'!${1}${2}:${1}${3})", SheetName, columnName, Top, Bottom);
            return formula;
        }

        private List<String> GetColumnNames()
        {
            string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return alphabet.ToList().Select(c => c.ToString()).ToList();
        }
    }
}
