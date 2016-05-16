using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

using DrawingChart = DocumentFormat.OpenXml.Drawing.Charts.Chart;
using DrawingValues = DocumentFormat.OpenXml.Drawing.Charts.Values;

namespace OpenXMLTools
{
    public class ScatterChartSeriesFormatter
    {
        public ScatterChartSeriesFormatter(ScatterChartSeries series)
        {
            _series = series;
        }

        public void SetValues(string worksheetName, int xColumn, int yColumn, int rowCount)
        {
            var xValues = _series.GetFirstChild<XValues>();
            var xRefs = xValues.Elements<NumberReference>().First();
            var beginningXCell = new CellReference(2, xColumn);
            var xFormula = string.Format("('{0}'!${1}${2}:${1}${3})", worksheetName, beginningXCell.ColumnName, beginningXCell.RowIndex, rowCount + 1);
            xRefs.Formula = new Formula(xFormula);
            xRefs.NumberingCache = new NumberingCache();

            var yValues = _series.GetFirstChild<YValues>();
            var yRefs = yValues.Elements<NumberReference>().First();
            var beginningYCell = new CellReference(2, yColumn);
            var yFormula = string.Format("('{0}'!${1}${2}:${1}${3})", worksheetName, beginningYCell.ColumnName, beginningYCell.RowIndex, rowCount + 1);
            yRefs.Formula = new Formula(yFormula);
            yRefs.NumberingCache = new NumberingCache();
        }



        private ScatterChartSeries _series;
    }
}
