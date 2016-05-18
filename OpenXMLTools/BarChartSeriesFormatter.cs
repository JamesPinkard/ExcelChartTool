using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLTools
{
    public class BarChartSeriesFormatter : ISeriesFormatter
    {
        public BarChartSeriesFormatter(BarChartSeries series)
        {
            _series = series;
        }

        public void SetValues(string worksheetName, int xColumn, int yColumn, int rowCount)
        {
            var xValues = _series.GetFirstChild<CategoryAxisData>();
            var xRefs = xValues.Elements<NumberReference>().First();
            var beginningXCell = new CellReference(2, xColumn);
            var xFormula = string.Format("('{0}'!${1}${2}:${1}${3})", worksheetName, beginningXCell.ColumnName, beginningXCell.RowIndex, rowCount + 1);
            xRefs.Formula = new Formula(xFormula);
            xRefs.NumberingCache = new NumberingCache();

            var yValues = _series.GetFirstChild<Values>();
            var yRefs = yValues.Elements<NumberReference>().First();
            var beginningYCell = new CellReference(2, yColumn);
            var yFormula = string.Format("('{0}'!${1}${2}:${1}${3})", worksheetName, beginningYCell.ColumnName, beginningYCell.RowIndex, rowCount + 1);
            yRefs.Formula = new Formula(yFormula);
            yRefs.NumberingCache = new NumberingCache();
        }

        public void SetSeriesFormula(string xRangeFormula, string yRangeFormula)
        {
            SetXFormula(xRangeFormula);
            SetYFormula(yRangeFormula);
        }

        public void SetXFormula(string xRangeFormula)
        {
            var xValues = _series.GetFirstChild<CategoryAxisData>();
            var xRefs = xValues.Elements<NumberReference>().First();
            xRefs.Formula = new Formula(xRangeFormula);
            xRefs.NumberingCache = new NumberingCache();
        }

        public void SetYFormula(string yRangeFormula)
        {
            var yValues = _series.GetFirstChild<Values>();
            var yRefs = yValues.Elements<NumberReference>().First();
            yRefs.Formula = new Formula(yRangeFormula);
            yRefs.NumberingCache = new NumberingCache();
        }

        public void SetSeriesTitle(string title)
        {
            _series.SeriesText = new SeriesText(new NumericValue() { Text = title });
        }

        private BarChartSeries _series;
    }
}
