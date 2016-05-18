using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class BarChartMediator
    {
        public BarChartMediator(BarChart scatterChart)
        {
            _scatterChart = scatterChart;
        }

        public BarChartSeries GetSeries(string seriesName)
        {
            var series = _scatterChart.Elements<BarChartSeries>().Where(s => s.SeriesText.NumericValue.Text == seriesName).First();
            return series;
        }

        public BarChartSeriesFormatter GetSeriesFormatter(string seriesName)
        {
            var series = GetSeries(seriesName);
            return new BarChartSeriesFormatter(series);
        }

        BarChart _scatterChart;
    }
}
