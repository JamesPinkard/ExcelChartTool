using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class BarChartMediator : IChartMediator
    {
        public BarChartMediator(BarChart barChart)
        {
            _barChart = barChart;
        }

        public bool HasSeries(string seriesName)
        {
            var series = GetSeries(seriesName);
            if (series == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public ISeriesFormatter GetSeriesFormatter(string seriesName)
        {
            var series = GetSeries(seriesName);
            return new BarChartSeriesFormatter(series);
        }

        private BarChartSeries GetSeries(string seriesName)
        {
            var series = _barChart.Elements<BarChartSeries>().Where(s => s.SeriesText.NumericValue.Text == seriesName).FirstOrDefault();
            return series;
        }

        BarChart _barChart;
    }
}
