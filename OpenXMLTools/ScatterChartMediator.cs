using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class ScatterChartMediator
    {
        public ScatterChartMediator(ScatterChart scatterChart)
        {
            _scatterChart = scatterChart;
        }

        public ScatterChartSeries GetSeries(string seriesName)
        {
            var series = _scatterChart.Elements<ScatterChartSeries>().Where(s => s.SeriesText.NumericValue.Text == seriesName).First();
            return series;            
        }

        public ScatterChartSeriesFormatter GetSeriesFormatter(string seriesName)
        {
            var series = GetSeries(seriesName);
            return new ScatterChartSeriesFormatter(series);
        }

        ScatterChart _scatterChart;
    }
}
