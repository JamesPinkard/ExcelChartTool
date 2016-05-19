using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class ScatterChartMediator : IChartMediator
    {
        public ScatterChartMediator(ScatterChart scatterChart)
        {
            _scatterChart = scatterChart;
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
            return new ScatterChartSeriesFormatter(series);
        }

        private ScatterChartSeries GetSeries(string seriesName)
        {
            var series = _scatterChart.Elements<ScatterChartSeries>().Where(s => s.SeriesText.NumericValue.Text == seriesName).FirstOrDefault();
            return series;            
        }

        ScatterChart _scatterChart;
    }
}
