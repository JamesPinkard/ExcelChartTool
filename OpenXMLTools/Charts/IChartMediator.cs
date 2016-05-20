using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public interface IChartMediator
    {
        bool HasSeries(string seriesName);
        ISeriesFormatter GetSeriesFormatter(string seriesName);
    }
}
