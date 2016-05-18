using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public interface ISeriesFormatter
    {
        void SetSeriesFormula(string xRangeFormula, string yRangeFormula);
        void SetXFormula(string xRangeFormula);
        void SetYFormula(string yRangeFormula);
        void SetSeriesTitle(string title);
    }
}
