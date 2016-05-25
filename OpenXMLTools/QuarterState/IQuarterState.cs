using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public interface IQuarterState
    {
        bool VerifyQuarter(MountainViewField field);
        IQuarterState NextQuarter();
    }
}
