using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class FirstQuarterState : QuarterStateBase
    {
        public override IQuarterState NextQuarter()
        {
            return new SecondQuarterState();
        }

        protected override bool VerifyMonth(int month)
        {
            if (month >= 4)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
    }
}
