using System;

namespace OpenXMLTools
{
    public class ThirdQuarterState : QuarterStateBase
    {
        public override IQuarterState NextQuarter()
        {
            return new FourthQuarterState();
        }

        protected override bool VerifyMonth(int month)
        {
            if (month >= 10)
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