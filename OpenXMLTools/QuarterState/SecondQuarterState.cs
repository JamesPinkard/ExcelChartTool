using System;

namespace OpenXMLTools
{
    public class SecondQuarterState : QuarterStateBase
    {
        public override IQuarterState NextQuarter()
        {
            return new ThirdQuarterState();
        }

        protected override bool VerifyMonth(int month)
        {
            if (month >= 7)
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