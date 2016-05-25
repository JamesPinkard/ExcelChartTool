using System;

namespace OpenXMLTools
{
    public class FourthQuarterState : QuarterStateBase
    {
        public override IQuarterState NextQuarter()
        {
            return new FirstQuarterState();
        }

        protected override bool VerifyMonth(int month)
        {
            if (month < 10)
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