using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public abstract class QuarterStateBase : IQuarterState
    {       

        public bool VerifyQuarter(MountainViewField field)
        {
            if (!field.ValidDate())
            {
                throw new ArgumentException("Field must have valid Date");
            }

            return VerifyMonth(field.MeasureTime.Month);
        }

        protected abstract bool VerifyMonth(int month);

        public abstract IQuarterState NextQuarter();
    }
}
