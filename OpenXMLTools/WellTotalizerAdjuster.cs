using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    class WellTotalizerAdjuster
    {

        public void AdjustReading(IEnumerable<MountainViewField> fields, string wellName, DateTime startTime, int addedValue)
        {
            var extractionFields = fields.Where(f => f.StationName.Equals(wellName))
                .Where(f => f.MeasureTime > startTime)
                .OrderBy(f => f.MeasureTime);                

            foreach(var f in extractionFields)
            {
                f.TotalizerReading += addedValue;
            }
        }

    }
}
