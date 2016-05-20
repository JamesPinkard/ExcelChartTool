using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    class UniqueWeekFieldQuery
    {
        internal IEnumerable<int> GetUniqueWeekIndices(IEnumerable<MountainViewField> fields)
        {
            HashSet<int> weekIndexes = new HashSet<int>();

            foreach (var f in fields)
            {
                weekIndexes.Add(f.GetWeek());
            }

            return weekIndexes;
        }
    }
}
