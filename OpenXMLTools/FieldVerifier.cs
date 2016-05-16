using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    internal class FieldVerifier
    {
        public FieldVerifier(Dictionary<int, List<MountainViewField>> weeklyMeasurements)
        {
            _weeklyMeasurements = weeklyMeasurements;
        }


        public int GetAlternativeIndex(int weekIndex)
        {
            int? alternativeIndex = PriorIndex(weekIndex);
            if (alternativeIndex == null) alternativeIndex = AdvancedIndex(weekIndex);
            if (alternativeIndex == null) alternativeIndex = weekIndex;
            return (int)alternativeIndex;
        }

        public int? PriorIndex(int weekIndex)
        {
            int alternativeIndex = weekIndex - 1;
            while (!Contains(alternativeIndex))
            {
                alternativeIndex--;
                if (alternativeIndex < 0)
                {
                    return null;                    
                }
            }

            return alternativeIndex;
        }


        public int? AdvancedIndex(int weekIndex)
        {
            int alternativeIndex = weekIndex + 1;
            while (InvalidIndex(weekIndex, alternativeIndex))
            {
                alternativeIndex++;
                if (alternativeIndex > _weeklyMeasurements.Count())
                {
                    return null;
                }
            }

            return alternativeIndex;
        }


        private bool InvalidIndex(int weekIndex, int alternativeIndex)
        {
            return !Contains(alternativeIndex) && alternativeIndex != weekIndex;
        }

        public bool Contains(int weekIndex)
        {
            return _weeklyMeasurements.ContainsKey(weekIndex);
        }
        readonly Dictionary<int, List<MountainViewField>> _weeklyMeasurements = new Dictionary<int, List<MountainViewField>>();
    }
}
