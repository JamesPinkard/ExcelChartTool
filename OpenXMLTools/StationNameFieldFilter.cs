using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class StationNameFieldFilter : IFieldFilter
    {
        public StationNameFieldFilter(string stationName)
        {
            _stationNames.Add(stationName);
        }

        public StationNameFieldFilter(IEnumerable<string> stationNames)
        {
            _stationNames.AddRange(stationNames);
        }

        public IEnumerable<MountainViewField> Filter(IEnumerable<MountainViewField> fields)
        {
            // Get all fields that belong to stations listed in station names
            var filteredFields = fields.Where(f => _stationNames.Contains(f.StationName));
            return filteredFields;
        }

        private List<string> _stationNames = new List<string>();
    }
}
