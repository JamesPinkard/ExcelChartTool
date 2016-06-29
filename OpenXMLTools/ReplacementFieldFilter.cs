using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class ReplacementFieldFilter : IFieldFilter
    {
        public ReplacementFieldFilter(StationNameFieldFilter fieldFilter, string stationName)
        {
            _fieldFilter = fieldFilter;
            _stationName = stationName;
        }

        public IEnumerable<MountainViewField> Filter(IEnumerable<MountainViewField> fields)
        {
            var filteredFields = _fieldFilter.Filter(fields);
            var replacementFields = fields.Where(f => f.StationName == _stationName);
            var replacedWeeks = replacementFields.Select(r => r.GetWeek());
            var existingFields = filteredFields.Where(f => !replacedWeeks.Contains(f.GetWeek()));
            var fieldList = existingFields.ToList();
            fieldList.AddRange(replacementFields);
            return fieldList;
        }

        private StationNameFieldFilter _fieldFilter;
        private string _stationName;
    }
}
