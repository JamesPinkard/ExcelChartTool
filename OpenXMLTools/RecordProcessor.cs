using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class RecordProcessor
    {
        public RecordProcessor(IEnumerable<MountainViewField> fields, IRecordQuery fieldQuery, IFieldFilter fieldFilter)
        {
            _fields = fields;
            _fieldQuery = fieldQuery;
            _fieldFilter = fieldFilter;
        }

        public IEnumerable<IRecord> ProcessRecords()
        {
            var filteredFields = _fieldFilter.Filter(_fields);
            var records = _fieldQuery.Query(filteredFields);
            return records;
        }        

        IEnumerable<MountainViewField> _fields;
        IRecordQuery _fieldQuery;
        IFieldFilter _fieldFilter;
    }
}
