using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class RecordProcessor
    {
        public RecordProcessor(IEnumerable<MountainViewField> fields, IRecordQuery fieldQuery): this(fields, fieldQuery, null)
        {

        }

        public RecordProcessor(IEnumerable<MountainViewField> fields, IRecordQuery fieldQuery, IFieldFilter fieldFilter)
        {
            _fields = fields;
            _fieldQuery = fieldQuery;
            _fieldFilter = fieldFilter;
        }

        

        public IFieldFilter FieldFilter
        {
            get { return _fieldFilter; }
            set { _fieldFilter = value; }
        }


        public IEnumerable<IRecord> ProcessRecords()
        {
            IEnumerable<MountainViewField> processedFields;
            if (_fieldFilter != null)
            {
                processedFields = _fieldFilter.Filter(_fields);
            }
            else
            {
                processedFields = _fields;
            }
            var records = _fieldQuery.Query(processedFields);
            return records;
        }        

        IEnumerable<MountainViewField> _fields;
        IRecordQuery _fieldQuery;
        IFieldFilter _fieldFilter;
    }
}
