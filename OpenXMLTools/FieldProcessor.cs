using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class FieldProcessor
    {
        public FieldProcessor(IRowTable table, IFieldParser parser)
        {
            _table = table;
            _parser = parser;
        }

        public IEnumerable<MountainViewField> ProcessFields()
        {
            var rows = _table.GetRows();
            var fields = _parser.Parse(rows);
            return fields;
        }

        private IFieldParser _parser;
        private readonly IRowTable _table;
    }
}
