using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace OpenXMLTools
{
    public class SimpleFieldParser : IFieldParser
    {
        public SimpleFieldParser(IFieldMapper fieldMapper)
        {
            _fieldMapper = fieldMapper;
        }

        public IEnumerable<MountainViewField> Parse(IEnumerable<Row> rows)
        {
            List<MountainViewField> fieldList = new List<MountainViewField>();

            var bodyRows = rows.Where(r => r.RowIndex.Value > 1);
            foreach (Row br in bodyRows)
            {
                MountainViewField field;

                if (_fieldMapper.TryParse(br, out field))
                {
                    fieldList.Add(field);
                }
                else
                {
                    continue;
                }                
            }

            return fieldList;
        }

        IFieldMapper _fieldMapper;
    }
}