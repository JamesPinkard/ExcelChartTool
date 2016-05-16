using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    class FieldParserFactory
    {
        public FieldParserFactory(List<SharedStringItem> sharedStringList)
        {
            _sharedStringList = sharedStringList;
        }

        public IFieldParser MakeParser()
        {
            CellParser cellParser = new CellParser(_sharedStringList);
            IRowParser rowParser = new FourColumnRowParser(cellParser);
            IFieldMapper fieldMapper = new FieldMapper(rowParser);
            IFieldParser parser = new SimpleFieldParser(fieldMapper);
            return parser;
        }

        private List<SharedStringItem> _sharedStringList;
    }
}
