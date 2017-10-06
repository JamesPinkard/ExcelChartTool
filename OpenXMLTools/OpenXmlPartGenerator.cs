using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    public class OpenXmlPartGenerator
    {
        private readonly WorkbookPart _workbookPart;

        public OpenXmlPartGenerator(WorkbookPart workbookPart)
        {
            if (workbookPart == null)
                throw new ArgumentException("workbookPart");

            _workbookPart = workbookPart;
        }

        public T AddNewPart<T>(string id) where T : OpenXmlPart, IFixedContentTypePart
        {
            return _workbookPart.AddNewPart<T>(id);
        }

        public string GetIdOfPart(OpenXmlPart part)
        {
            return _workbookPart.GetIdOfPart(part);
        }
    }
}
