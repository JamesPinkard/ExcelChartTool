using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLTools
{
    class ChartsheetMaker
    {
        private readonly SheetGenerator _sheetGenerator;
        private readonly OpenXmlPartGenerator _openXmlPartGenerator;
        private readonly IChartsheetPartBuilder _sheetPartBuilder;

        public ChartsheetMaker(SheetGenerator sheetGenerator, OpenXmlPartGenerator openXmlPartGenerator, IChartsheetPartBuilder sheetPartBuilder)
        {
            _sheetGenerator = sheetGenerator;
            _openXmlPartGenerator = openXmlPartGenerator;
            _sheetPartBuilder = sheetPartBuilder;
        }

        public void MakeChartSheet(string name, string id)
        {
            var cumulativeChart = _openXmlPartGenerator.AddNewPart<ChartsheetPart>(id);
            var cumulativeVolumeChartSheet = _sheetGenerator.CreateSheet(name, _openXmlPartGenerator.GetIdOfPart(cumulativeChart));
            _sheetPartBuilder.BuildSheetPart(cumulativeChart);
        }

    }
}
