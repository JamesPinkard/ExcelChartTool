using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLTools
{
    internal interface IChartsheetPartBuilder
    {
        void BuildSheetPart(ChartsheetPart part);
    }
}