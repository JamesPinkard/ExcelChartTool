using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;


namespace OpenXMLTools
{
    class ExcelExtensionListGenerator
    {
        // Creates an WorkbookExtensionList instance and adds its children.
        public WorkbookExtensionList GenerateWorkbookExtensionList()
        {
            WorkbookExtensionList workbookExtensionList1 = new WorkbookExtensionList();

            WorkbookExtension workbookExtension1 = new WorkbookExtension() { Uri = "{140A7094-0E35-4892-8432-C4D2E57EDEB5}" };
            workbookExtension1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.WorkbookProperties workbookProperties1 = new X15.WorkbookProperties() { ChartTrackingReferenceBase = true };

            workbookExtension1.Append(workbookProperties1);

            workbookExtensionList1.Append(workbookExtension1);
            return workbookExtensionList1;
        }


    }
}
