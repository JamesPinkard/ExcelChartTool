using DocumentFormat.OpenXml.Packaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using Cs = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLTools
{
    class FlowRateChartGenerator : IChartsheetPartBuilder
    {
        // Adds child parts and generates content of the specified part.
        public void BuildSheetPart(ChartsheetPart part)
        {
            DrawingsPart drawingsPart1 = part.AddNewPart<DrawingsPart>("rId2");
            GenerateDrawingsPart1Content(drawingsPart1);

            ChartPart chartPart1 = drawingsPart1.AddNewPart<ChartPart>("rId1");
            GenerateChartPart1Content(chartPart1);

            ChartColorStylePart chartColorStylePart1 = chartPart1.AddNewPart<ChartColorStylePart>("rId2");
            GenerateChartColorStylePart1Content(chartColorStylePart1);

            ChartStylePart chartStylePart1 = chartPart1.AddNewPart<ChartStylePart>("rId1");
            GenerateChartStylePart1Content(chartStylePart1);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1 = part.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart1Content(spreadsheetPrinterSettingsPart1);

            GeneratePartContent(part);

        }

        // Generates content of drawingsPart1.
        private void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.AbsoluteAnchor absoluteAnchor1 = new Xdr.AbsoluteAnchor();
            Xdr.Position position1 = new Xdr.Position() { X = 0L, Y = 0L };
            Xdr.Extent extent1 = new Xdr.Extent() { Cx = 8667750L, Cy = 6296025L };

            Xdr.GraphicFrame graphicFrame1 = new Xdr.GraphicFrame() { Macro = "" };

            Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new Xdr.NonVisualGraphicFrameProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Chart 1" };

            Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Xdr.NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoGrouping = true };

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);

            Xdr.Transform transform1 = new Xdr.Transform();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };

            transform1.Append(offset1);
            transform1.Append(extents1);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            C.ChartReference chartReference1 = new C.ChartReference() { Id = "rId1" };
            chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartReference1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData1.Append(chartReference1);

            graphic1.Append(graphicData1);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            absoluteAnchor1.Append(position1);
            absoluteAnchor1.Append(extent1);
            absoluteAnchor1.Append(graphicFrame1);
            absoluteAnchor1.Append(clientData1);

            worksheetDrawing1.Append(absoluteAnchor1);

            drawingsPart1.WorksheetDrawing = worksheetDrawing1;
        }

        // Generates content of chartPart1.
        private void GenerateChartPart1Content(ChartPart chartPart1)
        {
            C.ChartSpace chartSpace1 = new C.ChartSpace();
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            C.Date1904 date19041 = new C.Date1904() { Val = false };
            C.EditingLanguage editingLanguage1 = new C.EditingLanguage() { Val = "en-US" };
            C.RoundedCorners roundedCorners1 = new C.RoundedCorners() { Val = false };

            AlternateContent alternateContent1 = new AlternateContent();
            alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "c14" };
            alternateContentChoice1.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
            C14.Style style1 = new C14.Style() { Val = 102 };

            alternateContentChoice1.Append(style1);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
            C.Style style2 = new C.Style() { Val = 2 };

            alternateContentFallback1.Append(style2);

            alternateContent1.Append(alternateContentChoice1);
            alternateContent1.Append(alternateContentFallback1);

            C.Chart chart1 = new C.Chart();
            C.AutoTitleDeleted autoTitleDeleted1 = new C.AutoTitleDeleted() { Val = true };

            C.PlotArea plotArea1 = new C.PlotArea();
            C.Layout layout1 = new C.Layout();

            C.BarChart barChart1 = new C.BarChart();
            C.BarDirection barDirection1 = new C.BarDirection() { Val = C.BarDirectionValues.Column };
            C.BarGrouping barGrouping1 = new C.BarGrouping() { Val = C.BarGroupingValues.Clustered };
            C.VaryColors varyColors1 = new C.VaryColors() { Val = false };

            C.BarChartSeries barChartSeries1 = new C.BarChartSeries();
            C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

            C.SeriesText seriesText1 = new C.SeriesText();
            C.NumericValue numericValue1 = new C.NumericValue();
            numericValue1.Text = "Injection";

            seriesText1.Append(numericValue1);

            C.ChartShapeProperties chartShapeProperties1 = new C.ChartShapeProperties();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            solidFill1.Append(schemeColor1);

            A.Outline outline1 = new A.Outline();
            A.NoFill noFill1 = new A.NoFill();

            outline1.Append(noFill1);
            A.EffectList effectList1 = new A.EffectList();

            chartShapeProperties1.Append(solidFill1);
            chartShapeProperties1.Append(outline1);
            chartShapeProperties1.Append(effectList1);
            C.InvertIfNegative invertIfNegative1 = new C.InvertIfNegative() { Val = false };

            C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

            C.NumberReference numberReference1 = new C.NumberReference();
            C.Formula formula1 = new C.Formula();
            formula1.Text = "(\'[1]Weekly Inj_Rates\'!$H$15,\'[1]Weekly Inj_Rates\'!$H$34,\'[1]Weekly Inj_Rates\'!$H$53,\'[1]Weekly Inj_Rates\'!$H$76,\'[1]Weekly Inj_Rates\'!$H$91,\'[1]Weekly Inj_Rates\'!$H$106,\'[1]Weekly Inj_Rates\'!$H$129,\'[1]Weekly Inj_Rates\'!$H$144,\'[1]Weekly Inj_Rates\'!$H$159,\'[1]Weekly Inj_Rates\'!$H$173,\'[1]Weekly Inj_Rates\'!$H$184,\'[1]Weekly Inj_Rates\'!$H$199,\'[1]Weekly Inj_Rates\'!$H$214,\'[1]Weekly Inj_Rates\'!$H$225,\'[1]Weekly Inj_Rates\'!$H$236,\'[1]Weekly Inj_Rates\'!$H$273,\'[1]Weekly Inj_Rates\'!$H$288,\'[1]Weekly Inj_Rates\'!$H$299,\'[1]Weekly Inj_Rates\'!$H$314,\'[1]Weekly Inj_Rates\'!$H$329,\'[1]Weekly Inj_Rates\'!$H$344,\'[1]Weekly Inj_Rates\'!$H$359,\'[1]Weekly Inj_Rates\'!$H$370,\'[1]Weekly Inj_Rates\'!$H$381,\'[1]Weekly Inj_Rates\'!$H$396,\'[1]Weekly Inj_Rates\'!$H$411,\'[1]Weekly Inj_Rates\'!$H$422,\'[1]Weekly Inj_Rates\'!$H$433,\'[1]Weekly Inj_Rates\'!$H$457,\'[1]Weekly Inj_Rates\'!$H$474,\'[1]Weekly Inj_Rates\'!$H$490,\'[1]Weekly Inj_Rates\'!$H$502,\'[1]Weekly Inj_Rates\'!$H$512,\'[1]Weekly Inj_Rates\'!$H$524,\'[1]Weekly Inj_Rates\'!$H$540,\'[1]Weekly Inj_Rates\'!$H$556,\'[1]Weekly Inj_Rates\'!$H$572,\'[1]Weekly Inj_Rates\'!$H$584,\'[1]Weekly Inj_Rates\'!$H$596,\'[1]Weekly Inj_Rates\'!$H$608)";

            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode();
            formatCode1.Text = "d\\-mmm\\-yy";
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)40U };

            C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue2 = new C.NumericValue();
            numericValue2.Text = "42170";

            numericPoint1.Append(numericValue2);

            C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue3 = new C.NumericValue();
            numericValue3.Text = "42177";

            numericPoint2.Append(numericValue3);

            C.NumericPoint numericPoint3 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue4 = new C.NumericValue();
            numericValue4.Text = "42184";

            numericPoint3.Append(numericValue4);

            C.NumericPoint numericPoint4 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue5 = new C.NumericValue();
            numericValue5.Text = "42198";

            numericPoint4.Append(numericValue5);

            C.NumericPoint numericPoint5 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue6 = new C.NumericValue();
            numericValue6.Text = "42205";

            numericPoint5.Append(numericValue6);

            C.NumericPoint numericPoint6 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue7 = new C.NumericValue();
            numericValue7.Text = "42212";

            numericPoint6.Append(numericValue7);

            C.NumericPoint numericPoint7 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue8 = new C.NumericValue();
            numericValue8.Text = "42219";

            numericPoint7.Append(numericValue8);

            C.NumericPoint numericPoint8 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue9 = new C.NumericValue();
            numericValue9.Text = "42226";

            numericPoint8.Append(numericValue9);

            C.NumericPoint numericPoint9 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue10 = new C.NumericValue();
            numericValue10.Text = "42233";

            numericPoint9.Append(numericValue10);

            C.NumericPoint numericPoint10 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue11 = new C.NumericValue();
            numericValue11.Text = "42240.697916666664";

            numericPoint10.Append(numericValue11);

            C.NumericPoint numericPoint11 = new C.NumericPoint() { Index = (UInt32Value)10U };
            C.NumericValue numericValue12 = new C.NumericValue();
            numericValue12.Text = "42249.337500000001";

            numericPoint11.Append(numericValue12);

            C.NumericPoint numericPoint12 = new C.NumericPoint() { Index = (UInt32Value)11U };
            C.NumericValue numericValue13 = new C.NumericValue();
            numericValue13.Text = "42254.347222222219";

            numericPoint12.Append(numericValue13);

            C.NumericPoint numericPoint13 = new C.NumericPoint() { Index = (UInt32Value)12U };
            C.NumericValue numericValue14 = new C.NumericValue();
            numericValue14.Text = "42261.321527777778";

            numericPoint13.Append(numericValue14);

            C.NumericPoint numericPoint14 = new C.NumericPoint() { Index = (UInt32Value)13U };
            C.NumericValue numericValue15 = new C.NumericValue();
            numericValue15.Text = "42268.665972222225";

            numericPoint14.Append(numericValue15);

            C.NumericPoint numericPoint15 = new C.NumericPoint() { Index = (UInt32Value)14U };
            C.NumericValue numericValue16 = new C.NumericValue();
            numericValue16.Text = "42275.436111111114";

            numericPoint15.Append(numericValue16);

            C.NumericPoint numericPoint16 = new C.NumericPoint() { Index = (UInt32Value)15U };
            C.NumericValue numericValue17 = new C.NumericValue();
            numericValue17.Text = "42279.411805555559";

            numericPoint16.Append(numericValue17);

            C.NumericPoint numericPoint17 = new C.NumericPoint() { Index = (UInt32Value)16U };
            C.NumericValue numericValue18 = new C.NumericValue();
            numericValue18.Text = "42289.338888888888";

            numericPoint17.Append(numericValue18);

            C.NumericPoint numericPoint18 = new C.NumericPoint() { Index = (UInt32Value)17U };
            C.NumericValue numericValue19 = new C.NumericValue();
            numericValue19.Text = "42296.396527777775";

            numericPoint18.Append(numericValue19);

            C.NumericPoint numericPoint19 = new C.NumericPoint() { Index = (UInt32Value)18U };
            C.NumericValue numericValue20 = new C.NumericValue();
            numericValue20.Text = "42303.357638888891";

            numericPoint19.Append(numericValue20);

            C.NumericPoint numericPoint20 = new C.NumericPoint() { Index = (UInt32Value)19U };
            C.NumericValue numericValue21 = new C.NumericValue();
            numericValue21.Text = "42310.379861111112";

            numericPoint20.Append(numericValue21);

            C.NumericPoint numericPoint21 = new C.NumericPoint() { Index = (UInt32Value)20U };
            C.NumericValue numericValue22 = new C.NumericValue();
            numericValue22.Text = "42317.574999999997";

            numericPoint21.Append(numericValue22);

            C.NumericPoint numericPoint22 = new C.NumericPoint() { Index = (UInt32Value)21U };
            C.NumericValue numericValue23 = new C.NumericValue();
            numericValue23.Text = "42324.416666666664";

            numericPoint22.Append(numericValue23);

            C.NumericPoint numericPoint23 = new C.NumericPoint() { Index = (UInt32Value)22U };
            C.NumericValue numericValue24 = new C.NumericValue();
            numericValue24.Text = "42331.348611111112";

            numericPoint23.Append(numericValue24);

            C.NumericPoint numericPoint24 = new C.NumericPoint() { Index = (UInt32Value)23U };
            C.NumericValue numericValue25 = new C.NumericValue();
            numericValue25.Text = "42338.527777777781";

            numericPoint24.Append(numericValue25);

            C.NumericPoint numericPoint25 = new C.NumericPoint() { Index = (UInt32Value)24U };
            C.NumericValue numericValue26 = new C.NumericValue();
            numericValue26.Text = "42345.645833333336";

            numericPoint25.Append(numericValue26);

            C.NumericPoint numericPoint26 = new C.NumericPoint() { Index = (UInt32Value)25U };
            C.NumericValue numericValue27 = new C.NumericValue();
            numericValue27.Text = "42352.334027777775";

            numericPoint26.Append(numericValue27);

            C.NumericPoint numericPoint27 = new C.NumericPoint() { Index = (UInt32Value)26U };
            C.NumericValue numericValue28 = new C.NumericValue();
            numericValue28.Text = "42359.444444444445";

            numericPoint27.Append(numericValue28);

            C.NumericPoint numericPoint28 = new C.NumericPoint() { Index = (UInt32Value)27U };
            C.NumericValue numericValue29 = new C.NumericValue();
            numericValue29.Text = "42366.636805555558";

            numericPoint28.Append(numericValue29);

            C.NumericPoint numericPoint29 = new C.NumericPoint() { Index = (UInt32Value)28U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue30 = new C.NumericValue();
            numericValue30.Text = "42373.343055555553";

            numericPoint29.Append(numericValue30);

            C.NumericPoint numericPoint30 = new C.NumericPoint() { Index = (UInt32Value)29U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue31 = new C.NumericValue();
            numericValue31.Text = "42380.372916666667";

            numericPoint30.Append(numericValue31);

            C.NumericPoint numericPoint31 = new C.NumericPoint() { Index = (UInt32Value)30U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue32 = new C.NumericValue();
            numericValue32.Text = "42387.590277777781";

            numericPoint31.Append(numericValue32);

            C.NumericPoint numericPoint32 = new C.NumericPoint() { Index = (UInt32Value)31U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue33 = new C.NumericValue();
            numericValue33.Text = "42394.630555555559";

            numericPoint32.Append(numericValue33);

            C.NumericPoint numericPoint33 = new C.NumericPoint() { Index = (UInt32Value)32U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue34 = new C.NumericValue();
            numericValue34.Text = "42412.60833333333";

            numericPoint33.Append(numericValue34);

            C.NumericPoint numericPoint34 = new C.NumericPoint() { Index = (UInt32Value)33U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue35 = new C.NumericValue();
            numericValue35.Text = "42416.474999999999";

            numericPoint34.Append(numericValue35);

            C.NumericPoint numericPoint35 = new C.NumericPoint() { Index = (UInt32Value)34U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue36 = new C.NumericValue();
            numericValue36.Text = "42422.419444444444";

            numericPoint35.Append(numericValue36);

            C.NumericPoint numericPoint36 = new C.NumericPoint() { Index = (UInt32Value)35U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue37 = new C.NumericValue();
            numericValue37.Text = "42429.363888888889";

            numericPoint36.Append(numericValue37);

            C.NumericPoint numericPoint37 = new C.NumericPoint() { Index = (UInt32Value)36U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue38 = new C.NumericValue();
            numericValue38.Text = "42436.354166666664";

            numericPoint37.Append(numericValue38);

            C.NumericPoint numericPoint38 = new C.NumericPoint() { Index = (UInt32Value)37U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue39 = new C.NumericValue();
            numericValue39.Text = "42443.369444444441";

            numericPoint38.Append(numericValue39);

            C.NumericPoint numericPoint39 = new C.NumericPoint() { Index = (UInt32Value)38U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue40 = new C.NumericValue();
            numericValue40.Text = "42450.362500000003";

            numericPoint39.Append(numericValue40);

            C.NumericPoint numericPoint40 = new C.NumericPoint() { Index = (UInt32Value)39U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue41 = new C.NumericValue();
            numericValue41.Text = "42457.690972222219";

            numericPoint40.Append(numericValue41);

            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount1);
            numberingCache1.Append(numericPoint1);
            numberingCache1.Append(numericPoint2);
            numberingCache1.Append(numericPoint3);
            numberingCache1.Append(numericPoint4);
            numberingCache1.Append(numericPoint5);
            numberingCache1.Append(numericPoint6);
            numberingCache1.Append(numericPoint7);
            numberingCache1.Append(numericPoint8);
            numberingCache1.Append(numericPoint9);
            numberingCache1.Append(numericPoint10);
            numberingCache1.Append(numericPoint11);
            numberingCache1.Append(numericPoint12);
            numberingCache1.Append(numericPoint13);
            numberingCache1.Append(numericPoint14);
            numberingCache1.Append(numericPoint15);
            numberingCache1.Append(numericPoint16);
            numberingCache1.Append(numericPoint17);
            numberingCache1.Append(numericPoint18);
            numberingCache1.Append(numericPoint19);
            numberingCache1.Append(numericPoint20);
            numberingCache1.Append(numericPoint21);
            numberingCache1.Append(numericPoint22);
            numberingCache1.Append(numericPoint23);
            numberingCache1.Append(numericPoint24);
            numberingCache1.Append(numericPoint25);
            numberingCache1.Append(numericPoint26);
            numberingCache1.Append(numericPoint27);
            numberingCache1.Append(numericPoint28);
            numberingCache1.Append(numericPoint29);
            numberingCache1.Append(numericPoint30);
            numberingCache1.Append(numericPoint31);
            numberingCache1.Append(numericPoint32);
            numberingCache1.Append(numericPoint33);
            numberingCache1.Append(numericPoint34);
            numberingCache1.Append(numericPoint35);
            numberingCache1.Append(numericPoint36);
            numberingCache1.Append(numericPoint37);
            numberingCache1.Append(numericPoint38);
            numberingCache1.Append(numericPoint39);
            numberingCache1.Append(numericPoint40);

            numberReference1.Append(formula1);
            numberReference1.Append(numberingCache1);

            categoryAxisData1.Append(numberReference1);

            C.Values values1 = new C.Values();

            C.NumberReference numberReference2 = new C.NumberReference();
            C.Formula formula2 = new C.Formula();
            formula2.Text = "(\'[1]Weekly Inj_Rates\'!$F$15,\'[1]Weekly Inj_Rates\'!$F$34,\'[1]Weekly Inj_Rates\'!$F$53,\'[1]Weekly Inj_Rates\'!$F$76,\'[1]Weekly Inj_Rates\'!$F$91,\'[1]Weekly Inj_Rates\'!$F$106,\'[1]Weekly Inj_Rates\'!$F$129,\'[1]Weekly Inj_Rates\'!$F$144,\'[1]Weekly Inj_Rates\'!$F$159,\'[1]Weekly Inj_Rates\'!$F$174,\'[1]Weekly Inj_Rates\'!$F$185,\'[1]Weekly Inj_Rates\'!$F$200,\'[1]Weekly Inj_Rates\'!$F$215,\'[1]Weekly Inj_Rates\'!$F$226,\'[1]Weekly Inj_Rates\'!$F$237,\'[1]Weekly Inj_Rates\'!$F$273,\'[1]Weekly Inj_Rates\'!$F$288,\'[1]Weekly Inj_Rates\'!$F$299,\'[1]Weekly Inj_Rates\'!$F$314,\'[1]Weekly Inj_Rates\'!$F$329,\'[1]Weekly Inj_Rates\'!$F$344,\'[1]Weekly Inj_Rates\'!$F$359,\'[1]Weekly Inj_Rates\'!$F$370,\'[1]Weekly Inj_Rates\'!$F$381,\'[1]Weekly Inj_Rates\'!$F$396,\'[1]Weekly Inj_Rates\'!$F$411,\'[1]Weekly Inj_Rates\'!$F$422,\'[1]Weekly Inj_Rates\'!$F$433,\'[1]Weekly Inj_Rates\'!$F$457,\'[1]Weekly Inj_Rates\'!$F$474,\'[1]Weekly Inj_Rates\'!$F$490,\'[1]Weekly Inj_Rates\'!$F$502,\'[1]Weekly Inj_Rates\'!$F$512,\'[1]Weekly Inj_Rates\'!$F$524,\'[1]Weekly Inj_Rates\'!$F$540,\'[1]Weekly Inj_Rates\'!$F$556,\'[1]Weekly Inj_Rates\'!$F$572,\'[1]Weekly Inj_Rates\'!$F$584,\'[1]Weekly Inj_Rates\'!$F$596,\'[1]Weekly Inj_Rates\'!$F$608)";

            C.NumberingCache numberingCache2 = new C.NumberingCache();
            C.FormatCode formatCode2 = new C.FormatCode();
            formatCode2.Text = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)";
            C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)40U };

            C.NumericPoint numericPoint41 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue42 = new C.NumericValue();
            numericValue42.Text = "2.7821275199704827";

            numericPoint41.Append(numericValue42);

            C.NumericPoint numericPoint42 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue43 = new C.NumericValue();
            numericValue43.Text = "8.2738590871807975";

            numericPoint42.Append(numericValue43);

            C.NumericPoint numericPoint43 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue44 = new C.NumericValue();
            numericValue44.Text = "15.010187636507489";

            numericPoint43.Append(numericValue44);

            C.NumericPoint numericPoint44 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue45 = new C.NumericValue();
            numericValue45.Text = "2.8873844598842138";

            numericPoint44.Append(numericValue45);

            C.NumericPoint numericPoint45 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue46 = new C.NumericValue();
            numericValue46.Text = "11.170580026137818";

            numericPoint45.Append(numericValue46);

            C.NumericPoint numericPoint46 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue47 = new C.NumericValue();
            numericValue47.Text = "10.179576465248022";

            numericPoint46.Append(numericValue47);

            C.NumericPoint numericPoint47 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue48 = new C.NumericValue();
            numericValue48.Text = "2.9035226872961455";

            numericPoint47.Append(numericValue48);

            C.NumericPoint numericPoint48 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue49 = new C.NumericValue();
            numericValue49.Text = "9.6505406439582639";

            numericPoint48.Append(numericValue49);

            C.NumericPoint numericPoint49 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue50 = new C.NumericValue();
            numericValue50.Text = "10.04225189703266";

            numericPoint49.Append(numericValue50);

            C.NumericPoint numericPoint50 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue51 = new C.NumericValue();
            numericValue51.Text = "3.7651240421730465";

            numericPoint50.Append(numericValue51);

            C.NumericPoint numericPoint51 = new C.NumericPoint() { Index = (UInt32Value)10U };
            C.NumericValue numericValue52 = new C.NumericValue();
            numericValue52.Text = "7.9878114225998189";

            numericPoint51.Append(numericValue52);

            C.NumericPoint numericPoint52 = new C.NumericPoint() { Index = (UInt32Value)11U };
            C.NumericValue numericValue53 = new C.NumericValue();
            numericValue53.Text = "12.573319233362334";

            numericPoint52.Append(numericValue53);

            C.NumericPoint numericPoint53 = new C.NumericPoint() { Index = (UInt32Value)12U };
            C.NumericValue numericValue54 = new C.NumericValue();
            numericValue54.Text = "10.800192189557835";

            numericPoint53.Append(numericValue54);

            C.NumericPoint numericPoint54 = new C.NumericPoint() { Index = (UInt32Value)13U };
            C.NumericValue numericValue55 = new C.NumericValue();
            numericValue55.Text = "11.112383850729181";

            numericPoint54.Append(numericValue55);

            C.NumericPoint numericPoint55 = new C.NumericPoint() { Index = (UInt32Value)14U };
            C.NumericValue numericValue56 = new C.NumericValue();
            numericValue56.Text = "9.9515304791015975";

            numericPoint55.Append(numericValue56);

            C.NumericPoint numericPoint56 = new C.NumericPoint() { Index = (UInt32Value)15U };
            C.NumericValue numericValue57 = new C.NumericValue();
            numericValue57.Text = "10.470052459141662";

            numericPoint56.Append(numericValue57);

            C.NumericPoint numericPoint57 = new C.NumericPoint() { Index = (UInt32Value)16U };
            C.NumericValue numericValue58 = new C.NumericValue();
            numericValue58.Text = "10.417876492873519";

            numericPoint57.Append(numericValue58);

            C.NumericPoint numericPoint58 = new C.NumericPoint() { Index = (UInt32Value)17U };
            C.NumericValue numericValue59 = new C.NumericValue();
            numericValue59.Text = "10.021841692803909";

            numericPoint58.Append(numericValue59);

            C.NumericPoint numericPoint59 = new C.NumericPoint() { Index = (UInt32Value)18U };
            C.NumericValue numericValue60 = new C.NumericValue();
            numericValue60.Text = "11.281883257078803";

            numericPoint59.Append(numericValue60);

            C.NumericPoint numericPoint60 = new C.NumericPoint() { Index = (UInt32Value)19U };
            C.NumericValue numericValue61 = new C.NumericValue();
            numericValue61.Text = "6.2454895291370374";

            numericPoint60.Append(numericValue61);

            C.NumericPoint numericPoint61 = new C.NumericPoint() { Index = (UInt32Value)20U };
            C.NumericValue numericValue62 = new C.NumericValue();
            numericValue62.Text = "8.2552420037894336";

            numericPoint61.Append(numericValue62);

            C.NumericPoint numericPoint62 = new C.NumericPoint() { Index = (UInt32Value)21U };
            C.NumericValue numericValue63 = new C.NumericValue();
            numericValue63.Text = "6.8309586186339732";

            numericPoint62.Append(numericValue63);

            C.NumericPoint numericPoint63 = new C.NumericPoint() { Index = (UInt32Value)22U };
            C.NumericValue numericValue64 = new C.NumericValue();
            numericValue64.Text = "8.3609436559622345";

            numericPoint63.Append(numericValue64);

            C.NumericPoint numericPoint64 = new C.NumericPoint() { Index = (UInt32Value)23U };
            C.NumericValue numericValue65 = new C.NumericValue();
            numericValue65.Text = "7.2161212072577552";

            numericPoint64.Append(numericValue65);

            C.NumericPoint numericPoint65 = new C.NumericPoint() { Index = (UInt32Value)24U };
            C.NumericValue numericValue66 = new C.NumericValue();
            numericValue66.Text = "10.017457407124313";

            numericPoint65.Append(numericValue66);

            C.NumericPoint numericPoint66 = new C.NumericPoint() { Index = (UInt32Value)25U };
            C.NumericValue numericValue67 = new C.NumericValue();
            numericValue67.Text = "10.329247811635588";

            numericPoint66.Append(numericValue67);

            C.NumericPoint numericPoint67 = new C.NumericPoint() { Index = (UInt32Value)26U };
            C.NumericValue numericValue68 = new C.NumericValue();
            numericValue68.Text = "9.6209627059695375";

            numericPoint67.Append(numericValue68);

            C.NumericPoint numericPoint68 = new C.NumericPoint() { Index = (UInt32Value)27U };
            C.NumericValue numericValue69 = new C.NumericValue();
            numericValue69.Text = "8.6400200693113618";

            numericPoint68.Append(numericValue69);

            C.NumericPoint numericPoint69 = new C.NumericPoint() { Index = (UInt32Value)28U };
            C.NumericValue numericValue70 = new C.NumericValue();
            numericValue70.Text = "8.7465973406725475";

            numericPoint69.Append(numericValue70);

            C.NumericPoint numericPoint70 = new C.NumericPoint() { Index = (UInt32Value)29U };
            C.NumericValue numericValue71 = new C.NumericValue();
            numericValue71.Text = "7.8992782515202009";

            numericPoint70.Append(numericValue71);

            C.NumericPoint numericPoint71 = new C.NumericPoint() { Index = (UInt32Value)30U };
            C.NumericValue numericValue72 = new C.NumericValue();
            numericValue72.Text = "9.6994729205628918";

            numericPoint71.Append(numericValue72);

            C.NumericPoint numericPoint72 = new C.NumericPoint() { Index = (UInt32Value)31U };
            C.NumericValue numericValue73 = new C.NumericValue();
            numericValue73.Text = "7.9441916008219877";

            numericPoint72.Append(numericValue73);

            C.NumericPoint numericPoint73 = new C.NumericPoint() { Index = (UInt32Value)32U };
            C.NumericValue numericValue74 = new C.NumericValue();
            numericValue74.Text = "1.9442490747946204";

            numericPoint73.Append(numericValue74);

            C.NumericPoint numericPoint74 = new C.NumericPoint() { Index = (UInt32Value)33U };
            C.NumericValue numericValue75 = new C.NumericValue();
            numericValue75.Text = "5.0652264007660497";

            numericPoint74.Append(numericValue75);

            C.NumericPoint numericPoint75 = new C.NumericPoint() { Index = (UInt32Value)34U };
            C.NumericValue numericValue76 = new C.NumericValue();
            numericValue76.Text = "8.6023794035489161";

            numericPoint75.Append(numericValue76);

            C.NumericPoint numericPoint76 = new C.NumericPoint() { Index = (UInt32Value)35U };
            C.NumericValue numericValue77 = new C.NumericValue();
            numericValue77.Text = "8.5930296352042888";

            numericPoint76.Append(numericValue77);

            C.NumericPoint numericPoint77 = new C.NumericPoint() { Index = (UInt32Value)36U };
            C.NumericValue numericValue78 = new C.NumericValue();
            numericValue78.Text = "9.7476274898294957";

            numericPoint77.Append(numericValue78);

            C.NumericPoint numericPoint78 = new C.NumericPoint() { Index = (UInt32Value)37U };
            C.NumericValue numericValue79 = new C.NumericValue();
            numericValue79.Text = "11.642572161952504";

            numericPoint78.Append(numericValue79);

            C.NumericPoint numericPoint79 = new C.NumericPoint() { Index = (UInt32Value)38U };
            C.NumericValue numericValue80 = new C.NumericValue();
            numericValue80.Text = "9.5059476906740557";

            numericPoint79.Append(numericValue80);

            C.NumericPoint numericPoint80 = new C.NumericPoint() { Index = (UInt32Value)39U };
            C.NumericValue numericValue81 = new C.NumericValue();
            numericValue81.Text = "10.488972485560243";

            numericPoint80.Append(numericValue81);

            numberingCache2.Append(formatCode2);
            numberingCache2.Append(pointCount2);
            numberingCache2.Append(numericPoint41);
            numberingCache2.Append(numericPoint42);
            numberingCache2.Append(numericPoint43);
            numberingCache2.Append(numericPoint44);
            numberingCache2.Append(numericPoint45);
            numberingCache2.Append(numericPoint46);
            numberingCache2.Append(numericPoint47);
            numberingCache2.Append(numericPoint48);
            numberingCache2.Append(numericPoint49);
            numberingCache2.Append(numericPoint50);
            numberingCache2.Append(numericPoint51);
            numberingCache2.Append(numericPoint52);
            numberingCache2.Append(numericPoint53);
            numberingCache2.Append(numericPoint54);
            numberingCache2.Append(numericPoint55);
            numberingCache2.Append(numericPoint56);
            numberingCache2.Append(numericPoint57);
            numberingCache2.Append(numericPoint58);
            numberingCache2.Append(numericPoint59);
            numberingCache2.Append(numericPoint60);
            numberingCache2.Append(numericPoint61);
            numberingCache2.Append(numericPoint62);
            numberingCache2.Append(numericPoint63);
            numberingCache2.Append(numericPoint64);
            numberingCache2.Append(numericPoint65);
            numberingCache2.Append(numericPoint66);
            numberingCache2.Append(numericPoint67);
            numberingCache2.Append(numericPoint68);
            numberingCache2.Append(numericPoint69);
            numberingCache2.Append(numericPoint70);
            numberingCache2.Append(numericPoint71);
            numberingCache2.Append(numericPoint72);
            numberingCache2.Append(numericPoint73);
            numberingCache2.Append(numericPoint74);
            numberingCache2.Append(numericPoint75);
            numberingCache2.Append(numericPoint76);
            numberingCache2.Append(numericPoint77);
            numberingCache2.Append(numericPoint78);
            numberingCache2.Append(numericPoint79);
            numberingCache2.Append(numericPoint80);

            numberReference2.Append(formula2);
            numberReference2.Append(numberingCache2);

            values1.Append(numberReference2);

            C.BarSerExtensionList barSerExtensionList1 = new C.BarSerExtensionList();
            barSerExtensionList1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");

            C.BarSerExtension barSerExtension1 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            barSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-5AC9-435B-A402-5E3C89F8E9A8}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            barSerExtension1.Append(openXmlUnknownElement1);

            barSerExtensionList1.Append(barSerExtension1);

            barChartSeries1.Append(index1);
            barChartSeries1.Append(order1);
            barChartSeries1.Append(seriesText1);
            barChartSeries1.Append(chartShapeProperties1);
            barChartSeries1.Append(invertIfNegative1);
            barChartSeries1.Append(categoryAxisData1);
            barChartSeries1.Append(values1);
            barChartSeries1.Append(barSerExtensionList1);

            C.BarChartSeries barChartSeries2 = new C.BarChartSeries();
            C.Index index2 = new C.Index() { Val = (UInt32Value)1U };
            C.Order order2 = new C.Order() { Val = (UInt32Value)1U };

            C.SeriesText seriesText2 = new C.SeriesText();
            C.NumericValue numericValue82 = new C.NumericValue();
            numericValue82.Text = "Extraction";

            seriesText2.Append(numericValue82);

            C.ChartShapeProperties chartShapeProperties2 = new C.ChartShapeProperties();

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };

            solidFill2.Append(schemeColor2);

            A.Outline outline2 = new A.Outline();
            A.NoFill noFill2 = new A.NoFill();

            outline2.Append(noFill2);
            A.EffectList effectList2 = new A.EffectList();

            chartShapeProperties2.Append(solidFill2);
            chartShapeProperties2.Append(outline2);
            chartShapeProperties2.Append(effectList2);
            C.InvertIfNegative invertIfNegative2 = new C.InvertIfNegative() { Val = false };

            C.CategoryAxisData categoryAxisData2 = new C.CategoryAxisData();

            C.NumberReference numberReference3 = new C.NumberReference();
            C.Formula formula3 = new C.Formula();
            formula3.Text = "(\'[1]Weekly Inj_Rates\'!$H$15,\'[1]Weekly Inj_Rates\'!$H$34,\'[1]Weekly Inj_Rates\'!$H$53,\'[1]Weekly Inj_Rates\'!$H$76,\'[1]Weekly Inj_Rates\'!$H$91,\'[1]Weekly Inj_Rates\'!$H$106,\'[1]Weekly Inj_Rates\'!$H$129,\'[1]Weekly Inj_Rates\'!$H$144,\'[1]Weekly Inj_Rates\'!$H$159,\'[1]Weekly Inj_Rates\'!$H$173,\'[1]Weekly Inj_Rates\'!$H$184,\'[1]Weekly Inj_Rates\'!$H$199,\'[1]Weekly Inj_Rates\'!$H$214,\'[1]Weekly Inj_Rates\'!$H$225,\'[1]Weekly Inj_Rates\'!$H$236,\'[1]Weekly Inj_Rates\'!$H$273,\'[1]Weekly Inj_Rates\'!$H$288,\'[1]Weekly Inj_Rates\'!$H$299,\'[1]Weekly Inj_Rates\'!$H$314,\'[1]Weekly Inj_Rates\'!$H$329,\'[1]Weekly Inj_Rates\'!$H$344,\'[1]Weekly Inj_Rates\'!$H$359,\'[1]Weekly Inj_Rates\'!$H$370,\'[1]Weekly Inj_Rates\'!$H$381,\'[1]Weekly Inj_Rates\'!$H$396,\'[1]Weekly Inj_Rates\'!$H$411,\'[1]Weekly Inj_Rates\'!$H$422,\'[1]Weekly Inj_Rates\'!$H$433,\'[1]Weekly Inj_Rates\'!$H$457,\'[1]Weekly Inj_Rates\'!$H$474,\'[1]Weekly Inj_Rates\'!$H$490,\'[1]Weekly Inj_Rates\'!$H$502,\'[1]Weekly Inj_Rates\'!$H$512,\'[1]Weekly Inj_Rates\'!$H$524,\'[1]Weekly Inj_Rates\'!$H$540,\'[1]Weekly Inj_Rates\'!$H$556,\'[1]Weekly Inj_Rates\'!$H$572,\'[1]Weekly Inj_Rates\'!$H$584,\'[1]Weekly Inj_Rates\'!$H$596,\'[1]Weekly Inj_Rates\'!$H$608)";

            C.NumberingCache numberingCache3 = new C.NumberingCache();
            C.FormatCode formatCode3 = new C.FormatCode();
            formatCode3.Text = "d\\-mmm\\-yy";
            C.PointCount pointCount3 = new C.PointCount() { Val = (UInt32Value)40U };

            C.NumericPoint numericPoint81 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue83 = new C.NumericValue();
            numericValue83.Text = "42170";

            numericPoint81.Append(numericValue83);

            C.NumericPoint numericPoint82 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue84 = new C.NumericValue();
            numericValue84.Text = "42177";

            numericPoint82.Append(numericValue84);

            C.NumericPoint numericPoint83 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue85 = new C.NumericValue();
            numericValue85.Text = "42184";

            numericPoint83.Append(numericValue85);

            C.NumericPoint numericPoint84 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue86 = new C.NumericValue();
            numericValue86.Text = "42198";

            numericPoint84.Append(numericValue86);

            C.NumericPoint numericPoint85 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue87 = new C.NumericValue();
            numericValue87.Text = "42205";

            numericPoint85.Append(numericValue87);

            C.NumericPoint numericPoint86 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue88 = new C.NumericValue();
            numericValue88.Text = "42212";

            numericPoint86.Append(numericValue88);

            C.NumericPoint numericPoint87 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue89 = new C.NumericValue();
            numericValue89.Text = "42219";

            numericPoint87.Append(numericValue89);

            C.NumericPoint numericPoint88 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue90 = new C.NumericValue();
            numericValue90.Text = "42226";

            numericPoint88.Append(numericValue90);

            C.NumericPoint numericPoint89 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue91 = new C.NumericValue();
            numericValue91.Text = "42233";

            numericPoint89.Append(numericValue91);

            C.NumericPoint numericPoint90 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue92 = new C.NumericValue();
            numericValue92.Text = "42240.697916666664";

            numericPoint90.Append(numericValue92);

            C.NumericPoint numericPoint91 = new C.NumericPoint() { Index = (UInt32Value)10U };
            C.NumericValue numericValue93 = new C.NumericValue();
            numericValue93.Text = "42249.337500000001";

            numericPoint91.Append(numericValue93);

            C.NumericPoint numericPoint92 = new C.NumericPoint() { Index = (UInt32Value)11U };
            C.NumericValue numericValue94 = new C.NumericValue();
            numericValue94.Text = "42254.347222222219";

            numericPoint92.Append(numericValue94);

            C.NumericPoint numericPoint93 = new C.NumericPoint() { Index = (UInt32Value)12U };
            C.NumericValue numericValue95 = new C.NumericValue();
            numericValue95.Text = "42261.321527777778";

            numericPoint93.Append(numericValue95);

            C.NumericPoint numericPoint94 = new C.NumericPoint() { Index = (UInt32Value)13U };
            C.NumericValue numericValue96 = new C.NumericValue();
            numericValue96.Text = "42268.665972222225";

            numericPoint94.Append(numericValue96);

            C.NumericPoint numericPoint95 = new C.NumericPoint() { Index = (UInt32Value)14U };
            C.NumericValue numericValue97 = new C.NumericValue();
            numericValue97.Text = "42275.436111111114";

            numericPoint95.Append(numericValue97);

            C.NumericPoint numericPoint96 = new C.NumericPoint() { Index = (UInt32Value)15U };
            C.NumericValue numericValue98 = new C.NumericValue();
            numericValue98.Text = "42279.411805555559";

            numericPoint96.Append(numericValue98);

            C.NumericPoint numericPoint97 = new C.NumericPoint() { Index = (UInt32Value)16U };
            C.NumericValue numericValue99 = new C.NumericValue();
            numericValue99.Text = "42289.338888888888";

            numericPoint97.Append(numericValue99);

            C.NumericPoint numericPoint98 = new C.NumericPoint() { Index = (UInt32Value)17U };
            C.NumericValue numericValue100 = new C.NumericValue();
            numericValue100.Text = "42296.396527777775";

            numericPoint98.Append(numericValue100);

            C.NumericPoint numericPoint99 = new C.NumericPoint() { Index = (UInt32Value)18U };
            C.NumericValue numericValue101 = new C.NumericValue();
            numericValue101.Text = "42303.357638888891";

            numericPoint99.Append(numericValue101);

            C.NumericPoint numericPoint100 = new C.NumericPoint() { Index = (UInt32Value)19U };
            C.NumericValue numericValue102 = new C.NumericValue();
            numericValue102.Text = "42310.379861111112";

            numericPoint100.Append(numericValue102);

            C.NumericPoint numericPoint101 = new C.NumericPoint() { Index = (UInt32Value)20U };
            C.NumericValue numericValue103 = new C.NumericValue();
            numericValue103.Text = "42317.574999999997";

            numericPoint101.Append(numericValue103);

            C.NumericPoint numericPoint102 = new C.NumericPoint() { Index = (UInt32Value)21U };
            C.NumericValue numericValue104 = new C.NumericValue();
            numericValue104.Text = "42324.416666666664";

            numericPoint102.Append(numericValue104);

            C.NumericPoint numericPoint103 = new C.NumericPoint() { Index = (UInt32Value)22U };
            C.NumericValue numericValue105 = new C.NumericValue();
            numericValue105.Text = "42331.348611111112";

            numericPoint103.Append(numericValue105);

            C.NumericPoint numericPoint104 = new C.NumericPoint() { Index = (UInt32Value)23U };
            C.NumericValue numericValue106 = new C.NumericValue();
            numericValue106.Text = "42338.527777777781";

            numericPoint104.Append(numericValue106);

            C.NumericPoint numericPoint105 = new C.NumericPoint() { Index = (UInt32Value)24U };
            C.NumericValue numericValue107 = new C.NumericValue();
            numericValue107.Text = "42345.645833333336";

            numericPoint105.Append(numericValue107);

            C.NumericPoint numericPoint106 = new C.NumericPoint() { Index = (UInt32Value)25U };
            C.NumericValue numericValue108 = new C.NumericValue();
            numericValue108.Text = "42352.334027777775";

            numericPoint106.Append(numericValue108);

            C.NumericPoint numericPoint107 = new C.NumericPoint() { Index = (UInt32Value)26U };
            C.NumericValue numericValue109 = new C.NumericValue();
            numericValue109.Text = "42359.444444444445";

            numericPoint107.Append(numericValue109);

            C.NumericPoint numericPoint108 = new C.NumericPoint() { Index = (UInt32Value)27U };
            C.NumericValue numericValue110 = new C.NumericValue();
            numericValue110.Text = "42366.636805555558";

            numericPoint108.Append(numericValue110);

            C.NumericPoint numericPoint109 = new C.NumericPoint() { Index = (UInt32Value)28U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue111 = new C.NumericValue();
            numericValue111.Text = "42373.343055555553";

            numericPoint109.Append(numericValue111);

            C.NumericPoint numericPoint110 = new C.NumericPoint() { Index = (UInt32Value)29U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue112 = new C.NumericValue();
            numericValue112.Text = "42380.372916666667";

            numericPoint110.Append(numericValue112);

            C.NumericPoint numericPoint111 = new C.NumericPoint() { Index = (UInt32Value)30U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue113 = new C.NumericValue();
            numericValue113.Text = "42387.590277777781";

            numericPoint111.Append(numericValue113);

            C.NumericPoint numericPoint112 = new C.NumericPoint() { Index = (UInt32Value)31U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue114 = new C.NumericValue();
            numericValue114.Text = "42394.630555555559";

            numericPoint112.Append(numericValue114);

            C.NumericPoint numericPoint113 = new C.NumericPoint() { Index = (UInt32Value)32U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue115 = new C.NumericValue();
            numericValue115.Text = "42412.60833333333";

            numericPoint113.Append(numericValue115);

            C.NumericPoint numericPoint114 = new C.NumericPoint() { Index = (UInt32Value)33U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue116 = new C.NumericValue();
            numericValue116.Text = "42416.474999999999";

            numericPoint114.Append(numericValue116);

            C.NumericPoint numericPoint115 = new C.NumericPoint() { Index = (UInt32Value)34U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue117 = new C.NumericValue();
            numericValue117.Text = "42422.419444444444";

            numericPoint115.Append(numericValue117);

            C.NumericPoint numericPoint116 = new C.NumericPoint() { Index = (UInt32Value)35U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue118 = new C.NumericValue();
            numericValue118.Text = "42429.363888888889";

            numericPoint116.Append(numericValue118);

            C.NumericPoint numericPoint117 = new C.NumericPoint() { Index = (UInt32Value)36U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue119 = new C.NumericValue();
            numericValue119.Text = "42436.354166666664";

            numericPoint117.Append(numericValue119);

            C.NumericPoint numericPoint118 = new C.NumericPoint() { Index = (UInt32Value)37U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue120 = new C.NumericValue();
            numericValue120.Text = "42443.369444444441";

            numericPoint118.Append(numericValue120);

            C.NumericPoint numericPoint119 = new C.NumericPoint() { Index = (UInt32Value)38U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue121 = new C.NumericValue();
            numericValue121.Text = "42450.362500000003";

            numericPoint119.Append(numericValue121);

            C.NumericPoint numericPoint120 = new C.NumericPoint() { Index = (UInt32Value)39U, FormatCode = "m/d/yyyy\\ h:mm" };
            C.NumericValue numericValue122 = new C.NumericValue();
            numericValue122.Text = "42457.690972222219";

            numericPoint120.Append(numericValue122);

            numberingCache3.Append(formatCode3);
            numberingCache3.Append(pointCount3);
            numberingCache3.Append(numericPoint81);
            numberingCache3.Append(numericPoint82);
            numberingCache3.Append(numericPoint83);
            numberingCache3.Append(numericPoint84);
            numberingCache3.Append(numericPoint85);
            numberingCache3.Append(numericPoint86);
            numberingCache3.Append(numericPoint87);
            numberingCache3.Append(numericPoint88);
            numberingCache3.Append(numericPoint89);
            numberingCache3.Append(numericPoint90);
            numberingCache3.Append(numericPoint91);
            numberingCache3.Append(numericPoint92);
            numberingCache3.Append(numericPoint93);
            numberingCache3.Append(numericPoint94);
            numberingCache3.Append(numericPoint95);
            numberingCache3.Append(numericPoint96);
            numberingCache3.Append(numericPoint97);
            numberingCache3.Append(numericPoint98);
            numberingCache3.Append(numericPoint99);
            numberingCache3.Append(numericPoint100);
            numberingCache3.Append(numericPoint101);
            numberingCache3.Append(numericPoint102);
            numberingCache3.Append(numericPoint103);
            numberingCache3.Append(numericPoint104);
            numberingCache3.Append(numericPoint105);
            numberingCache3.Append(numericPoint106);
            numberingCache3.Append(numericPoint107);
            numberingCache3.Append(numericPoint108);
            numberingCache3.Append(numericPoint109);
            numberingCache3.Append(numericPoint110);
            numberingCache3.Append(numericPoint111);
            numberingCache3.Append(numericPoint112);
            numberingCache3.Append(numericPoint113);
            numberingCache3.Append(numericPoint114);
            numberingCache3.Append(numericPoint115);
            numberingCache3.Append(numericPoint116);
            numberingCache3.Append(numericPoint117);
            numberingCache3.Append(numericPoint118);
            numberingCache3.Append(numericPoint119);
            numberingCache3.Append(numericPoint120);

            numberReference3.Append(formula3);
            numberReference3.Append(numberingCache3);

            categoryAxisData2.Append(numberReference3);

            C.Values values2 = new C.Values();

            C.NumberReference numberReference4 = new C.NumberReference();
            C.Formula formula4 = new C.Formula();
            formula4.Text = "(\'[1]Weekly Ext_Rates\'!$G$24,\'[1]Weekly Ext_Rates\'!$G$43,\'[1]Weekly Ext_Rates\'!$G$62,\'[1]Weekly Ext_Rates\'!$G$81,\'[1]Weekly Ext_Rates\'!$G$96,\'[1]Weekly Ext_Rates\'!$G$115,\'[1]Weekly Ext_Rates\'!$G$157,\'[1]Weekly Ext_Rates\'!$G$182,\'[1]Weekly Ext_Rates\'!$G$206,\'[1]Weekly Ext_Rates\'!$G$233,\'[1]Weekly Ext_Rates\'!$G$252,\'[1]Weekly Ext_Rates\'!$G$276,\'[1]Weekly Ext_Rates\'!$G$303,\'[1]Weekly Ext_Rates\'!$G$322,\'[1]Weekly Ext_Rates\'!$G$341,\'[1]Weekly Ext_Rates\'!$G$409,\'[1]Weekly Ext_Rates\'!$G$436,\'[1]Weekly Ext_Rates\'!$G$463,\'[1]Weekly Ext_Rates\'!$G$490,\'[1]Weekly Ext_Rates\'!$G$517,\'[1]Weekly Ext_Rates\'!$G$544,\'[1]Weekly Ext_Rates\'!$G$571,\'[1]Weekly Ext_Rates\'!$G$590,\'[1]Weekly Ext_Rates\'!$G$601,\'[1]Weekly Ext_Rates\'!$G$612,\'[1]Weekly Ext_Rates\'!$G$623,\'[1]Weekly Ext_Rates\'!$G$634,\'[1]Weekly Ext_Rates\'!$G$644,\'[1]Weekly Ext_Rates\'!$G$687,\'[1]Weekly Ext_Rates\'!$G$698,\'[1]Weekly Ext_Rates\'!$G$710,\'[1]Weekly Ext_Rates\'!$G$722,\'[1]Weekly Ext_Rates\'!$G$734,\'[1]Weekly Ext_Rates\'!$G$746,\'[1]Weekly Ext_Rates\'!$G$758,\'[1]Weekly Ext_Rates\'!$G$785,\'[1]Weekly Ext_Rates\'!$G$813,\'[1]Weekly Ext_Rates\'!$G$833,\'[1]Weekly Ext_Rates\'!$G$873)";

            C.NumberingCache numberingCache4 = new C.NumberingCache();
            C.FormatCode formatCode4 = new C.FormatCode();
            formatCode4.Text = "0.00";
            C.PointCount pointCount4 = new C.PointCount() { Val = (UInt32Value)39U };

            C.NumericPoint numericPoint121 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue123 = new C.NumericValue();
            numericValue123.Text = "3.2672697549804566";

            numericPoint121.Append(numericValue123);

            C.NumericPoint numericPoint122 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue124 = new C.NumericValue();
            numericValue124.Text = "9.2263865620263061";

            numericPoint122.Append(numericValue124);

            C.NumericPoint numericPoint123 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue125 = new C.NumericValue();
            numericValue125.Text = "13.681127982651793";

            numericPoint123.Append(numericValue125);

            C.NumericPoint numericPoint124 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue126 = new C.NumericValue();
            numericValue126.Text = "2.6592960182840635";

            numericPoint124.Append(numericValue126);

            C.NumericPoint numericPoint125 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue127 = new C.NumericValue();
            numericValue127.Text = "7.4426441757008881";

            numericPoint125.Append(numericValue127);

            C.NumericPoint numericPoint126 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue128 = new C.NumericValue();
            numericValue128.Text = "5.6015123244620817";

            numericPoint126.Append(numericValue128);

            C.NumericPoint numericPoint127 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue129 = new C.NumericValue();
            numericValue129.Text = "5.0887618369093079";

            numericPoint127.Append(numericValue129);

            C.NumericPoint numericPoint128 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue130 = new C.NumericValue();
            numericValue130.Text = "18.742774255575259";

            numericPoint128.Append(numericValue130);

            C.NumericPoint numericPoint129 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue131 = new C.NumericValue();
            numericValue131.Text = "12.003589416226429";

            numericPoint129.Append(numericValue131);

            C.NumericPoint numericPoint130 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue132 = new C.NumericValue();
            numericValue132.Text = "7.9048913157028133";

            numericPoint130.Append(numericValue132);

            C.NumericPoint numericPoint131 = new C.NumericPoint() { Index = (UInt32Value)10U };
            C.NumericValue numericValue133 = new C.NumericValue();
            numericValue133.Text = "14.665558129991911";

            numericPoint131.Append(numericValue133);

            C.NumericPoint numericPoint132 = new C.NumericPoint() { Index = (UInt32Value)11U };
            C.NumericValue numericValue134 = new C.NumericValue();
            numericValue134.Text = "12.380763366868639";

            numericPoint132.Append(numericValue134);

            C.NumericPoint numericPoint133 = new C.NumericPoint() { Index = (UInt32Value)12U };
            C.NumericValue numericValue135 = new C.NumericValue();
            numericValue135.Text = "10.387718321125073";

            numericPoint133.Append(numericValue135);

            C.NumericPoint numericPoint134 = new C.NumericPoint() { Index = (UInt32Value)13U };
            C.NumericValue numericValue136 = new C.NumericValue();
            numericValue136.Text = "9.2990869725053749";

            numericPoint134.Append(numericValue136);

            C.NumericPoint numericPoint135 = new C.NumericPoint() { Index = (UInt32Value)14U };
            C.NumericValue numericValue137 = new C.NumericValue();
            numericValue137.Text = "18.221619436565931";

            numericPoint135.Append(numericValue137);

            C.NumericPoint numericPoint136 = new C.NumericPoint() { Index = (UInt32Value)15U };
            C.NumericValue numericValue138 = new C.NumericValue();
            numericValue138.Text = "15.848566528659374";

            numericPoint136.Append(numericValue138);

            C.NumericPoint numericPoint137 = new C.NumericPoint() { Index = (UInt32Value)16U };
            C.NumericValue numericValue139 = new C.NumericValue();
            numericValue139.Text = "14.660465558590964";

            numericPoint137.Append(numericValue139);

            C.NumericPoint numericPoint138 = new C.NumericPoint() { Index = (UInt32Value)17U };
            C.NumericValue numericValue140 = new C.NumericValue();
            numericValue140.Text = "8.7491305966567907";

            numericPoint138.Append(numericValue140);

            C.NumericPoint numericPoint139 = new C.NumericPoint() { Index = (UInt32Value)18U };
            C.NumericValue numericValue141 = new C.NumericValue();
            numericValue141.Text = "15.33843097020573";

            numericPoint139.Append(numericValue141);

            C.NumericPoint numericPoint140 = new C.NumericPoint() { Index = (UInt32Value)19U };
            C.NumericValue numericValue142 = new C.NumericValue();
            numericValue142.Text = "14.950672359520716";

            numericPoint140.Append(numericValue142);

            C.NumericPoint numericPoint141 = new C.NumericPoint() { Index = (UInt32Value)20U };
            C.NumericValue numericValue143 = new C.NumericValue();
            numericValue143.Text = "15.335116113330439";

            numericPoint141.Append(numericValue143);

            C.NumericPoint numericPoint142 = new C.NumericPoint() { Index = (UInt32Value)21U };
            C.NumericValue numericValue144 = new C.NumericValue();
            numericValue144.Text = "13.394851398807432";

            numericPoint142.Append(numericValue144);

            C.NumericPoint numericPoint143 = new C.NumericPoint() { Index = (UInt32Value)22U };
            C.NumericValue numericValue145 = new C.NumericValue();
            numericValue145.Text = "15.593622735486486";

            numericPoint143.Append(numericValue145);

            C.NumericPoint numericPoint144 = new C.NumericPoint() { Index = (UInt32Value)23U };
            C.NumericValue numericValue146 = new C.NumericValue();
            numericValue146.Text = "15.085073708190903";

            numericPoint144.Append(numericValue146);

            C.NumericPoint numericPoint145 = new C.NumericPoint() { Index = (UInt32Value)24U };
            C.NumericValue numericValue147 = new C.NumericValue();
            numericValue147.Text = "20.304568932022079";

            numericPoint145.Append(numericValue147);

            C.NumericPoint numericPoint146 = new C.NumericPoint() { Index = (UInt32Value)25U };
            C.NumericValue numericValue148 = new C.NumericValue();
            numericValue148.Text = "19.761578453200041";

            numericPoint146.Append(numericValue148);

            C.NumericPoint numericPoint147 = new C.NumericPoint() { Index = (UInt32Value)26U };
            C.NumericValue numericValue149 = new C.NumericValue();
            numericValue149.Text = "18.12904033767834";

            numericPoint147.Append(numericValue149);

            C.NumericPoint numericPoint148 = new C.NumericPoint() { Index = (UInt32Value)27U };
            C.NumericValue numericValue150 = new C.NumericValue();
            numericValue150.Text = "15.28590107069822";

            numericPoint148.Append(numericValue150);

            C.NumericPoint numericPoint149 = new C.NumericPoint() { Index = (UInt32Value)28U };
            C.NumericValue numericValue151 = new C.NumericValue();
            numericValue151.Text = "16.131771273293516";

            numericPoint149.Append(numericValue151);

            C.NumericPoint numericPoint150 = new C.NumericPoint() { Index = (UInt32Value)29U };
            C.NumericValue numericValue152 = new C.NumericValue();
            numericValue152.Text = "17.981858880299484";

            numericPoint150.Append(numericValue152);

            C.NumericPoint numericPoint151 = new C.NumericPoint() { Index = (UInt32Value)30U };
            C.NumericValue numericValue153 = new C.NumericValue();
            numericValue153.Text = "16.878771198185113";

            numericPoint151.Append(numericValue153);

            C.NumericPoint numericPoint152 = new C.NumericPoint() { Index = (UInt32Value)31U };
            C.NumericValue numericValue154 = new C.NumericValue();
            numericValue154.Text = "9.0415236554635285";

            numericPoint152.Append(numericValue154);

            C.NumericPoint numericPoint153 = new C.NumericPoint() { Index = (UInt32Value)32U };
            C.NumericValue numericValue155 = new C.NumericValue();
            numericValue155.Text = "3.4121827924497659";

            numericPoint153.Append(numericValue155);

            C.NumericPoint numericPoint154 = new C.NumericPoint() { Index = (UInt32Value)33U };
            C.NumericValue numericValue156 = new C.NumericValue();
            numericValue156.Text = "0.24764791455825383";

            numericPoint154.Append(numericValue156);

            C.NumericPoint numericPoint155 = new C.NumericPoint() { Index = (UInt32Value)34U };
            C.NumericValue numericValue157 = new C.NumericValue();
            numericValue157.Text = "14.447321207159874";

            numericPoint155.Append(numericValue157);

            C.NumericPoint numericPoint156 = new C.NumericPoint() { Index = (UInt32Value)35U };
            C.NumericValue numericValue158 = new C.NumericValue();
            numericValue158.Text = "5.101837591791492";

            numericPoint156.Append(numericValue158);

            C.NumericPoint numericPoint157 = new C.NumericPoint() { Index = (UInt32Value)36U };
            C.NumericValue numericValue159 = new C.NumericValue();
            numericValue159.Text = "6.0475752393821907";

            numericPoint157.Append(numericValue159);

            C.NumericPoint numericPoint158 = new C.NumericPoint() { Index = (UInt32Value)37U };
            C.NumericValue numericValue160 = new C.NumericValue();
            numericValue160.Text = "6.2183426235731378";

            numericPoint158.Append(numericValue160);

            C.NumericPoint numericPoint159 = new C.NumericPoint() { Index = (UInt32Value)38U };
            C.NumericValue numericValue161 = new C.NumericValue();
            numericValue161.Text = "6.183536976732821";

            numericPoint159.Append(numericValue161);

            numberingCache4.Append(formatCode4);
            numberingCache4.Append(pointCount4);
            numberingCache4.Append(numericPoint121);
            numberingCache4.Append(numericPoint122);
            numberingCache4.Append(numericPoint123);
            numberingCache4.Append(numericPoint124);
            numberingCache4.Append(numericPoint125);
            numberingCache4.Append(numericPoint126);
            numberingCache4.Append(numericPoint127);
            numberingCache4.Append(numericPoint128);
            numberingCache4.Append(numericPoint129);
            numberingCache4.Append(numericPoint130);
            numberingCache4.Append(numericPoint131);
            numberingCache4.Append(numericPoint132);
            numberingCache4.Append(numericPoint133);
            numberingCache4.Append(numericPoint134);
            numberingCache4.Append(numericPoint135);
            numberingCache4.Append(numericPoint136);
            numberingCache4.Append(numericPoint137);
            numberingCache4.Append(numericPoint138);
            numberingCache4.Append(numericPoint139);
            numberingCache4.Append(numericPoint140);
            numberingCache4.Append(numericPoint141);
            numberingCache4.Append(numericPoint142);
            numberingCache4.Append(numericPoint143);
            numberingCache4.Append(numericPoint144);
            numberingCache4.Append(numericPoint145);
            numberingCache4.Append(numericPoint146);
            numberingCache4.Append(numericPoint147);
            numberingCache4.Append(numericPoint148);
            numberingCache4.Append(numericPoint149);
            numberingCache4.Append(numericPoint150);
            numberingCache4.Append(numericPoint151);
            numberingCache4.Append(numericPoint152);
            numberingCache4.Append(numericPoint153);
            numberingCache4.Append(numericPoint154);
            numberingCache4.Append(numericPoint155);
            numberingCache4.Append(numericPoint156);
            numberingCache4.Append(numericPoint157);
            numberingCache4.Append(numericPoint158);
            numberingCache4.Append(numericPoint159);

            numberReference4.Append(formula4);
            numberReference4.Append(numberingCache4);

            values2.Append(numberReference4);

            C.BarSerExtensionList barSerExtensionList2 = new C.BarSerExtensionList();
            barSerExtensionList2.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");

            C.BarSerExtension barSerExtension2 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            barSerExtension2.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000001-5AC9-435B-A402-5E3C89F8E9A8}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            barSerExtension2.Append(openXmlUnknownElement2);

            barSerExtensionList2.Append(barSerExtension2);

            barChartSeries2.Append(index2);
            barChartSeries2.Append(order2);
            barChartSeries2.Append(seriesText2);
            barChartSeries2.Append(chartShapeProperties2);
            barChartSeries2.Append(invertIfNegative2);
            barChartSeries2.Append(categoryAxisData2);
            barChartSeries2.Append(values2);
            barChartSeries2.Append(barSerExtensionList2);

            C.DataLabels dataLabels1 = new C.DataLabels();
            C.ShowLegendKey showLegendKey1 = new C.ShowLegendKey() { Val = false };
            C.ShowValue showValue1 = new C.ShowValue() { Val = false };
            C.ShowCategoryName showCategoryName1 = new C.ShowCategoryName() { Val = false };
            C.ShowSeriesName showSeriesName1 = new C.ShowSeriesName() { Val = false };
            C.ShowPercent showPercent1 = new C.ShowPercent() { Val = false };
            C.ShowBubbleSize showBubbleSize1 = new C.ShowBubbleSize() { Val = false };

            dataLabels1.Append(showLegendKey1);
            dataLabels1.Append(showValue1);
            dataLabels1.Append(showCategoryName1);
            dataLabels1.Append(showSeriesName1);
            dataLabels1.Append(showPercent1);
            dataLabels1.Append(showBubbleSize1);
            C.GapWidth gapWidth1 = new C.GapWidth() { Val = (UInt16Value)219U };
            C.Overlap overlap1 = new C.Overlap() { Val = -27 };
            C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)458742528U };
            C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)458741744U };

            barChart1.Append(barDirection1);
            barChart1.Append(barGrouping1);
            barChart1.Append(varyColors1);
            barChart1.Append(barChartSeries1);
            barChart1.Append(barChartSeries2);
            barChart1.Append(dataLabels1);
            barChart1.Append(gapWidth1);
            barChart1.Append(overlap1);
            barChart1.Append(axisId1);
            barChart1.Append(axisId2);

            C.LineChart lineChart1 = new C.LineChart();
            C.Grouping grouping1 = new C.Grouping() { Val = C.GroupingValues.Standard };
            C.VaryColors varyColors2 = new C.VaryColors() { Val = false };

            C.LineChartSeries lineChartSeries1 = new C.LineChartSeries();
            C.Index index3 = new C.Index() { Val = (UInt32Value)2U };
            C.Order order3 = new C.Order() { Val = (UInt32Value)2U };

            C.SeriesText seriesText3 = new C.SeriesText();
            C.NumericValue numericValue162 = new C.NumericValue();
            numericValue162.Text = "Milestone Weekly Flowrate";

            seriesText3.Append(numericValue162);

            C.ChartShapeProperties chartShapeProperties3 = new C.ChartShapeProperties();

            A.Outline outline3 = new A.Outline() { Width = 22225, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FF0000" };

            solidFill3.Append(rgbColorModelHex1);
            A.Round round1 = new A.Round();

            outline3.Append(solidFill3);
            outline3.Append(round1);
            A.EffectList effectList3 = new A.EffectList();

            chartShapeProperties3.Append(outline3);
            chartShapeProperties3.Append(effectList3);

            C.Marker marker1 = new C.Marker();
            C.Symbol symbol1 = new C.Symbol() { Val = C.MarkerStyleValues.None };

            marker1.Append(symbol1);

            C.CategoryAxisData categoryAxisData3 = new C.CategoryAxisData();

            C.NumberReference numberReference5 = new C.NumberReference();
            C.Formula formula5 = new C.Formula();
            formula5.Text = "([1]Summary!$B$12,[1]Summary!$B$13,[1]Summary!$B$14,[1]Summary!$B$15,[1]Summary!$B$16,[1]Summary!$B$17,[1]Summary!$B$18,[1]Summary!$B$19,[1]Summary!$B$20,[1]Summary!$B$21,[1]Summary!$B$22,[1]Summary!$B$23,[1]Summary!$B$24,[1]Summary!$B$25,[1]Summary!$B$26,[1]Summary!$B$27,[1]Summary!$B$28,[1]Summary!$B$29,[1]Summary!$B$30,[1]Summary!$B$31,[1]Summary!$B$32,[1]Summary!$B$33,[1]Summary!$B$34,[1]Summary!$B$35,[1]Summary!$B$36,[1]Summary!$B$37,[1]Summary!$B$38,[1]Summary!$B$39,[1]Summary!$B$40,[1]Summary!$B$41,[1]Summary!$B$42,[1]Summary!$B$43,[1]Summary!$B$44,[1]Summary!$B$45,[1]Summary!$B$46,[1]Summary!$B$47,[1]Summary!$B$48,[1]Summary!$B$49,[1]Summary!$B$50,[1]Summary!$B$51)";

            C.NumberingCache numberingCache5 = new C.NumberingCache();
            C.FormatCode formatCode5 = new C.FormatCode();
            formatCode5.Text = "d\\-mmm\\-yy";
            C.PointCount pointCount5 = new C.PointCount() { Val = (UInt32Value)40U };

            C.NumericPoint numericPoint160 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue163 = new C.NumericValue();
            numericValue163.Text = "42170";

            numericPoint160.Append(numericValue163);

            C.NumericPoint numericPoint161 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue164 = new C.NumericValue();
            numericValue164.Text = "42177";

            numericPoint161.Append(numericValue164);

            C.NumericPoint numericPoint162 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue165 = new C.NumericValue();
            numericValue165.Text = "42184";

            numericPoint162.Append(numericValue165);

            C.NumericPoint numericPoint163 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue166 = new C.NumericValue();
            numericValue166.Text = "42198";

            numericPoint163.Append(numericValue166);

            C.NumericPoint numericPoint164 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue167 = new C.NumericValue();
            numericValue167.Text = "42205";

            numericPoint164.Append(numericValue167);

            C.NumericPoint numericPoint165 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue168 = new C.NumericValue();
            numericValue168.Text = "42212";

            numericPoint165.Append(numericValue168);

            C.NumericPoint numericPoint166 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue169 = new C.NumericValue();
            numericValue169.Text = "42219";

            numericPoint166.Append(numericValue169);

            C.NumericPoint numericPoint167 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue170 = new C.NumericValue();
            numericValue170.Text = "42226";

            numericPoint167.Append(numericValue170);

            C.NumericPoint numericPoint168 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue171 = new C.NumericValue();
            numericValue171.Text = "42233";

            numericPoint168.Append(numericValue171);

            C.NumericPoint numericPoint169 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue172 = new C.NumericValue();
            numericValue172.Text = "42240";

            numericPoint169.Append(numericValue172);

            C.NumericPoint numericPoint170 = new C.NumericPoint() { Index = (UInt32Value)10U };
            C.NumericValue numericValue173 = new C.NumericValue();
            numericValue173.Text = "42249";

            numericPoint170.Append(numericValue173);

            C.NumericPoint numericPoint171 = new C.NumericPoint() { Index = (UInt32Value)11U };
            C.NumericValue numericValue174 = new C.NumericValue();
            numericValue174.Text = "42254";

            numericPoint171.Append(numericValue174);

            C.NumericPoint numericPoint172 = new C.NumericPoint() { Index = (UInt32Value)12U };
            C.NumericValue numericValue175 = new C.NumericValue();
            numericValue175.Text = "42261";

            numericPoint172.Append(numericValue175);

            C.NumericPoint numericPoint173 = new C.NumericPoint() { Index = (UInt32Value)13U };
            C.NumericValue numericValue176 = new C.NumericValue();
            numericValue176.Text = "42268";

            numericPoint173.Append(numericValue176);

            C.NumericPoint numericPoint174 = new C.NumericPoint() { Index = (UInt32Value)14U };
            C.NumericValue numericValue177 = new C.NumericValue();
            numericValue177.Text = "42275";

            numericPoint174.Append(numericValue177);

            C.NumericPoint numericPoint175 = new C.NumericPoint() { Index = (UInt32Value)15U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue178 = new C.NumericValue();
            numericValue178.Text = "42279.642361111109";

            numericPoint175.Append(numericValue178);

            C.NumericPoint numericPoint176 = new C.NumericPoint() { Index = (UInt32Value)16U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue179 = new C.NumericValue();
            numericValue179.Text = "42289.361805555556";

            numericPoint176.Append(numericValue179);

            C.NumericPoint numericPoint177 = new C.NumericPoint() { Index = (UInt32Value)17U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue180 = new C.NumericValue();
            numericValue180.Text = "42296.554166666669";

            numericPoint177.Append(numericValue180);

            C.NumericPoint numericPoint178 = new C.NumericPoint() { Index = (UInt32Value)18U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue181 = new C.NumericValue();
            numericValue181.Text = "42303.370138888888";

            numericPoint178.Append(numericValue181);

            C.NumericPoint numericPoint179 = new C.NumericPoint() { Index = (UInt32Value)19U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue182 = new C.NumericValue();
            numericValue182.Text = "42310.397222222222";

            numericPoint179.Append(numericValue182);

            C.NumericPoint numericPoint180 = new C.NumericPoint() { Index = (UInt32Value)20U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue183 = new C.NumericValue();
            numericValue183.Text = "42317.643055555556";

            numericPoint180.Append(numericValue183);

            C.NumericPoint numericPoint181 = new C.NumericPoint() { Index = (UInt32Value)21U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue184 = new C.NumericValue();
            numericValue184.Text = "42324.479166666664";

            numericPoint181.Append(numericValue184);

            C.NumericPoint numericPoint182 = new C.NumericPoint() { Index = (UInt32Value)22U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue185 = new C.NumericValue();
            numericValue185.Text = "42331.404861111114";

            numericPoint182.Append(numericValue185);

            C.NumericPoint numericPoint183 = new C.NumericPoint() { Index = (UInt32Value)23U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue186 = new C.NumericValue();
            numericValue186.Text = "42338.544444444444";

            numericPoint183.Append(numericValue186);

            C.NumericPoint numericPoint184 = new C.NumericPoint() { Index = (UInt32Value)24U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue187 = new C.NumericValue();
            numericValue187.Text = "42345.614583333336";

            numericPoint184.Append(numericValue187);

            C.NumericPoint numericPoint185 = new C.NumericPoint() { Index = (UInt32Value)25U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue188 = new C.NumericValue();
            numericValue188.Text = "42352.366666666669";

            numericPoint185.Append(numericValue188);

            C.NumericPoint numericPoint186 = new C.NumericPoint() { Index = (UInt32Value)26U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue189 = new C.NumericValue();
            numericValue189.Text = "42359.385416666664";

            numericPoint186.Append(numericValue189);

            C.NumericPoint numericPoint187 = new C.NumericPoint() { Index = (UInt32Value)27U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue190 = new C.NumericValue();
            numericValue190.Text = "42366.541666666664";

            numericPoint187.Append(numericValue190);

            C.NumericPoint numericPoint188 = new C.NumericPoint() { Index = (UInt32Value)28U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue191 = new C.NumericValue();
            numericValue191.Text = "42373.343055555553";

            numericPoint188.Append(numericValue191);

            C.NumericPoint numericPoint189 = new C.NumericPoint() { Index = (UInt32Value)29U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue192 = new C.NumericValue();
            numericValue192.Text = "42380.372916666667";

            numericPoint189.Append(numericValue192);

            C.NumericPoint numericPoint190 = new C.NumericPoint() { Index = (UInt32Value)30U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue193 = new C.NumericValue();
            numericValue193.Text = "42387.590277777781";

            numericPoint190.Append(numericValue193);

            C.NumericPoint numericPoint191 = new C.NumericPoint() { Index = (UInt32Value)31U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue194 = new C.NumericValue();
            numericValue194.Text = "42394.630555555559";

            numericPoint191.Append(numericValue194);

            C.NumericPoint numericPoint192 = new C.NumericPoint() { Index = (UInt32Value)32U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue195 = new C.NumericValue();
            numericValue195.Text = "42412.60833333333";

            numericPoint192.Append(numericValue195);

            C.NumericPoint numericPoint193 = new C.NumericPoint() { Index = (UInt32Value)33U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue196 = new C.NumericValue();
            numericValue196.Text = "42416.474999999999";

            numericPoint193.Append(numericValue196);

            C.NumericPoint numericPoint194 = new C.NumericPoint() { Index = (UInt32Value)34U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue197 = new C.NumericValue();
            numericValue197.Text = "42422.419444444444";

            numericPoint194.Append(numericValue197);

            C.NumericPoint numericPoint195 = new C.NumericPoint() { Index = (UInt32Value)35U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue198 = new C.NumericValue();
            numericValue198.Text = "42429.363888888889";

            numericPoint195.Append(numericValue198);

            C.NumericPoint numericPoint196 = new C.NumericPoint() { Index = (UInt32Value)36U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue199 = new C.NumericValue();
            numericValue199.Text = "42436.354166666664";

            numericPoint196.Append(numericValue199);

            C.NumericPoint numericPoint197 = new C.NumericPoint() { Index = (UInt32Value)37U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue200 = new C.NumericValue();
            numericValue200.Text = "42443.369444444441";

            numericPoint197.Append(numericValue200);

            C.NumericPoint numericPoint198 = new C.NumericPoint() { Index = (UInt32Value)38U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue201 = new C.NumericValue();
            numericValue201.Text = "42450.362500000003";

            numericPoint198.Append(numericValue201);

            C.NumericPoint numericPoint199 = new C.NumericPoint() { Index = (UInt32Value)39U, FormatCode = "[$-409]d\\-mmm\\-yy;@" };
            C.NumericValue numericValue202 = new C.NumericValue();
            numericValue202.Text = "42457.690972222219";

            numericPoint199.Append(numericValue202);

            numberingCache5.Append(formatCode5);
            numberingCache5.Append(pointCount5);
            numberingCache5.Append(numericPoint160);
            numberingCache5.Append(numericPoint161);
            numberingCache5.Append(numericPoint162);
            numberingCache5.Append(numericPoint163);
            numberingCache5.Append(numericPoint164);
            numberingCache5.Append(numericPoint165);
            numberingCache5.Append(numericPoint166);
            numberingCache5.Append(numericPoint167);
            numberingCache5.Append(numericPoint168);
            numberingCache5.Append(numericPoint169);
            numberingCache5.Append(numericPoint170);
            numberingCache5.Append(numericPoint171);
            numberingCache5.Append(numericPoint172);
            numberingCache5.Append(numericPoint173);
            numberingCache5.Append(numericPoint174);
            numberingCache5.Append(numericPoint175);
            numberingCache5.Append(numericPoint176);
            numberingCache5.Append(numericPoint177);
            numberingCache5.Append(numericPoint178);
            numberingCache5.Append(numericPoint179);
            numberingCache5.Append(numericPoint180);
            numberingCache5.Append(numericPoint181);
            numberingCache5.Append(numericPoint182);
            numberingCache5.Append(numericPoint183);
            numberingCache5.Append(numericPoint184);
            numberingCache5.Append(numericPoint185);
            numberingCache5.Append(numericPoint186);
            numberingCache5.Append(numericPoint187);
            numberingCache5.Append(numericPoint188);
            numberingCache5.Append(numericPoint189);
            numberingCache5.Append(numericPoint190);
            numberingCache5.Append(numericPoint191);
            numberingCache5.Append(numericPoint192);
            numberingCache5.Append(numericPoint193);
            numberingCache5.Append(numericPoint194);
            numberingCache5.Append(numericPoint195);
            numberingCache5.Append(numericPoint196);
            numberingCache5.Append(numericPoint197);
            numberingCache5.Append(numericPoint198);
            numberingCache5.Append(numericPoint199);

            numberReference5.Append(formula5);
            numberReference5.Append(numberingCache5);

            categoryAxisData3.Append(numberReference5);

            C.Values values3 = new C.Values();

            C.NumberReference numberReference6 = new C.NumberReference();
            C.Formula formula6 = new C.Formula();
            formula6.Text = "[1]Summary!$C$12:$C$51";

            C.NumberingCache numberingCache6 = new C.NumberingCache();
            C.FormatCode formatCode6 = new C.FormatCode();
            formatCode6.Text = "General";
            C.PointCount pointCount6 = new C.PointCount() { Val = (UInt32Value)40U };

            C.NumericPoint numericPoint200 = new C.NumericPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue203 = new C.NumericValue();
            numericValue203.Text = "28.3";

            numericPoint200.Append(numericValue203);

            C.NumericPoint numericPoint201 = new C.NumericPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue204 = new C.NumericValue();
            numericValue204.Text = "28.3";

            numericPoint201.Append(numericValue204);

            C.NumericPoint numericPoint202 = new C.NumericPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue205 = new C.NumericValue();
            numericValue205.Text = "28.3";

            numericPoint202.Append(numericValue205);

            C.NumericPoint numericPoint203 = new C.NumericPoint() { Index = (UInt32Value)3U };
            C.NumericValue numericValue206 = new C.NumericValue();
            numericValue206.Text = "28.3";

            numericPoint203.Append(numericValue206);

            C.NumericPoint numericPoint204 = new C.NumericPoint() { Index = (UInt32Value)4U };
            C.NumericValue numericValue207 = new C.NumericValue();
            numericValue207.Text = "28.3";

            numericPoint204.Append(numericValue207);

            C.NumericPoint numericPoint205 = new C.NumericPoint() { Index = (UInt32Value)5U };
            C.NumericValue numericValue208 = new C.NumericValue();
            numericValue208.Text = "28.3";

            numericPoint205.Append(numericValue208);

            C.NumericPoint numericPoint206 = new C.NumericPoint() { Index = (UInt32Value)6U };
            C.NumericValue numericValue209 = new C.NumericValue();
            numericValue209.Text = "28.3";

            numericPoint206.Append(numericValue209);

            C.NumericPoint numericPoint207 = new C.NumericPoint() { Index = (UInt32Value)7U };
            C.NumericValue numericValue210 = new C.NumericValue();
            numericValue210.Text = "28.3";

            numericPoint207.Append(numericValue210);

            C.NumericPoint numericPoint208 = new C.NumericPoint() { Index = (UInt32Value)8U };
            C.NumericValue numericValue211 = new C.NumericValue();
            numericValue211.Text = "28.3";

            numericPoint208.Append(numericValue211);

            C.NumericPoint numericPoint209 = new C.NumericPoint() { Index = (UInt32Value)9U };
            C.NumericValue numericValue212 = new C.NumericValue();
            numericValue212.Text = "28.3";

            numericPoint209.Append(numericValue212);

            C.NumericPoint numericPoint210 = new C.NumericPoint() { Index = (UInt32Value)10U };
            C.NumericValue numericValue213 = new C.NumericValue();
            numericValue213.Text = "28.3";

            numericPoint210.Append(numericValue213);

            C.NumericPoint numericPoint211 = new C.NumericPoint() { Index = (UInt32Value)11U };
            C.NumericValue numericValue214 = new C.NumericValue();
            numericValue214.Text = "28.3";

            numericPoint211.Append(numericValue214);

            C.NumericPoint numericPoint212 = new C.NumericPoint() { Index = (UInt32Value)12U };
            C.NumericValue numericValue215 = new C.NumericValue();
            numericValue215.Text = "28.3";

            numericPoint212.Append(numericValue215);

            C.NumericPoint numericPoint213 = new C.NumericPoint() { Index = (UInt32Value)13U };
            C.NumericValue numericValue216 = new C.NumericValue();
            numericValue216.Text = "28.3";

            numericPoint213.Append(numericValue216);

            C.NumericPoint numericPoint214 = new C.NumericPoint() { Index = (UInt32Value)14U };
            C.NumericValue numericValue217 = new C.NumericValue();
            numericValue217.Text = "28.3";

            numericPoint214.Append(numericValue217);

            C.NumericPoint numericPoint215 = new C.NumericPoint() { Index = (UInt32Value)15U };
            C.NumericValue numericValue218 = new C.NumericValue();
            numericValue218.Text = "28.3";

            numericPoint215.Append(numericValue218);

            C.NumericPoint numericPoint216 = new C.NumericPoint() { Index = (UInt32Value)16U };
            C.NumericValue numericValue219 = new C.NumericValue();
            numericValue219.Text = "28.3";

            numericPoint216.Append(numericValue219);

            C.NumericPoint numericPoint217 = new C.NumericPoint() { Index = (UInt32Value)17U };
            C.NumericValue numericValue220 = new C.NumericValue();
            numericValue220.Text = "28.3";

            numericPoint217.Append(numericValue220);

            C.NumericPoint numericPoint218 = new C.NumericPoint() { Index = (UInt32Value)18U };
            C.NumericValue numericValue221 = new C.NumericValue();
            numericValue221.Text = "28.3";

            numericPoint218.Append(numericValue221);

            C.NumericPoint numericPoint219 = new C.NumericPoint() { Index = (UInt32Value)19U };
            C.NumericValue numericValue222 = new C.NumericValue();
            numericValue222.Text = "28.3";

            numericPoint219.Append(numericValue222);

            C.NumericPoint numericPoint220 = new C.NumericPoint() { Index = (UInt32Value)20U };
            C.NumericValue numericValue223 = new C.NumericValue();
            numericValue223.Text = "28.3";

            numericPoint220.Append(numericValue223);

            C.NumericPoint numericPoint221 = new C.NumericPoint() { Index = (UInt32Value)21U };
            C.NumericValue numericValue224 = new C.NumericValue();
            numericValue224.Text = "28.3";

            numericPoint221.Append(numericValue224);

            C.NumericPoint numericPoint222 = new C.NumericPoint() { Index = (UInt32Value)22U };
            C.NumericValue numericValue225 = new C.NumericValue();
            numericValue225.Text = "28.3";

            numericPoint222.Append(numericValue225);

            C.NumericPoint numericPoint223 = new C.NumericPoint() { Index = (UInt32Value)23U };
            C.NumericValue numericValue226 = new C.NumericValue();
            numericValue226.Text = "28.3";

            numericPoint223.Append(numericValue226);

            C.NumericPoint numericPoint224 = new C.NumericPoint() { Index = (UInt32Value)24U };
            C.NumericValue numericValue227 = new C.NumericValue();
            numericValue227.Text = "28.3";

            numericPoint224.Append(numericValue227);

            C.NumericPoint numericPoint225 = new C.NumericPoint() { Index = (UInt32Value)25U };
            C.NumericValue numericValue228 = new C.NumericValue();
            numericValue228.Text = "28.3";

            numericPoint225.Append(numericValue228);

            C.NumericPoint numericPoint226 = new C.NumericPoint() { Index = (UInt32Value)26U };
            C.NumericValue numericValue229 = new C.NumericValue();
            numericValue229.Text = "28.3";

            numericPoint226.Append(numericValue229);

            C.NumericPoint numericPoint227 = new C.NumericPoint() { Index = (UInt32Value)27U };
            C.NumericValue numericValue230 = new C.NumericValue();
            numericValue230.Text = "28.3";

            numericPoint227.Append(numericValue230);

            C.NumericPoint numericPoint228 = new C.NumericPoint() { Index = (UInt32Value)28U };
            C.NumericValue numericValue231 = new C.NumericValue();
            numericValue231.Text = "28.3";

            numericPoint228.Append(numericValue231);

            C.NumericPoint numericPoint229 = new C.NumericPoint() { Index = (UInt32Value)29U };
            C.NumericValue numericValue232 = new C.NumericValue();
            numericValue232.Text = "28.3";

            numericPoint229.Append(numericValue232);

            C.NumericPoint numericPoint230 = new C.NumericPoint() { Index = (UInt32Value)30U };
            C.NumericValue numericValue233 = new C.NumericValue();
            numericValue233.Text = "28.3";

            numericPoint230.Append(numericValue233);

            C.NumericPoint numericPoint231 = new C.NumericPoint() { Index = (UInt32Value)31U };
            C.NumericValue numericValue234 = new C.NumericValue();
            numericValue234.Text = "28.3";

            numericPoint231.Append(numericValue234);

            C.NumericPoint numericPoint232 = new C.NumericPoint() { Index = (UInt32Value)32U };
            C.NumericValue numericValue235 = new C.NumericValue();
            numericValue235.Text = "28.3";

            numericPoint232.Append(numericValue235);

            C.NumericPoint numericPoint233 = new C.NumericPoint() { Index = (UInt32Value)33U };
            C.NumericValue numericValue236 = new C.NumericValue();
            numericValue236.Text = "28.3";

            numericPoint233.Append(numericValue236);

            C.NumericPoint numericPoint234 = new C.NumericPoint() { Index = (UInt32Value)34U };
            C.NumericValue numericValue237 = new C.NumericValue();
            numericValue237.Text = "28.3";

            numericPoint234.Append(numericValue237);

            C.NumericPoint numericPoint235 = new C.NumericPoint() { Index = (UInt32Value)35U };
            C.NumericValue numericValue238 = new C.NumericValue();
            numericValue238.Text = "28.3";

            numericPoint235.Append(numericValue238);

            C.NumericPoint numericPoint236 = new C.NumericPoint() { Index = (UInt32Value)36U };
            C.NumericValue numericValue239 = new C.NumericValue();
            numericValue239.Text = "28.3";

            numericPoint236.Append(numericValue239);

            C.NumericPoint numericPoint237 = new C.NumericPoint() { Index = (UInt32Value)37U };
            C.NumericValue numericValue240 = new C.NumericValue();
            numericValue240.Text = "28.3";

            numericPoint237.Append(numericValue240);

            C.NumericPoint numericPoint238 = new C.NumericPoint() { Index = (UInt32Value)38U };
            C.NumericValue numericValue241 = new C.NumericValue();
            numericValue241.Text = "28.3";

            numericPoint238.Append(numericValue241);

            C.NumericPoint numericPoint239 = new C.NumericPoint() { Index = (UInt32Value)39U };
            C.NumericValue numericValue242 = new C.NumericValue();
            numericValue242.Text = "28.3";

            numericPoint239.Append(numericValue242);

            numberingCache6.Append(formatCode6);
            numberingCache6.Append(pointCount6);
            numberingCache6.Append(numericPoint200);
            numberingCache6.Append(numericPoint201);
            numberingCache6.Append(numericPoint202);
            numberingCache6.Append(numericPoint203);
            numberingCache6.Append(numericPoint204);
            numberingCache6.Append(numericPoint205);
            numberingCache6.Append(numericPoint206);
            numberingCache6.Append(numericPoint207);
            numberingCache6.Append(numericPoint208);
            numberingCache6.Append(numericPoint209);
            numberingCache6.Append(numericPoint210);
            numberingCache6.Append(numericPoint211);
            numberingCache6.Append(numericPoint212);
            numberingCache6.Append(numericPoint213);
            numberingCache6.Append(numericPoint214);
            numberingCache6.Append(numericPoint215);
            numberingCache6.Append(numericPoint216);
            numberingCache6.Append(numericPoint217);
            numberingCache6.Append(numericPoint218);
            numberingCache6.Append(numericPoint219);
            numberingCache6.Append(numericPoint220);
            numberingCache6.Append(numericPoint221);
            numberingCache6.Append(numericPoint222);
            numberingCache6.Append(numericPoint223);
            numberingCache6.Append(numericPoint224);
            numberingCache6.Append(numericPoint225);
            numberingCache6.Append(numericPoint226);
            numberingCache6.Append(numericPoint227);
            numberingCache6.Append(numericPoint228);
            numberingCache6.Append(numericPoint229);
            numberingCache6.Append(numericPoint230);
            numberingCache6.Append(numericPoint231);
            numberingCache6.Append(numericPoint232);
            numberingCache6.Append(numericPoint233);
            numberingCache6.Append(numericPoint234);
            numberingCache6.Append(numericPoint235);
            numberingCache6.Append(numericPoint236);
            numberingCache6.Append(numericPoint237);
            numberingCache6.Append(numericPoint238);
            numberingCache6.Append(numericPoint239);

            numberReference6.Append(formula6);
            numberReference6.Append(numberingCache6);

            values3.Append(numberReference6);
            C.Smooth smooth1 = new C.Smooth() { Val = false };

            C.LineSerExtensionList lineSerExtensionList1 = new C.LineSerExtensionList();
            lineSerExtensionList1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");

            C.LineSerExtension lineSerExtension1 = new C.LineSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            lineSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000002-5AC9-435B-A402-5E3C89F8E9A8}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            lineSerExtension1.Append(openXmlUnknownElement3);

            lineSerExtensionList1.Append(lineSerExtension1);

            lineChartSeries1.Append(index3);
            lineChartSeries1.Append(order3);
            lineChartSeries1.Append(seriesText3);
            lineChartSeries1.Append(chartShapeProperties3);
            lineChartSeries1.Append(marker1);
            lineChartSeries1.Append(categoryAxisData3);
            lineChartSeries1.Append(values3);
            lineChartSeries1.Append(smooth1);
            lineChartSeries1.Append(lineSerExtensionList1);

            C.DataLabels dataLabels2 = new C.DataLabels();
            C.ShowLegendKey showLegendKey2 = new C.ShowLegendKey() { Val = false };
            C.ShowValue showValue2 = new C.ShowValue() { Val = false };
            C.ShowCategoryName showCategoryName2 = new C.ShowCategoryName() { Val = false };
            C.ShowSeriesName showSeriesName2 = new C.ShowSeriesName() { Val = false };
            C.ShowPercent showPercent2 = new C.ShowPercent() { Val = false };
            C.ShowBubbleSize showBubbleSize2 = new C.ShowBubbleSize() { Val = false };

            dataLabels2.Append(showLegendKey2);
            dataLabels2.Append(showValue2);
            dataLabels2.Append(showCategoryName2);
            dataLabels2.Append(showSeriesName2);
            dataLabels2.Append(showPercent2);
            dataLabels2.Append(showBubbleSize2);
            C.ShowMarker showMarker1 = new C.ShowMarker() { Val = true };
            C.Smooth smooth2 = new C.Smooth() { Val = false };
            C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)458742528U };
            C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)458741744U };

            lineChart1.Append(grouping1);
            lineChart1.Append(varyColors2);
            lineChart1.Append(lineChartSeries1);
            lineChart1.Append(dataLabels2);
            lineChart1.Append(showMarker1);
            lineChart1.Append(smooth2);
            lineChart1.Append(axisId3);
            lineChart1.Append(axisId4);

            C.CategoryAxis categoryAxis1 = new C.CategoryAxis();
            C.AxisId axisId5 = new C.AxisId() { Val = (UInt32Value)458742528U };

            C.Scaling scaling1 = new C.Scaling();
            C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };

            scaling1.Append(orientation1);
            C.Delete delete1 = new C.Delete() { Val = false };
            C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };

            C.Title title1 = new C.Title();

            C.ChartText chartText1 = new C.ChartText();

            C.RichText richText1 = new C.RichText();
            A.BodyProperties bodyProperties1 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill4 = new A.SolidFill();

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor3.Append(luminanceModulation1);
            schemeColor3.Append(luminanceOffset1);

            solidFill4.Append(schemeColor3);
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties1.Append(solidFill4);
            defaultRunProperties1.Append(latinFont1);
            defaultRunProperties1.Append(eastAsianFont1);
            defaultRunProperties1.Append(complexScriptFont1);

            paragraphProperties1.Append(defaultRunProperties1);

            A.Run run1 = new A.Run();
            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US" };
            A.Text text1 = new A.Text();
            text1.Text = "Week Start Date";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            richText1.Append(bodyProperties1);
            richText1.Append(listStyle1);
            richText1.Append(paragraph1);

            chartText1.Append(richText1);
            C.Layout layout2 = new C.Layout();
            C.Overlay overlay1 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties4 = new C.ChartShapeProperties();
            A.NoFill noFill3 = new A.NoFill();

            A.Outline outline4 = new A.Outline();
            A.NoFill noFill4 = new A.NoFill();

            outline4.Append(noFill4);
            A.EffectList effectList4 = new A.EffectList();

            chartShapeProperties4.Append(noFill3);
            chartShapeProperties4.Append(outline4);
            chartShapeProperties4.Append(effectList4);

            C.TextProperties textProperties1 = new C.TextProperties();
            A.BodyProperties bodyProperties2 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill5 = new A.SolidFill();

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor4.Append(luminanceModulation2);
            schemeColor4.Append(luminanceOffset2);

            solidFill5.Append(schemeColor4);
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill5);
            defaultRunProperties2.Append(latinFont2);
            defaultRunProperties2.Append(eastAsianFont2);
            defaultRunProperties2.Append(complexScriptFont2);

            paragraphProperties2.Append(defaultRunProperties2);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(endParagraphRunProperties1);

            textProperties1.Append(bodyProperties2);
            textProperties1.Append(listStyle2);
            textProperties1.Append(paragraph2);

            title1.Append(chartText1);
            title1.Append(layout2);
            title1.Append(overlay1);
            title1.Append(chartShapeProperties4);
            title1.Append(textProperties1);
            C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "[$-409]d\\-mmm;@", SourceLinked = false };
            C.MajorTickMark majorTickMark1 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark1 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

            C.ChartShapeProperties chartShapeProperties5 = new C.ChartShapeProperties();
            A.NoFill noFill5 = new A.NoFill();

            A.Outline outline5 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor5.Append(luminanceModulation3);
            schemeColor5.Append(luminanceOffset3);

            solidFill6.Append(schemeColor5);
            A.Round round2 = new A.Round();

            outline5.Append(solidFill6);
            outline5.Append(round2);
            A.EffectList effectList5 = new A.EffectList();

            chartShapeProperties5.Append(noFill5);
            chartShapeProperties5.Append(outline5);
            chartShapeProperties5.Append(effectList5);

            C.TextProperties textProperties2 = new C.TextProperties();
            A.BodyProperties bodyProperties3 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill7 = new A.SolidFill();

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor6.Append(luminanceModulation4);
            schemeColor6.Append(luminanceOffset4);

            solidFill7.Append(schemeColor6);
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties3.Append(solidFill7);
            defaultRunProperties3.Append(latinFont3);
            defaultRunProperties3.Append(eastAsianFont3);
            defaultRunProperties3.Append(complexScriptFont3);

            paragraphProperties3.Append(defaultRunProperties3);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(endParagraphRunProperties2);

            textProperties2.Append(bodyProperties3);
            textProperties2.Append(listStyle3);
            textProperties2.Append(paragraph3);
            C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)458741744U };
            C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.AutoLabeled autoLabeled1 = new C.AutoLabeled() { Val = false };
            C.LabelAlignment labelAlignment1 = new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center };
            C.LabelOffset labelOffset1 = new C.LabelOffset() { Val = (UInt16Value)100U };
            C.NoMultiLevelLabels noMultiLevelLabels1 = new C.NoMultiLevelLabels() { Val = false };

            categoryAxis1.Append(axisId5);
            categoryAxis1.Append(scaling1);
            categoryAxis1.Append(delete1);
            categoryAxis1.Append(axisPosition1);
            categoryAxis1.Append(title1);
            categoryAxis1.Append(numberingFormat1);
            categoryAxis1.Append(majorTickMark1);
            categoryAxis1.Append(minorTickMark1);
            categoryAxis1.Append(tickLabelPosition1);
            categoryAxis1.Append(chartShapeProperties5);
            categoryAxis1.Append(textProperties2);
            categoryAxis1.Append(crossingAxis1);
            categoryAxis1.Append(crosses1);
            categoryAxis1.Append(autoLabeled1);
            categoryAxis1.Append(labelAlignment1);
            categoryAxis1.Append(labelOffset1);
            categoryAxis1.Append(noMultiLevelLabels1);

            C.ValueAxis valueAxis1 = new C.ValueAxis();
            C.AxisId axisId6 = new C.AxisId() { Val = (UInt32Value)458741744U };

            C.Scaling scaling2 = new C.Scaling();
            C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };

            scaling2.Append(orientation2);
            C.Delete delete2 = new C.Delete() { Val = false };
            C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };

            C.MajorGridlines majorGridlines1 = new C.MajorGridlines();

            C.ChartShapeProperties chartShapeProperties6 = new C.ChartShapeProperties();

            A.Outline outline6 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill8 = new A.SolidFill();

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor7.Append(luminanceModulation5);
            schemeColor7.Append(luminanceOffset5);

            solidFill8.Append(schemeColor7);
            A.Round round3 = new A.Round();

            outline6.Append(solidFill8);
            outline6.Append(round3);
            A.EffectList effectList6 = new A.EffectList();

            chartShapeProperties6.Append(outline6);
            chartShapeProperties6.Append(effectList6);

            majorGridlines1.Append(chartShapeProperties6);

            C.Title title2 = new C.Title();

            C.ChartText chartText2 = new C.ChartText();

            C.RichText richText2 = new C.RichText();
            A.BodyProperties bodyProperties4 = new A.BodyProperties() { Rotation = -5400000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill9 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset6 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor8.Append(luminanceModulation6);
            schemeColor8.Append(luminanceOffset6);

            solidFill9.Append(schemeColor8);
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties4.Append(solidFill9);
            defaultRunProperties4.Append(latinFont4);
            defaultRunProperties4.Append(eastAsianFont4);
            defaultRunProperties4.Append(complexScriptFont4);

            paragraphProperties4.Append(defaultRunProperties4);

            A.Run run2 = new A.Run();
            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US" };
            A.Text text2 = new A.Text();
            text2.Text = "Flow";

            run2.Append(runProperties2);
            run2.Append(text2);

            A.Run run3 = new A.Run();
            A.RunProperties runProperties3 = new A.RunProperties() { Language = "en-US", Baseline = 0 };
            A.Text text3 = new A.Text();
            text3.Text = " Rate (GPM)";

            run3.Append(runProperties3);
            run3.Append(text3);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run2);
            paragraph4.Append(run3);
            paragraph4.Append(endParagraphRunProperties3);

            richText2.Append(bodyProperties4);
            richText2.Append(listStyle4);
            richText2.Append(paragraph4);

            chartText2.Append(richText2);
            C.Layout layout3 = new C.Layout();
            C.Overlay overlay2 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties7 = new C.ChartShapeProperties();
            A.NoFill noFill6 = new A.NoFill();

            A.Outline outline7 = new A.Outline();
            A.NoFill noFill7 = new A.NoFill();

            outline7.Append(noFill7);
            A.EffectList effectList7 = new A.EffectList();

            chartShapeProperties7.Append(noFill6);
            chartShapeProperties7.Append(outline7);
            chartShapeProperties7.Append(effectList7);

            C.TextProperties textProperties3 = new C.TextProperties();
            A.BodyProperties bodyProperties5 = new A.BodyProperties() { Rotation = -5400000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties() { FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill10 = new A.SolidFill();

            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset7 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor9.Append(luminanceModulation7);
            schemeColor9.Append(luminanceOffset7);

            solidFill10.Append(schemeColor9);
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties5.Append(solidFill10);
            defaultRunProperties5.Append(latinFont5);
            defaultRunProperties5.Append(eastAsianFont5);
            defaultRunProperties5.Append(complexScriptFont5);

            paragraphProperties5.Append(defaultRunProperties5);
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(endParagraphRunProperties4);

            textProperties3.Append(bodyProperties5);
            textProperties3.Append(listStyle5);
            textProperties3.Append(paragraph5);

            title2.Append(chartText2);
            title2.Append(layout3);
            title2.Append(overlay2);
            title2.Append(chartShapeProperties7);
            title2.Append(textProperties3);
            C.NumberingFormat numberingFormat2 = new C.NumberingFormat() { FormatCode = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)", SourceLinked = true };
            C.MajorTickMark majorTickMark2 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark2 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

            C.ChartShapeProperties chartShapeProperties8 = new C.ChartShapeProperties();
            A.NoFill noFill8 = new A.NoFill();

            A.Outline outline8 = new A.Outline();
            A.NoFill noFill9 = new A.NoFill();

            outline8.Append(noFill9);
            A.EffectList effectList8 = new A.EffectList();

            chartShapeProperties8.Append(noFill8);
            chartShapeProperties8.Append(outline8);
            chartShapeProperties8.Append(effectList8);

            C.TextProperties textProperties4 = new C.TextProperties();
            A.BodyProperties bodyProperties6 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle6 = new A.ListStyle();

            A.Paragraph paragraph6 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties6 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill11 = new A.SolidFill();

            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset8 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor10.Append(luminanceModulation8);
            schemeColor10.Append(luminanceOffset8);

            solidFill11.Append(schemeColor10);
            A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties6.Append(solidFill11);
            defaultRunProperties6.Append(latinFont6);
            defaultRunProperties6.Append(eastAsianFont6);
            defaultRunProperties6.Append(complexScriptFont6);

            paragraphProperties6.Append(defaultRunProperties6);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(endParagraphRunProperties5);

            textProperties4.Append(bodyProperties6);
            textProperties4.Append(listStyle6);
            textProperties4.Append(paragraph6);
            C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)458742528U };
            C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.Between };

            valueAxis1.Append(axisId6);
            valueAxis1.Append(scaling2);
            valueAxis1.Append(delete2);
            valueAxis1.Append(axisPosition2);
            valueAxis1.Append(majorGridlines1);
            valueAxis1.Append(title2);
            valueAxis1.Append(numberingFormat2);
            valueAxis1.Append(majorTickMark2);
            valueAxis1.Append(minorTickMark2);
            valueAxis1.Append(tickLabelPosition2);
            valueAxis1.Append(chartShapeProperties8);
            valueAxis1.Append(textProperties4);
            valueAxis1.Append(crossingAxis2);
            valueAxis1.Append(crosses2);
            valueAxis1.Append(crossBetween1);

            C.ShapeProperties shapeProperties1 = new C.ShapeProperties();
            A.NoFill noFill10 = new A.NoFill();

            A.Outline outline9 = new A.Outline();
            A.NoFill noFill11 = new A.NoFill();

            outline9.Append(noFill11);
            A.EffectList effectList9 = new A.EffectList();

            shapeProperties1.Append(noFill10);
            shapeProperties1.Append(outline9);
            shapeProperties1.Append(effectList9);

            plotArea1.Append(layout1);
            plotArea1.Append(barChart1);
            plotArea1.Append(lineChart1);
            plotArea1.Append(categoryAxis1);
            plotArea1.Append(valueAxis1);
            plotArea1.Append(shapeProperties1);

            C.Legend legend1 = new C.Legend();
            C.LegendPosition legendPosition1 = new C.LegendPosition() { Val = C.LegendPositionValues.Bottom };

            C.LegendEntry legendEntry1 = new C.LegendEntry();
            C.Index index4 = new C.Index() { Val = (UInt32Value)0U };

            C.TextProperties textProperties5 = new C.TextProperties();
            A.BodyProperties bodyProperties7 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle7 = new A.ListStyle();

            A.Paragraph paragraph7 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties() { FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill12 = new A.SolidFill();

            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset9 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor11.Append(luminanceModulation9);
            schemeColor11.Append(luminanceOffset9);

            solidFill12.Append(schemeColor11);
            A.LatinFont latinFont7 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont7 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont7 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties7.Append(solidFill12);
            defaultRunProperties7.Append(latinFont7);
            defaultRunProperties7.Append(eastAsianFont7);
            defaultRunProperties7.Append(complexScriptFont7);

            paragraphProperties7.Append(defaultRunProperties7);
            A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(endParagraphRunProperties6);

            textProperties5.Append(bodyProperties7);
            textProperties5.Append(listStyle7);
            textProperties5.Append(paragraph7);

            legendEntry1.Append(index4);
            legendEntry1.Append(textProperties5);

            C.LegendEntry legendEntry2 = new C.LegendEntry();
            C.Index index5 = new C.Index() { Val = (UInt32Value)1U };

            C.TextProperties textProperties6 = new C.TextProperties();
            A.BodyProperties bodyProperties8 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle8 = new A.ListStyle();

            A.Paragraph paragraph8 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties8 = new A.DefaultRunProperties() { FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill13 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset10 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor12.Append(luminanceModulation10);
            schemeColor12.Append(luminanceOffset10);

            solidFill13.Append(schemeColor12);
            A.LatinFont latinFont8 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont8 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont8 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties8.Append(solidFill13);
            defaultRunProperties8.Append(latinFont8);
            defaultRunProperties8.Append(eastAsianFont8);
            defaultRunProperties8.Append(complexScriptFont8);

            paragraphProperties8.Append(defaultRunProperties8);
            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(endParagraphRunProperties7);

            textProperties6.Append(bodyProperties8);
            textProperties6.Append(listStyle8);
            textProperties6.Append(paragraph8);

            legendEntry2.Append(index5);
            legendEntry2.Append(textProperties6);
            C.Layout layout4 = new C.Layout();
            C.Overlay overlay3 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties9 = new C.ChartShapeProperties();
            A.NoFill noFill12 = new A.NoFill();

            A.Outline outline10 = new A.Outline();
            A.NoFill noFill13 = new A.NoFill();

            outline10.Append(noFill13);
            A.EffectList effectList10 = new A.EffectList();

            chartShapeProperties9.Append(noFill12);
            chartShapeProperties9.Append(outline10);
            chartShapeProperties9.Append(effectList10);

            C.TextProperties textProperties7 = new C.TextProperties();
            A.BodyProperties bodyProperties9 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle9 = new A.ListStyle();

            A.Paragraph paragraph9 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties9 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill14 = new A.SolidFill();

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset11 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor13.Append(luminanceModulation11);
            schemeColor13.Append(luminanceOffset11);

            solidFill14.Append(schemeColor13);
            A.LatinFont latinFont9 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont9 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont9 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties9.Append(solidFill14);
            defaultRunProperties9.Append(latinFont9);
            defaultRunProperties9.Append(eastAsianFont9);
            defaultRunProperties9.Append(complexScriptFont9);

            paragraphProperties9.Append(defaultRunProperties9);
            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(endParagraphRunProperties8);

            textProperties7.Append(bodyProperties9);
            textProperties7.Append(listStyle9);
            textProperties7.Append(paragraph9);

            legend1.Append(legendPosition1);
            legend1.Append(legendEntry1);
            legend1.Append(legendEntry2);
            legend1.Append(layout4);
            legend1.Append(overlay3);
            legend1.Append(chartShapeProperties9);
            legend1.Append(textProperties7);
            C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };
            C.DisplayBlanksAs displayBlanksAs1 = new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Gap };
            C.ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new C.ShowDataLabelsOverMaximum() { Val = false };

            chart1.Append(autoTitleDeleted1);
            chart1.Append(plotArea1);
            chart1.Append(legend1);
            chart1.Append(plotVisibleOnly1);
            chart1.Append(displayBlanksAs1);
            chart1.Append(showDataLabelsOverMaximum1);

            C.ShapeProperties shapeProperties2 = new C.ShapeProperties();

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill15.Append(schemeColor14);

            A.Outline outline11 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill16 = new A.SolidFill();

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset12 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor15.Append(luminanceModulation12);
            schemeColor15.Append(luminanceOffset12);

            solidFill16.Append(schemeColor15);
            A.Round round4 = new A.Round();

            outline11.Append(solidFill16);
            outline11.Append(round4);
            A.EffectList effectList11 = new A.EffectList();

            shapeProperties2.Append(solidFill15);
            shapeProperties2.Append(outline11);
            shapeProperties2.Append(effectList11);

            C.TextProperties textProperties8 = new C.TextProperties();
            A.BodyProperties bodyProperties10 = new A.BodyProperties();
            A.ListStyle listStyle10 = new A.ListStyle();

            A.Paragraph paragraph10 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties10 = new A.ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties10 = new A.DefaultRunProperties();

            paragraphProperties10.Append(defaultRunProperties10);
            A.EndParagraphRunProperties endParagraphRunProperties9 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(endParagraphRunProperties9);

            textProperties8.Append(bodyProperties10);
            textProperties8.Append(listStyle10);
            textProperties8.Append(paragraph10);

            chartSpace1.Append(date19041);
            chartSpace1.Append(editingLanguage1);
            chartSpace1.Append(roundedCorners1);
            chartSpace1.Append(alternateContent1);
            chartSpace1.Append(chart1);
            chartSpace1.Append(shapeProperties2);
            chartSpace1.Append(textProperties8);

            chartPart1.ChartSpace = chartSpace1;
        }

        // Generates content of chartColorStylePart1.
        private void GenerateChartColorStylePart1Content(ChartColorStylePart chartColorStylePart1)
        {
            Cs.ColorStyle colorStyle1 = new Cs.ColorStyle() { Method = "cycle", Id = (UInt32Value)10U };
            colorStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
            colorStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };
            A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent3 };
            A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent4 };
            A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent5 };
            A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent6 };
            Cs.ColorStyleVariation colorStyleVariation1 = new Cs.ColorStyleVariation();

            Cs.ColorStyleVariation colorStyleVariation2 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation13 = new A.LuminanceModulation() { Val = 60000 };

            colorStyleVariation2.Append(luminanceModulation13);

            Cs.ColorStyleVariation colorStyleVariation3 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation14 = new A.LuminanceModulation() { Val = 80000 };
            A.LuminanceOffset luminanceOffset13 = new A.LuminanceOffset() { Val = 20000 };

            colorStyleVariation3.Append(luminanceModulation14);
            colorStyleVariation3.Append(luminanceOffset13);

            Cs.ColorStyleVariation colorStyleVariation4 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation15 = new A.LuminanceModulation() { Val = 80000 };

            colorStyleVariation4.Append(luminanceModulation15);

            Cs.ColorStyleVariation colorStyleVariation5 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation16 = new A.LuminanceModulation() { Val = 60000 };
            A.LuminanceOffset luminanceOffset14 = new A.LuminanceOffset() { Val = 40000 };

            colorStyleVariation5.Append(luminanceModulation16);
            colorStyleVariation5.Append(luminanceOffset14);

            Cs.ColorStyleVariation colorStyleVariation6 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation17 = new A.LuminanceModulation() { Val = 50000 };

            colorStyleVariation6.Append(luminanceModulation17);

            Cs.ColorStyleVariation colorStyleVariation7 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation18 = new A.LuminanceModulation() { Val = 70000 };
            A.LuminanceOffset luminanceOffset15 = new A.LuminanceOffset() { Val = 30000 };

            colorStyleVariation7.Append(luminanceModulation18);
            colorStyleVariation7.Append(luminanceOffset15);

            Cs.ColorStyleVariation colorStyleVariation8 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation19 = new A.LuminanceModulation() { Val = 70000 };

            colorStyleVariation8.Append(luminanceModulation19);

            Cs.ColorStyleVariation colorStyleVariation9 = new Cs.ColorStyleVariation();
            A.LuminanceModulation luminanceModulation20 = new A.LuminanceModulation() { Val = 50000 };
            A.LuminanceOffset luminanceOffset16 = new A.LuminanceOffset() { Val = 50000 };

            colorStyleVariation9.Append(luminanceModulation20);
            colorStyleVariation9.Append(luminanceOffset16);

            colorStyle1.Append(schemeColor16);
            colorStyle1.Append(schemeColor17);
            colorStyle1.Append(schemeColor18);
            colorStyle1.Append(schemeColor19);
            colorStyle1.Append(schemeColor20);
            colorStyle1.Append(schemeColor21);
            colorStyle1.Append(colorStyleVariation1);
            colorStyle1.Append(colorStyleVariation2);
            colorStyle1.Append(colorStyleVariation3);
            colorStyle1.Append(colorStyleVariation4);
            colorStyle1.Append(colorStyleVariation5);
            colorStyle1.Append(colorStyleVariation6);
            colorStyle1.Append(colorStyleVariation7);
            colorStyle1.Append(colorStyleVariation8);
            colorStyle1.Append(colorStyleVariation9);

            chartColorStylePart1.ColorStyle = colorStyle1;
        }

        // Generates content of chartStylePart1.
        private void GenerateChartStylePart1Content(ChartStylePart chartStylePart1)
        {
            Cs.ChartStyle chartStyle1 = new Cs.ChartStyle() { Id = (UInt32Value)201U };
            chartStyle1.AddNamespaceDeclaration("cs", "http://schemas.microsoft.com/office/drawing/2012/chartStyle");
            chartStyle1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Cs.AxisTitle axisTitle1 = new Cs.AxisTitle();
            Cs.LineReference lineReference1 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference1 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference1 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference1 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor22 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation21 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset17 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor22.Append(luminanceModulation21);
            schemeColor22.Append(luminanceOffset17);

            fontReference1.Append(schemeColor22);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType1 = new Cs.TextCharacterPropertiesType() { FontSize = 1000, Kerning = 1200 };

            axisTitle1.Append(lineReference1);
            axisTitle1.Append(fillReference1);
            axisTitle1.Append(effectReference1);
            axisTitle1.Append(fontReference1);
            axisTitle1.Append(textCharacterPropertiesType1);

            Cs.CategoryAxis categoryAxis2 = new Cs.CategoryAxis();
            Cs.LineReference lineReference2 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference2 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference2 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference2 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor23 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation22 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset18 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor23.Append(luminanceModulation22);
            schemeColor23.Append(luminanceOffset18);

            fontReference2.Append(schemeColor23);

            Cs.ShapeProperties shapeProperties3 = new Cs.ShapeProperties();

            A.Outline outline12 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill17 = new A.SolidFill();

            A.SchemeColor schemeColor24 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation23 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset19 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor24.Append(luminanceModulation23);
            schemeColor24.Append(luminanceOffset19);

            solidFill17.Append(schemeColor24);
            A.Round round5 = new A.Round();

            outline12.Append(solidFill17);
            outline12.Append(round5);

            shapeProperties3.Append(outline12);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType2 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            categoryAxis2.Append(lineReference2);
            categoryAxis2.Append(fillReference2);
            categoryAxis2.Append(effectReference2);
            categoryAxis2.Append(fontReference2);
            categoryAxis2.Append(shapeProperties3);
            categoryAxis2.Append(textCharacterPropertiesType2);

            Cs.ChartArea chartArea1 = new Cs.ChartArea() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
            Cs.LineReference lineReference3 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference3 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference3 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference3 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor25 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference3.Append(schemeColor25);

            Cs.ShapeProperties shapeProperties4 = new Cs.ShapeProperties();

            A.SolidFill solidFill18 = new A.SolidFill();
            A.SchemeColor schemeColor26 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill18.Append(schemeColor26);

            A.Outline outline13 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill19 = new A.SolidFill();

            A.SchemeColor schemeColor27 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation24 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset20 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor27.Append(luminanceModulation24);
            schemeColor27.Append(luminanceOffset20);

            solidFill19.Append(schemeColor27);
            A.Round round6 = new A.Round();

            outline13.Append(solidFill19);
            outline13.Append(round6);

            shapeProperties4.Append(solidFill18);
            shapeProperties4.Append(outline13);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType3 = new Cs.TextCharacterPropertiesType() { FontSize = 1000, Kerning = 1200 };

            chartArea1.Append(lineReference3);
            chartArea1.Append(fillReference3);
            chartArea1.Append(effectReference3);
            chartArea1.Append(fontReference3);
            chartArea1.Append(shapeProperties4);
            chartArea1.Append(textCharacterPropertiesType3);

            Cs.DataLabel dataLabel1 = new Cs.DataLabel();
            Cs.LineReference lineReference4 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference4 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference4 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference4 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor28 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation25 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset21 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor28.Append(luminanceModulation25);
            schemeColor28.Append(luminanceOffset21);

            fontReference4.Append(schemeColor28);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType4 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            dataLabel1.Append(lineReference4);
            dataLabel1.Append(fillReference4);
            dataLabel1.Append(effectReference4);
            dataLabel1.Append(fontReference4);
            dataLabel1.Append(textCharacterPropertiesType4);

            Cs.DataLabelCallout dataLabelCallout1 = new Cs.DataLabelCallout();
            Cs.LineReference lineReference5 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference5 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference5 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference5 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor29 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation26 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset22 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor29.Append(luminanceModulation26);
            schemeColor29.Append(luminanceOffset22);

            fontReference5.Append(schemeColor29);

            Cs.ShapeProperties shapeProperties5 = new Cs.ShapeProperties();

            A.SolidFill solidFill20 = new A.SolidFill();
            A.SchemeColor schemeColor30 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            solidFill20.Append(schemeColor30);

            A.Outline outline14 = new A.Outline();

            A.SolidFill solidFill21 = new A.SolidFill();

            A.SchemeColor schemeColor31 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation27 = new A.LuminanceModulation() { Val = 25000 };
            A.LuminanceOffset luminanceOffset23 = new A.LuminanceOffset() { Val = 75000 };

            schemeColor31.Append(luminanceModulation27);
            schemeColor31.Append(luminanceOffset23);

            solidFill21.Append(schemeColor31);

            outline14.Append(solidFill21);

            shapeProperties5.Append(solidFill20);
            shapeProperties5.Append(outline14);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType5 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            Cs.TextBodyProperties textBodyProperties1 = new Cs.TextBodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Clip, HorizontalOverflow = A.TextHorizontalOverflowValues.Clip, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 36576, TopInset = 18288, RightInset = 36576, BottomInset = 18288, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

            textBodyProperties1.Append(shapeAutoFit1);

            dataLabelCallout1.Append(lineReference5);
            dataLabelCallout1.Append(fillReference5);
            dataLabelCallout1.Append(effectReference5);
            dataLabelCallout1.Append(fontReference5);
            dataLabelCallout1.Append(shapeProperties5);
            dataLabelCallout1.Append(textCharacterPropertiesType5);
            dataLabelCallout1.Append(textBodyProperties1);

            Cs.DataPoint dataPoint1 = new Cs.DataPoint();
            Cs.LineReference lineReference6 = new Cs.LineReference() { Index = (UInt32Value)0U };

            Cs.FillReference fillReference6 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.StyleColor styleColor1 = new Cs.StyleColor() { Val = "auto" };

            fillReference6.Append(styleColor1);
            Cs.EffectReference effectReference6 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference6 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor32 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference6.Append(schemeColor32);

            dataPoint1.Append(lineReference6);
            dataPoint1.Append(fillReference6);
            dataPoint1.Append(effectReference6);
            dataPoint1.Append(fontReference6);

            Cs.DataPoint3D dataPoint3D1 = new Cs.DataPoint3D();
            Cs.LineReference lineReference7 = new Cs.LineReference() { Index = (UInt32Value)0U };

            Cs.FillReference fillReference7 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.StyleColor styleColor2 = new Cs.StyleColor() { Val = "auto" };

            fillReference7.Append(styleColor2);
            Cs.EffectReference effectReference7 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference7 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor33 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference7.Append(schemeColor33);

            dataPoint3D1.Append(lineReference7);
            dataPoint3D1.Append(fillReference7);
            dataPoint3D1.Append(effectReference7);
            dataPoint3D1.Append(fontReference7);

            Cs.DataPointLine dataPointLine1 = new Cs.DataPointLine();

            Cs.LineReference lineReference8 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor3 = new Cs.StyleColor() { Val = "auto" };

            lineReference8.Append(styleColor3);
            Cs.FillReference fillReference8 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.EffectReference effectReference8 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference8 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor34 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference8.Append(schemeColor34);

            Cs.ShapeProperties shapeProperties6 = new Cs.ShapeProperties();

            A.Outline outline15 = new A.Outline() { Width = 28575, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill22 = new A.SolidFill();
            A.SchemeColor schemeColor35 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill22.Append(schemeColor35);
            A.Round round7 = new A.Round();

            outline15.Append(solidFill22);
            outline15.Append(round7);

            shapeProperties6.Append(outline15);

            dataPointLine1.Append(lineReference8);
            dataPointLine1.Append(fillReference8);
            dataPointLine1.Append(effectReference8);
            dataPointLine1.Append(fontReference8);
            dataPointLine1.Append(shapeProperties6);

            Cs.DataPointMarker dataPointMarker1 = new Cs.DataPointMarker();

            Cs.LineReference lineReference9 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor4 = new Cs.StyleColor() { Val = "auto" };

            lineReference9.Append(styleColor4);

            Cs.FillReference fillReference9 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.StyleColor styleColor5 = new Cs.StyleColor() { Val = "auto" };

            fillReference9.Append(styleColor5);
            Cs.EffectReference effectReference9 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference9 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor36 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference9.Append(schemeColor36);

            Cs.ShapeProperties shapeProperties7 = new Cs.ShapeProperties();

            A.Outline outline16 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill23 = new A.SolidFill();
            A.SchemeColor schemeColor37 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill23.Append(schemeColor37);

            outline16.Append(solidFill23);

            shapeProperties7.Append(outline16);

            dataPointMarker1.Append(lineReference9);
            dataPointMarker1.Append(fillReference9);
            dataPointMarker1.Append(effectReference9);
            dataPointMarker1.Append(fontReference9);
            dataPointMarker1.Append(shapeProperties7);
            Cs.MarkerLayoutProperties markerLayoutProperties1 = new Cs.MarkerLayoutProperties() { Symbol = Cs.MarkerStyle.Circle, Size = 5 };

            Cs.DataPointWireframe dataPointWireframe1 = new Cs.DataPointWireframe();

            Cs.LineReference lineReference10 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor6 = new Cs.StyleColor() { Val = "auto" };

            lineReference10.Append(styleColor6);
            Cs.FillReference fillReference10 = new Cs.FillReference() { Index = (UInt32Value)1U };
            Cs.EffectReference effectReference10 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference10 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor38 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference10.Append(schemeColor38);

            Cs.ShapeProperties shapeProperties8 = new Cs.ShapeProperties();

            A.Outline outline17 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill24 = new A.SolidFill();
            A.SchemeColor schemeColor39 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill24.Append(schemeColor39);
            A.Round round8 = new A.Round();

            outline17.Append(solidFill24);
            outline17.Append(round8);

            shapeProperties8.Append(outline17);

            dataPointWireframe1.Append(lineReference10);
            dataPointWireframe1.Append(fillReference10);
            dataPointWireframe1.Append(effectReference10);
            dataPointWireframe1.Append(fontReference10);
            dataPointWireframe1.Append(shapeProperties8);

            Cs.DataTableStyle dataTableStyle1 = new Cs.DataTableStyle();
            Cs.LineReference lineReference11 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference11 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference11 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference11 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor40 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation28 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset24 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor40.Append(luminanceModulation28);
            schemeColor40.Append(luminanceOffset24);

            fontReference11.Append(schemeColor40);

            Cs.ShapeProperties shapeProperties9 = new Cs.ShapeProperties();
            A.NoFill noFill14 = new A.NoFill();

            A.Outline outline18 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill25 = new A.SolidFill();

            A.SchemeColor schemeColor41 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation29 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset25 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor41.Append(luminanceModulation29);
            schemeColor41.Append(luminanceOffset25);

            solidFill25.Append(schemeColor41);
            A.Round round9 = new A.Round();

            outline18.Append(solidFill25);
            outline18.Append(round9);

            shapeProperties9.Append(noFill14);
            shapeProperties9.Append(outline18);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType6 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            dataTableStyle1.Append(lineReference11);
            dataTableStyle1.Append(fillReference11);
            dataTableStyle1.Append(effectReference11);
            dataTableStyle1.Append(fontReference11);
            dataTableStyle1.Append(shapeProperties9);
            dataTableStyle1.Append(textCharacterPropertiesType6);

            Cs.DownBar downBar1 = new Cs.DownBar();
            Cs.LineReference lineReference12 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference12 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference12 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference12 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor42 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference12.Append(schemeColor42);

            Cs.ShapeProperties shapeProperties10 = new Cs.ShapeProperties();

            A.SolidFill solidFill26 = new A.SolidFill();

            A.SchemeColor schemeColor43 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
            A.LuminanceModulation luminanceModulation30 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset26 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor43.Append(luminanceModulation30);
            schemeColor43.Append(luminanceOffset26);

            solidFill26.Append(schemeColor43);

            A.Outline outline19 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill27 = new A.SolidFill();

            A.SchemeColor schemeColor44 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation31 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset27 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor44.Append(luminanceModulation31);
            schemeColor44.Append(luminanceOffset27);

            solidFill27.Append(schemeColor44);

            outline19.Append(solidFill27);

            shapeProperties10.Append(solidFill26);
            shapeProperties10.Append(outline19);

            downBar1.Append(lineReference12);
            downBar1.Append(fillReference12);
            downBar1.Append(effectReference12);
            downBar1.Append(fontReference12);
            downBar1.Append(shapeProperties10);

            Cs.DropLine dropLine1 = new Cs.DropLine();
            Cs.LineReference lineReference13 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference13 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference13 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference13 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor45 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference13.Append(schemeColor45);

            Cs.ShapeProperties shapeProperties11 = new Cs.ShapeProperties();

            A.Outline outline20 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill28 = new A.SolidFill();

            A.SchemeColor schemeColor46 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation32 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset28 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor46.Append(luminanceModulation32);
            schemeColor46.Append(luminanceOffset28);

            solidFill28.Append(schemeColor46);
            A.Round round10 = new A.Round();

            outline20.Append(solidFill28);
            outline20.Append(round10);

            shapeProperties11.Append(outline20);

            dropLine1.Append(lineReference13);
            dropLine1.Append(fillReference13);
            dropLine1.Append(effectReference13);
            dropLine1.Append(fontReference13);
            dropLine1.Append(shapeProperties11);

            Cs.ErrorBar errorBar1 = new Cs.ErrorBar();
            Cs.LineReference lineReference14 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference14 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference14 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference14 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor47 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference14.Append(schemeColor47);

            Cs.ShapeProperties shapeProperties12 = new Cs.ShapeProperties();

            A.Outline outline21 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill29 = new A.SolidFill();

            A.SchemeColor schemeColor48 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation33 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset29 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor48.Append(luminanceModulation33);
            schemeColor48.Append(luminanceOffset29);

            solidFill29.Append(schemeColor48);
            A.Round round11 = new A.Round();

            outline21.Append(solidFill29);
            outline21.Append(round11);

            shapeProperties12.Append(outline21);

            errorBar1.Append(lineReference14);
            errorBar1.Append(fillReference14);
            errorBar1.Append(effectReference14);
            errorBar1.Append(fontReference14);
            errorBar1.Append(shapeProperties12);

            Cs.Floor floor1 = new Cs.Floor();
            Cs.LineReference lineReference15 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference15 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference15 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference15 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor49 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference15.Append(schemeColor49);

            Cs.ShapeProperties shapeProperties13 = new Cs.ShapeProperties();
            A.NoFill noFill15 = new A.NoFill();

            A.Outline outline22 = new A.Outline();
            A.NoFill noFill16 = new A.NoFill();

            outline22.Append(noFill16);

            shapeProperties13.Append(noFill15);
            shapeProperties13.Append(outline22);

            floor1.Append(lineReference15);
            floor1.Append(fillReference15);
            floor1.Append(effectReference15);
            floor1.Append(fontReference15);
            floor1.Append(shapeProperties13);

            Cs.GridlineMajor gridlineMajor1 = new Cs.GridlineMajor();
            Cs.LineReference lineReference16 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference16 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference16 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference16 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor50 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference16.Append(schemeColor50);

            Cs.ShapeProperties shapeProperties14 = new Cs.ShapeProperties();

            A.Outline outline23 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill30 = new A.SolidFill();

            A.SchemeColor schemeColor51 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation34 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset30 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor51.Append(luminanceModulation34);
            schemeColor51.Append(luminanceOffset30);

            solidFill30.Append(schemeColor51);
            A.Round round12 = new A.Round();

            outline23.Append(solidFill30);
            outline23.Append(round12);

            shapeProperties14.Append(outline23);

            gridlineMajor1.Append(lineReference16);
            gridlineMajor1.Append(fillReference16);
            gridlineMajor1.Append(effectReference16);
            gridlineMajor1.Append(fontReference16);
            gridlineMajor1.Append(shapeProperties14);

            Cs.GridlineMinor gridlineMinor1 = new Cs.GridlineMinor();
            Cs.LineReference lineReference17 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference17 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference17 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference17 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor52 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference17.Append(schemeColor52);

            Cs.ShapeProperties shapeProperties15 = new Cs.ShapeProperties();

            A.Outline outline24 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill31 = new A.SolidFill();

            A.SchemeColor schemeColor53 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation35 = new A.LuminanceModulation() { Val = 5000 };
            A.LuminanceOffset luminanceOffset31 = new A.LuminanceOffset() { Val = 95000 };

            schemeColor53.Append(luminanceModulation35);
            schemeColor53.Append(luminanceOffset31);

            solidFill31.Append(schemeColor53);
            A.Round round13 = new A.Round();

            outline24.Append(solidFill31);
            outline24.Append(round13);

            shapeProperties15.Append(outline24);

            gridlineMinor1.Append(lineReference17);
            gridlineMinor1.Append(fillReference17);
            gridlineMinor1.Append(effectReference17);
            gridlineMinor1.Append(fontReference17);
            gridlineMinor1.Append(shapeProperties15);

            Cs.HiLoLine hiLoLine1 = new Cs.HiLoLine();
            Cs.LineReference lineReference18 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference18 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference18 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference18 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor54 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference18.Append(schemeColor54);

            Cs.ShapeProperties shapeProperties16 = new Cs.ShapeProperties();

            A.Outline outline25 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill32 = new A.SolidFill();

            A.SchemeColor schemeColor55 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation36 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset32 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor55.Append(luminanceModulation36);
            schemeColor55.Append(luminanceOffset32);

            solidFill32.Append(schemeColor55);
            A.Round round14 = new A.Round();

            outline25.Append(solidFill32);
            outline25.Append(round14);

            shapeProperties16.Append(outline25);

            hiLoLine1.Append(lineReference18);
            hiLoLine1.Append(fillReference18);
            hiLoLine1.Append(effectReference18);
            hiLoLine1.Append(fontReference18);
            hiLoLine1.Append(shapeProperties16);

            Cs.LeaderLine leaderLine1 = new Cs.LeaderLine();
            Cs.LineReference lineReference19 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference19 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference19 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference19 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor56 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference19.Append(schemeColor56);

            Cs.ShapeProperties shapeProperties17 = new Cs.ShapeProperties();

            A.Outline outline26 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill33 = new A.SolidFill();

            A.SchemeColor schemeColor57 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation37 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset33 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor57.Append(luminanceModulation37);
            schemeColor57.Append(luminanceOffset33);

            solidFill33.Append(schemeColor57);
            A.Round round15 = new A.Round();

            outline26.Append(solidFill33);
            outline26.Append(round15);

            shapeProperties17.Append(outline26);

            leaderLine1.Append(lineReference19);
            leaderLine1.Append(fillReference19);
            leaderLine1.Append(effectReference19);
            leaderLine1.Append(fontReference19);
            leaderLine1.Append(shapeProperties17);

            Cs.LegendStyle legendStyle1 = new Cs.LegendStyle();
            Cs.LineReference lineReference20 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference20 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference20 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference20 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor58 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation38 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset34 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor58.Append(luminanceModulation38);
            schemeColor58.Append(luminanceOffset34);

            fontReference20.Append(schemeColor58);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType7 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            legendStyle1.Append(lineReference20);
            legendStyle1.Append(fillReference20);
            legendStyle1.Append(effectReference20);
            legendStyle1.Append(fontReference20);
            legendStyle1.Append(textCharacterPropertiesType7);

            Cs.PlotArea plotArea2 = new Cs.PlotArea() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
            Cs.LineReference lineReference21 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference21 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference21 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference21 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor59 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference21.Append(schemeColor59);

            plotArea2.Append(lineReference21);
            plotArea2.Append(fillReference21);
            plotArea2.Append(effectReference21);
            plotArea2.Append(fontReference21);

            Cs.PlotArea3D plotArea3D1 = new Cs.PlotArea3D() { Modifiers = new ListValue<StringValue>() { InnerText = "allowNoFillOverride allowNoLineOverride" } };
            Cs.LineReference lineReference22 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference22 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference22 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference22 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor60 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference22.Append(schemeColor60);

            plotArea3D1.Append(lineReference22);
            plotArea3D1.Append(fillReference22);
            plotArea3D1.Append(effectReference22);
            plotArea3D1.Append(fontReference22);

            Cs.SeriesAxis seriesAxis1 = new Cs.SeriesAxis();
            Cs.LineReference lineReference23 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference23 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference23 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference23 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor61 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation39 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset35 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor61.Append(luminanceModulation39);
            schemeColor61.Append(luminanceOffset35);

            fontReference23.Append(schemeColor61);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType8 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            seriesAxis1.Append(lineReference23);
            seriesAxis1.Append(fillReference23);
            seriesAxis1.Append(effectReference23);
            seriesAxis1.Append(fontReference23);
            seriesAxis1.Append(textCharacterPropertiesType8);

            Cs.SeriesLine seriesLine1 = new Cs.SeriesLine();
            Cs.LineReference lineReference24 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference24 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference24 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference24 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor62 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference24.Append(schemeColor62);

            Cs.ShapeProperties shapeProperties18 = new Cs.ShapeProperties();

            A.Outline outline27 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill34 = new A.SolidFill();

            A.SchemeColor schemeColor63 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation40 = new A.LuminanceModulation() { Val = 35000 };
            A.LuminanceOffset luminanceOffset36 = new A.LuminanceOffset() { Val = 65000 };

            schemeColor63.Append(luminanceModulation40);
            schemeColor63.Append(luminanceOffset36);

            solidFill34.Append(schemeColor63);
            A.Round round16 = new A.Round();

            outline27.Append(solidFill34);
            outline27.Append(round16);

            shapeProperties18.Append(outline27);

            seriesLine1.Append(lineReference24);
            seriesLine1.Append(fillReference24);
            seriesLine1.Append(effectReference24);
            seriesLine1.Append(fontReference24);
            seriesLine1.Append(shapeProperties18);

            Cs.TitleStyle titleStyle1 = new Cs.TitleStyle();
            Cs.LineReference lineReference25 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference25 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference25 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference25 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor64 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation41 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset37 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor64.Append(luminanceModulation41);
            schemeColor64.Append(luminanceOffset37);

            fontReference25.Append(schemeColor64);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType9 = new Cs.TextCharacterPropertiesType() { FontSize = 1400, Bold = false, Kerning = 1200, Spacing = 0, Baseline = 0 };

            titleStyle1.Append(lineReference25);
            titleStyle1.Append(fillReference25);
            titleStyle1.Append(effectReference25);
            titleStyle1.Append(fontReference25);
            titleStyle1.Append(textCharacterPropertiesType9);

            Cs.TrendlineStyle trendlineStyle1 = new Cs.TrendlineStyle();

            Cs.LineReference lineReference26 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.StyleColor styleColor7 = new Cs.StyleColor() { Val = "auto" };

            lineReference26.Append(styleColor7);
            Cs.FillReference fillReference26 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference26 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference26 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor65 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference26.Append(schemeColor65);

            Cs.ShapeProperties shapeProperties19 = new Cs.ShapeProperties();

            A.Outline outline28 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill35 = new A.SolidFill();
            A.SchemeColor schemeColor66 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill35.Append(schemeColor66);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.SystemDot };

            outline28.Append(solidFill35);
            outline28.Append(presetDash1);

            shapeProperties19.Append(outline28);

            trendlineStyle1.Append(lineReference26);
            trendlineStyle1.Append(fillReference26);
            trendlineStyle1.Append(effectReference26);
            trendlineStyle1.Append(fontReference26);
            trendlineStyle1.Append(shapeProperties19);

            Cs.TrendlineLabel trendlineLabel1 = new Cs.TrendlineLabel();
            Cs.LineReference lineReference27 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference27 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference27 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference27 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor67 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation42 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset38 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor67.Append(luminanceModulation42);
            schemeColor67.Append(luminanceOffset38);

            fontReference27.Append(schemeColor67);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType10 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            trendlineLabel1.Append(lineReference27);
            trendlineLabel1.Append(fillReference27);
            trendlineLabel1.Append(effectReference27);
            trendlineLabel1.Append(fontReference27);
            trendlineLabel1.Append(textCharacterPropertiesType10);

            Cs.UpBar upBar1 = new Cs.UpBar();
            Cs.LineReference lineReference28 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference28 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference28 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference28 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor68 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            fontReference28.Append(schemeColor68);

            Cs.ShapeProperties shapeProperties20 = new Cs.ShapeProperties();

            A.SolidFill solidFill36 = new A.SolidFill();
            A.SchemeColor schemeColor69 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            solidFill36.Append(schemeColor69);

            A.Outline outline29 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill37 = new A.SolidFill();

            A.SchemeColor schemeColor70 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation43 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset39 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor70.Append(luminanceModulation43);
            schemeColor70.Append(luminanceOffset39);

            solidFill37.Append(schemeColor70);

            outline29.Append(solidFill37);

            shapeProperties20.Append(solidFill36);
            shapeProperties20.Append(outline29);

            upBar1.Append(lineReference28);
            upBar1.Append(fillReference28);
            upBar1.Append(effectReference28);
            upBar1.Append(fontReference28);
            upBar1.Append(shapeProperties20);

            Cs.ValueAxis valueAxis2 = new Cs.ValueAxis();
            Cs.LineReference lineReference29 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference29 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference29 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference29 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };

            A.SchemeColor schemeColor71 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation44 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset40 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor71.Append(luminanceModulation44);
            schemeColor71.Append(luminanceOffset40);

            fontReference29.Append(schemeColor71);
            Cs.TextCharacterPropertiesType textCharacterPropertiesType11 = new Cs.TextCharacterPropertiesType() { FontSize = 900, Kerning = 1200 };

            valueAxis2.Append(lineReference29);
            valueAxis2.Append(fillReference29);
            valueAxis2.Append(effectReference29);
            valueAxis2.Append(fontReference29);
            valueAxis2.Append(textCharacterPropertiesType11);

            Cs.Wall wall1 = new Cs.Wall();
            Cs.LineReference lineReference30 = new Cs.LineReference() { Index = (UInt32Value)0U };
            Cs.FillReference fillReference30 = new Cs.FillReference() { Index = (UInt32Value)0U };
            Cs.EffectReference effectReference30 = new Cs.EffectReference() { Index = (UInt32Value)0U };

            Cs.FontReference fontReference30 = new Cs.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor72 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference30.Append(schemeColor72);

            Cs.ShapeProperties shapeProperties21 = new Cs.ShapeProperties();
            A.NoFill noFill17 = new A.NoFill();

            A.Outline outline30 = new A.Outline();
            A.NoFill noFill18 = new A.NoFill();

            outline30.Append(noFill18);

            shapeProperties21.Append(noFill17);
            shapeProperties21.Append(outline30);

            wall1.Append(lineReference30);
            wall1.Append(fillReference30);
            wall1.Append(effectReference30);
            wall1.Append(fontReference30);
            wall1.Append(shapeProperties21);

            chartStyle1.Append(axisTitle1);
            chartStyle1.Append(categoryAxis2);
            chartStyle1.Append(chartArea1);
            chartStyle1.Append(dataLabel1);
            chartStyle1.Append(dataLabelCallout1);
            chartStyle1.Append(dataPoint1);
            chartStyle1.Append(dataPoint3D1);
            chartStyle1.Append(dataPointLine1);
            chartStyle1.Append(dataPointMarker1);
            chartStyle1.Append(markerLayoutProperties1);
            chartStyle1.Append(dataPointWireframe1);
            chartStyle1.Append(dataTableStyle1);
            chartStyle1.Append(downBar1);
            chartStyle1.Append(dropLine1);
            chartStyle1.Append(errorBar1);
            chartStyle1.Append(floor1);
            chartStyle1.Append(gridlineMajor1);
            chartStyle1.Append(gridlineMinor1);
            chartStyle1.Append(hiLoLine1);
            chartStyle1.Append(leaderLine1);
            chartStyle1.Append(legendStyle1);
            chartStyle1.Append(plotArea2);
            chartStyle1.Append(plotArea3D1);
            chartStyle1.Append(seriesAxis1);
            chartStyle1.Append(seriesLine1);
            chartStyle1.Append(titleStyle1);
            chartStyle1.Append(trendlineStyle1);
            chartStyle1.Append(trendlineLabel1);
            chartStyle1.Append(upBar1);
            chartStyle1.Append(valueAxis2);
            chartStyle1.Append(wall1);

            chartStylePart1.ChartStyle = chartStyle1;
        }

        // Generates content of spreadsheetPrinterSettingsPart1.
        private void GenerateSpreadsheetPrinterSettingsPart1Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart1Data);
            spreadsheetPrinterSettingsPart1.FeedData(data);
            data.Close();
        }

        // Generates content of part.
        private void GeneratePartContent(ChartsheetPart part)
        {
            Chartsheet chartsheet1 = new Chartsheet();
            chartsheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            ChartSheetProperties chartSheetProperties1 = new ChartSheetProperties();
            TabColor tabColor1 = new TabColor() { Rgb = "FFFFFF00" };

            chartSheetProperties1.Append(tabColor1);

            ChartSheetViews chartSheetViews1 = new ChartSheetViews();
            ChartSheetView chartSheetView1 = new ChartSheetView() { WorkbookViewId = (UInt32Value)0U };

            chartSheetViews1.Append(chartSheetView1);
            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            ChartSheetPageSetup chartSheetPageSetup1 = new ChartSheetPageSetup() { Orientation = OrientationValues.Landscape, Id = "rId1" };
            Drawing drawing1 = new Drawing() { Id = "rId2" };

            chartsheet1.Append(chartSheetProperties1);
            chartsheet1.Append(chartSheetViews1);
            chartsheet1.Append(pageMargins1);
            chartsheet1.Append(chartSheetPageSetup1);
            chartsheet1.Append(drawing1);

            part.Chartsheet = chartsheet1;
        }

        #region Binary Data
        private string spreadsheetPrinterSettingsPart1Data = "QQBkAG8AYgBlACAAUABEAEYAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAwbcANQFU++AAQIAAQDqCm8IZAABAA8AWAICAAEAWAIDAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFBSSVbiIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAYAAAAAAAQJxAnECcAABAnAAAAAAAAAACwAKQDAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAwAAAAAAAAAwAhAAXEsDAGhDBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX3+euQUAAAAEAAAA/wD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALAAAABTTVRKAAAAABAAoABBAGQAbwBiAGUAIABQAEQARgAgAEMAbwBuAHYAZQByAHQAZQByAAAAUmVzb2x1dGlvbgAxMjAwZHBpAFBhZ2VTaXplAExldHRlcgBQYWdlUmVnaW9uAABMZWFkaW5nRWRnZQAASW5wdXRTbG90ACpVc2VGb3JtVHJheVRhYmxlAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAIAAEVCREEAAAEAAQAAAAEAAAABAAAAAQAAAAAAAABIAGkAZwBoACAAUQB1AGEAbABpAHQAeQAgAFAAcgBpAG4AdAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAA=";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }
}
