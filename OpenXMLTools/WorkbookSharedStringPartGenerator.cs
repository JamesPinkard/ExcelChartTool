using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;


namespace OpenXMLTools
{
    class WorkbookSharedStringPartGenerator
    {
        // Adds child parts and generates content of the specified part.
        public void CreateSharedStringTablePart(SharedStringTablePart part)
        {
            GeneratePartContent(part);

        }

        // Generates content of part.
        private void GeneratePartContent(SharedStringTablePart part)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)1550U, UniqueCount = (UInt32Value)32U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text1 = new Text();
            text1.Text = "RPW-6/7";

            sharedStringItem1.Append(text1);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text2 = new Text();
            text2.Text = "RPW-07";

            sharedStringItem2.Append(text2);

            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text3 = new Text();
            text3.Text = "RPW-06";

            sharedStringItem3.Append(text3);

            SharedStringItem sharedStringItem4 = new SharedStringItem();
            Text text4 = new Text();
            text4.Text = "RPW-03";

            sharedStringItem4.Append(text4);

            SharedStringItem sharedStringItem5 = new SharedStringItem();
            Text text5 = new Text();
            text5.Text = "IW-07";

            sharedStringItem5.Append(text5);

            SharedStringItem sharedStringItem6 = new SharedStringItem();
            Text text6 = new Text();
            text6.Text = "IW-06";

            sharedStringItem6.Append(text6);

            SharedStringItem sharedStringItem7 = new SharedStringItem();
            Text text7 = new Text();
            text7.Text = "IW-03";

            sharedStringItem7.Append(text7);

            SharedStringItem sharedStringItem8 = new SharedStringItem();
            Text text8 = new Text();
            text8.Text = "IW-02";

            sharedStringItem8.Append(text8);

            SharedStringItem sharedStringItem9 = new SharedStringItem();
            Text text9 = new Text();
            text9.Text = "EW-18";

            sharedStringItem9.Append(text9);

            SharedStringItem sharedStringItem10 = new SharedStringItem();
            Text text10 = new Text();
            text10.Text = "EW-17";

            sharedStringItem10.Append(text10);

            SharedStringItem sharedStringItem11 = new SharedStringItem();
            Text text11 = new Text();
            text11.Text = "EW-16";

            sharedStringItem11.Append(text11);

            SharedStringItem sharedStringItem12 = new SharedStringItem();
            Text text12 = new Text();
            text12.Text = "EW-12";

            sharedStringItem12.Append(text12);

            SharedStringItem sharedStringItem13 = new SharedStringItem();
            Text text13 = new Text();
            text13.Text = "EW-11";

            sharedStringItem13.Append(text13);

            SharedStringItem sharedStringItem14 = new SharedStringItem();
            Text text14 = new Text();
            text14.Text = "EW-08";

            sharedStringItem14.Append(text14);

            SharedStringItem sharedStringItem15 = new SharedStringItem();
            Text text15 = new Text();
            text15.Text = "EW-07";

            sharedStringItem15.Append(text15);

            SharedStringItem sharedStringItem16 = new SharedStringItem();
            Text text16 = new Text();
            text16.Text = "EW-06";

            sharedStringItem16.Append(text16);

            SharedStringItem sharedStringItem17 = new SharedStringItem();
            Text text17 = new Text();
            text17.Text = "corrected transcription error per O&M form";

            sharedStringItem17.Append(text17);

            SharedStringItem sharedStringItem18 = new SharedStringItem();
            Text text18 = new Text();
            text18.Text = "counter rollover at 999,999";

            sharedStringItem18.Append(text18);

            SharedStringItem sharedStringItem19 = new SharedStringItem();
            Text text19 = new Text();
            text19.Text = "why red?";

            sharedStringItem19.Append(text19);

            SharedStringItem sharedStringItem20 = new SharedStringItem();
            Text text20 = new Text();
            text20.Text = "EW-13";

            sharedStringItem20.Append(text20);

            SharedStringItem sharedStringItem21 = new SharedStringItem();
            Text text21 = new Text();
            text21.Text = "corrected time per O&M form";

            sharedStringItem21.Append(text21);

            SharedStringItem sharedStringItem22 = new SharedStringItem();
            Text text22 = new Text();
            text22.Text = "\"time\" value not on O&M form, source unknown";

            sharedStringItem22.Append(text22);

            SharedStringItem sharedStringItem23 = new SharedStringItem();
            Text text23 = new Text();
            text23.Text = "time not recorded on form";

            sharedStringItem23.Append(text23);

            SharedStringItem sharedStringItem24 = new SharedStringItem();
            Text text24 = new Text();
            text24.Text = "duplicate entry, source of \"time\" value unknown";

            sharedStringItem24.Append(text24);

            SharedStringItem sharedStringItem25 = new SharedStringItem();
            Text text25 = new Text();
            text25.Text = "system down, no field reading";

            sharedStringItem25.Append(text25);

            SharedStringItem sharedStringItem26 = new SharedStringItem();
            Text text26 = new Text();
            text26.Text = "2nd reading on 6/19/15, no change in field reading";

            sharedStringItem26.Append(text26);

            SharedStringItem sharedStringItem27 = new SharedStringItem();
            Text text27 = new Text();
            text27.Text = "pink highlighting and/or red text indicate questions/issues that need to be addressed.";

            sharedStringItem27.Append(text27);

            SharedStringItem sharedStringItem28 = new SharedStringItem();
            Text text28 = new Text();
            text28.Text = "Comments";

            sharedStringItem28.Append(text28);

            SharedStringItem sharedStringItem29 = new SharedStringItem();
            Text text29 = new Text();
            text29.Text = "Field Reading";

            sharedStringItem29.Append(text29);

            SharedStringItem sharedStringItem30 = new SharedStringItem();
            Text text30 = new Text();
            text30.Text = "Time";

            sharedStringItem30.Append(text30);

            SharedStringItem sharedStringItem31 = new SharedStringItem();
            Text text31 = new Text();
            text31.Text = "Date";

            sharedStringItem31.Append(text31);

            SharedStringItem sharedStringItem32 = new SharedStringItem();
            Text text32 = new Text();
            text32.Text = "Station";

            sharedStringItem32.Append(text32);

            sharedStringTable1.Append(sharedStringItem1);
            sharedStringTable1.Append(sharedStringItem2);
            sharedStringTable1.Append(sharedStringItem3);
            sharedStringTable1.Append(sharedStringItem4);
            sharedStringTable1.Append(sharedStringItem5);
            sharedStringTable1.Append(sharedStringItem6);
            sharedStringTable1.Append(sharedStringItem7);
            sharedStringTable1.Append(sharedStringItem8);
            sharedStringTable1.Append(sharedStringItem9);
            sharedStringTable1.Append(sharedStringItem10);
            sharedStringTable1.Append(sharedStringItem11);
            sharedStringTable1.Append(sharedStringItem12);
            sharedStringTable1.Append(sharedStringItem13);
            sharedStringTable1.Append(sharedStringItem14);
            sharedStringTable1.Append(sharedStringItem15);
            sharedStringTable1.Append(sharedStringItem16);
            sharedStringTable1.Append(sharedStringItem17);
            sharedStringTable1.Append(sharedStringItem18);
            sharedStringTable1.Append(sharedStringItem19);
            sharedStringTable1.Append(sharedStringItem20);
            sharedStringTable1.Append(sharedStringItem21);
            sharedStringTable1.Append(sharedStringItem22);
            sharedStringTable1.Append(sharedStringItem23);
            sharedStringTable1.Append(sharedStringItem24);
            sharedStringTable1.Append(sharedStringItem25);
            sharedStringTable1.Append(sharedStringItem26);
            sharedStringTable1.Append(sharedStringItem27);
            sharedStringTable1.Append(sharedStringItem28);
            sharedStringTable1.Append(sharedStringItem29);
            sharedStringTable1.Append(sharedStringItem30);
            sharedStringTable1.Append(sharedStringItem31);
            sharedStringTable1.Append(sharedStringItem32);

            part.SharedStringTable = sharedStringTable1;
        }
    }
}
