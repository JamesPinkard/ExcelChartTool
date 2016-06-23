using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace OpenXMLTools
{
    public class ExtractionWellFieldModifier
    {
        public ExtractionWellFieldModifier()
        {

        }

        public  IEnumerable<MountainViewField> Modify(IEnumerable<MountainViewField> fields)
        {
            var extractionFields = fields.Where(f => f.StationName.StartsWith("EW"));
            var orderedFields = extractionFields.OrderBy(f => f.MeasureTime);
            var fieldGroups = orderedFields.GroupBy(f => f.StationName);

            var logger = File.CreateText(@".\ew_error_log.txt");

            foreach (var group in fieldGroups)
            {
                var previousField = group.First();
                foreach (var thisField in group.Skip(1))
                {
                    if (thisField.TotalizerReading < previousField.TotalizerReading)
                    {
                        logger.WriteLine(@"lower reading than previous-{0} vs {1}", previousField.ToString(), thisField.ToString());
                        var laterFields = group.Where(e => e.MeasureTime >= thisField.MeasureTime);
                        foreach (var lf in laterFields)
                        {
                            lf.TotalizerReading += 1000000;
                        }
                    }
                    previousField = thisField;
                }
            }

            foreach (var ew in extractionFields)
            {
                ew.TotalizerReading = Convert.ToInt32(ew.TotalizerReading * 0.7);
            }
            logger.Close();
            return fields;
        }


    }
}
