using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class ReadingErrorLogger
    {
        public void Log(IEnumerable<MountainViewField> fields)
        {            
            var orderedFields = fields.OrderBy(f => f.MeasureTime);
            var fieldGroups = orderedFields.GroupBy(f => f.StationName);
            var logger = File.CreateText(@".\error_log.txt");

            foreach (var group in fieldGroups)
            {
                var previousField = group.First();                
                foreach (var thisField in group.Skip(1))
                {
                    if (thisField.TotalizerReading < previousField.TotalizerReading)
                    {
                        logger.WriteLine(@"lower reading than previous-{0} vs {1}", previousField.ToString(), thisField.ToString());                        
                    }
                     
                    previousField = thisField;
                }
            }
            logger.Close();
            
        }
    }
}
