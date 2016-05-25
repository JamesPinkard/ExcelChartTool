using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class StationTableParser
    {
        public IEnumerable<StationTable> CompileStationTables(IEnumerable<MountainViewField> fields)
        {
            List<StationTable> stationMeasurements = new List<StationTable>();

            IEnumerable<IGrouping<string, MountainViewField>> query = fields.GroupBy(field => field.StationName);

            foreach (IGrouping<string, MountainViewField> wellFields in query)
            {
                StationTable stationTable = new StationTable(wellFields.Key);

                foreach (var field in wellFields)
                {
                    stationTable.AddField(field);
                }

                stationMeasurements.Add(stationTable);
            }

            return stationMeasurements;
        }
    }
}
