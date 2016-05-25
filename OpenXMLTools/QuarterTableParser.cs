using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLTools
{
    public class QuarterTableParser
    {
        public QuarterTableParser(IQuarterState quarterState)
        {
            _quarterState = quarterState;
        }

        public IEnumerable<QuarterTable> Parse(IEnumerable<MountainViewField> fields)
        {
            List<QuarterTable> quarterTables = new List<QuarterTable>();
            List<MountainViewField> quarterFields = new List<MountainViewField>();

            foreach (var field in fields)
            {
                if (_quarterState.VerifyQuarter(field))
                {
                    quarterFields.Add(field);
                }
                else
                {
                    ResetQuarterState();
                    var fieldDict = new FieldWeekDictionary(quarterFields);
                    quarterTables.Add(new QuarterTable(fieldDict));
                    quarterFields.Clear();
                    quarterFields.Add(field);
                }
            }

            if (quarterFields.Count() > 0)
            {
                ResetQuarterState();
                var fieldDict = new FieldWeekDictionary(quarterFields);
                quarterTables.Add(new QuarterTable(fieldDict));
                quarterFields.Clear();                
            }

            return quarterTables;
        }

        private void ResetQuarterState()
        {
            _quarterState = _quarterState.NextQuarter();
        }

        private IQuarterState _quarterState;
    }
}
