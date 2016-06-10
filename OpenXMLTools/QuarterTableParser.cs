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
            _originalState = quarterState;
        }

        public IEnumerable<QuarterTable> Parse(IEnumerable<MountainViewField> fields)
        {
            List<QuarterTable> quarterTables = new List<QuarterTable>();
            List<MountainViewField> quarterFields = new List<MountainViewField>();
            _previousField = fields.First();

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
                    quarterTables.Add(new QuarterTable(fieldDict, _previousField));
                    _previousField = quarterFields.Last();
                    quarterFields = new List<MountainViewField>();
                    quarterFields.Add(field);
                }
            }

            if (quarterFields.Count() > 0)
            {
                ResetQuarterState();
                var fieldDict = new FieldWeekDictionary(quarterFields);
                quarterTables.Add(new QuarterTable(fieldDict, _previousField));
                quarterFields = new List<MountainViewField>();
            }

            ResetToOriginalState();
            return quarterTables;
        }

        private void ResetQuarterState()
        {
            _quarterState = _quarterState.NextQuarter();
        }

        private void ResetToOriginalState()
        {
            _quarterState = _originalState;
        }

        private IQuarterState _quarterState;
        private IQuarterState _originalState;
        private MountainViewField _previousField;
    }
}
