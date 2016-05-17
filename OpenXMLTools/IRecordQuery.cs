using System.Collections.Generic;

namespace OpenXMLTools
{
    public interface IRecordQuery
    {
        IEnumerable<IRecord> Query(IEnumerable<MountainViewField> fields);
    }
}