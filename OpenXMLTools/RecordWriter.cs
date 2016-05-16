using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace OpenXMLTools
{
    public class RecordWriter
    {
        public RecordWriter(string fileName)
        {
            _fileName = fileName;
        }

        public void Write(IEnumerable<IRecord> records)
        {            
            using (StreamWriter writer = new StreamWriter(_fileName))
            {
                foreach (var record in records)
                {
                    writer.WriteLine(record.ToString());
                }
            }
        }

        private string _fileName;
    }
}
