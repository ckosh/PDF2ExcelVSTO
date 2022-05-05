using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2ExcelVsto
{
    class SLExcelData
    {
        public List<string> Headers { get; set; }
        public List<List<string>> DataRows { get; set; }
        public string SheetName { get; set; }

        public SLExcelData()
        {
            Headers = new List<string>();
            DataRows = new List<List<string>>();
        }

    }
}
