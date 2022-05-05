using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using static PDF2ExcelVsto.ClassBatim;


namespace PDF2ExcelVsto
{
    class ClassbatimManager
    {
        public ClassFilesHandle filesHandle;
        public List<ClassBatim> allBatim;
        List<string> batimCSVfile;
        ClassExcelOperations excelOperations;

        public ClassbatimManager(ClassFilesHandle fhd, ClassExcelOperations excel)
        {
            filesHandle = fhd;
            excelOperations = excel;
        }

        public void convertBatimtoExcel()
        {
            allBatim = new List<ClassBatim>();
            batimCSVfile = filesHandle.getCSVFiles("batim");
            foreach(string csvfile in batimCSVfile)
            {
                int currentRow;
                SLExcelData slExcelData = new SLExcelData();
                slExcelData.DataRows = ClassUtils.File2Data(csvfile);
                string fn = Path.GetFileName(csvfile);
                fn = fn.Replace("csv", "pdf");
                ClassBatim batim = new ClassBatim(slExcelData, fn);
                currentRow = batim.buildHeader();
                ClassMapBatim TaboMap = new ClassMapBatim(batim.slExcelData.DataRows);

            }
        }

    }

}
