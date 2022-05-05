using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2ExcelVsto
{
    class ClassBatim
    {
        public SLExcelData slExcelData;
        public ClassHeader header{ get; set; }
        public TabooType tabooType{ get; set; }
        public Nozar nozar{ get; set; }
        public string PDFFileName;

        public ClassBatim(SLExcelData data, string pdfFileName)
        {
            header = new ClassHeader();
            slExcelData = data;
            nozar = null;
            PDFFileName = pdfFileName;
        }

        public class Nozar
        {
            public List<string> nozar{ get; set; }
        }

        public class ClassHeader
        {
            public string dateCalendar{ get; set; }
            public string dateHebrew{ get; set; }
            public string time{ get; set; }
            public String nesachNumber{ get; set; }
            public List<string> tabooHeader = new List<string>();
            public String gush{ get; set; }
            public String helka{ get; set; }
            public string tatHelka{ get; set; }
            public string headerFoot{ get; set; }
        }
        public enum TabooType
        {
            Zehuiot,
            MeshutafAll,
            MeshutafTat
        }
        public int buildHeader()
        {
            int retRow = 0;
            int i;
            List<string> dataRow = new List<string>();


            header.dateCalendar = ClassUtils.buildCombinedline(slExcelData.DataRows[0]);
            header.dateHebrew = ClassUtils.buildCombinedline(slExcelData.DataRows[1]);
            header.time = ClassUtils.buildCombinedline(slExcelData.DataRows[2]);
            header.nesachNumber = ClassUtils.buildCombinedline(slExcelData.DataRows[3]);

            for (i = 4; i < slExcelData.DataRows.Count; i++)
            {
                dataRow = slExcelData.DataRows[i];
                if (ClassUtils.isArrayIncludString(dataRow, "הזכויות") > -1)
                {
                    tabooType = TabooType.Zehuiot;
                }
                else if (ClassUtils.isArrayIncludString(dataRow, "משותפים") > -1)
                {
                    tabooType = TabooType.MeshutafAll;
                }
                else if (ClassUtils.isArrayIncludString(dataRow, "גוש") > -1)
                {

                    if (dataRow.Count == 4) // Tat Chelka 
                    {
                        header.gush = dataRow[2];
                        header.tatHelka = null;
                        header.helka = dataRow[0];
                    }
                    else
                    {
                        //string[] helk = words[1].Split(' ');
                        //header.gush = helk[1];
                        //header.helka = words[2];
                        //header.tatHelka = "";
                    }
                    retRow = i + 1;
                    header.headerFoot = ClassUtils.buildCombinedline(dataRow);
                    header.tabooHeader.Add(header.headerFoot);
                    break;
                }
                header.tabooHeader.Add(ClassUtils.buildCombinedline(dataRow));
            }
            slExcelData.DataRows = ClassUtils.RemoveHeaderSection(slExcelData.DataRows, header.headerFoot, header.dateCalendar);

            return retRow;
        }

    }
}
