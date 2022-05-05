using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2ExcelVsto
{
    class Classbatim : ClassBase
    {
        public SLExcelData slExcelData;
        public ClassHeader header { get; set; }
        public TabooType tabooType { get; set; }
        public ClassNozar nozar { get; set; }
        public string PDFFileName;
        public BatimCommonProperty batimproperty;
        public List<TatHelka> tatHelkot = new List<TatHelka>();
        //        public ClassNozar classnozar ;
        public int endOfdata;
        public int presenttatHelka;

        public Classbatim(SLExcelData data, string pdfFileName)
        {
            header = new ClassHeader();
            slExcelData = data;
            nozar = new ClassNozar();
            PDFFileName = pdfFileName;
            presenttatHelka = -1;

        }
        public class ClassNozar
        {
            public string shtar;
            public string date;
            public string shtarType;
            public int line;

            public ClassNozar()
            {
                line = -1;
                shtar = "";
                date = "";
                shtarType = "";
            }
        }
        public class BatimCommonProperty
        {
            public int line;
            public int lineAddress;
            public int lineRemark;
            public int lineRemark1;
            public int linezikot;
            public string rashuiot;
            public string areasqmr;
            public string numOfTatHelkot;
            public string takanon;
            public string shtarYozer;
            public string tikbaitMeshutaf;
            public string addtress;
            public List<string> smallremarks;
            public string tikYozer;
            public List<string> remarks;
            public List<string> zikotHnah;

            public BatimCommonProperty()
            {
                line = -1;
                lineAddress = -1;
                lineRemark = -1;
                lineRemark1 = -1;
                linezikot = -1;
                rashuiot = "";
                areasqmr = "";
                numOfTatHelkot = "";
                takanon = "";
                shtarYozer = "";
                tikbaitMeshutaf = "";
                addtress = "";
                smallremarks = new List<string>();
                remarks = new List<string>();
                zikotHnah = new List<string>();

            }
        }
        public class Owner
        {
            public int line;
            public string transaction;
            public string name;
            public string idType;
            public string idNumber;
            public string part;
            public string shtar;
            public Owner()
            {
                line = -1;
                transaction = "";
                name = "";
                idType = "";
                idNumber = "";
                part = "";
                shtar = "";
            }
        };

        public class Attachment
        {
            public int line;
            public List<string> mark;
            public List<string> color;
            public List<string> description;
            public List<string> commonWith;
            public List<string> area;
            public Attachment()
            {
                line = -1;
                mark = new List<string>();
                color = new List<string>();
                description = new List<string>();
                commonWith = new List<string>();
                area = new List<string>();
            }
        };
        public class MortgageTatHelka
        {
            public int line;
            public string mtype;
            public string Name;
            public string idType;
            public string idNumber;
            public string part;
            public string shtar;
            public string grade;

            public MortgageTatHelka()
            {
                line = -1;
                mtype = "";
                Name = "";
                idType = "";
                idNumber = "";
                part = "";
                shtar = "";
                grade = "";
            }
        };
        public class TatHelka
        {
            public int line;
            public int lineOwners;
            public int lineRemarks;
            public int lineMortgage;
            public int lineLeasing;
            public int lineAttachments;
            public int lineZikotTatHelka;
            public int number;
            public string shetah;
            public string floor;
            public string entrance;
            public string agaf;
            public string building;
            public string partincommon;
            public List<Owner> owners;
            public List<Attachment> attachments;
            public List<MortgageTatHelka> mortgageTatHelkas;
            public List<Remark> remarks;
            public List<Leasing> leasings;
            public List<string> zikotTatHelkas;

            public TatHelka()
            {
                line = -1;
                number = -1;
                lineOwners = -1;
                lineRemarks = -1;
                lineMortgage = -1;
                lineLeasing = -1;
                lineAttachments = -1;
                lineZikotTatHelka = -1;
                shetah = "";
                floor = "";
                entrance = "";
                agaf = "";
                building = "";
                partincommon = "";
                owners = new List<Owner>();
                attachments = new List<Attachment>();
                mortgageTatHelkas = new List<MortgageTatHelka>();
                remarks = new List<Remark>();
                leasings = new List<Leasing>();
                zikotTatHelkas = new List<string>();
            }
        };

        public class Leasing
        {
            public int line;
            public List<string> leasingType;
            public List<string> Name;
            public List<string> idtype;
            public List<string> id;
            public List<string> part;
            public List<string> shtar;
            public List<string> rama;
            public List<string> endDate;
            public List<string> remarks;
            public List<string> partpropery;

            public Leasing()
            {
                line = -1;
                leasingType = new List<string>();
                Name = new List<string>();
                idtype = new List<string>();
                id = new List<string>();
                part = new List<string>();
                shtar = new List<string>();
                rama = new List<string>();
                endDate = new List<string>();
                remarks = new List<string>();
                partpropery = new List<string>();
            }
        }
        public class Remark
        {
            public int line;
            public string remarkType;
            public List<string> name;
            public List<string> idType;
            public List<string> idNumber;
            public List<string> part;
            public List<string> shtar;
            public List<string> remarklines;

            public Remark()
            {
                line = -1;
                remarkType = "";
                name = new List<string>();
                idType = new List<string>();
                idNumber = new List<string>();
                part = new List<string>();
                shtar = new List<string>();
                remarklines = new List<string>();
            }
        };

        public enum TabooType
        {
            Zehuiot,
            MeshutafAll,
            MeshutafTat
        }

        public class ClassHeader
        {
            public string dateCalendar { get; set; }
            public string dateHebrew { get; set; }
            public string time { get; set; }
            public String nesachNumber { get; set; }
            public List<string> tabooHeader = new List<string>();
            public String gush { get; set; }
            public String helka { get; set; }
            public string tatHelka { get; set; }
            public string headerFoot { get; set; }
        }


        public int buildHeader()
        {
            int retRow = 0;
            int i;
            List<string> dataRow = new List<string>();
            List<string> temp = new List<string>();


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
                    header.tatHelka = null;
                    for (int j = 0; j < dataRow.Count; j++)
                    {
                        if (ClassUtils.isAllDigit(dataRow[j]) && dataRow[j + 1] == "חלקה:")
                        {
                            header.helka = dataRow[j];
                            j++;
                            continue;
                        }
                        else if (ClassUtils.isAllDigit(dataRow[j]))
                        {
                            header.gush = dataRow[j];
                            break;
                        }
                    }

                    retRow = i + 1;
                    temp = (slExcelData.DataRows[retRow]);
                    if (ClassUtils.isArrayIncludeAllStringsParam(temp, "משותף", "עם", "חלקות"))
                    {
                        header.headerFoot = ClassUtils.buildCombinedline(temp);
                    }
                    else
                    {
                        header.headerFoot = ClassUtils.buildCombinedline(dataRow);
                    }

                    header.tabooHeader.Add(header.headerFoot);
                    break;
                }
                header.tabooHeader.Add(ClassUtils.buildCombinedline(dataRow));
            }
            slExcelData.DataRows = ClassUtils.RemoveHeaderSection(slExcelData.DataRows, header.headerFoot, header.dateCalendar);

            return retRow;
        }
        public void MapMainSections()
        {
            int next = 0;
            int row = next;
            List<List<string>> NesachTaboo = slExcelData.DataRows;
            while (!ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "הרכוש", "המשותף"))
            {
                if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "הנכס", "נוצר"))
                {
 //                   nozar = new ClassNozar();
                    nozar.line = row;
                    next = row;
                    break;

                }
                row++;
            }

            while (!ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "תת", "חלקה"))
            {
                if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "הרכוש", "המשותף"))
                {
                    batimproperty = new BatimCommonProperty();
                    batimproperty.line = row;
                    next = row;
                    break;
                }
                row++;
            }

            for (row = next + 1; row < NesachTaboo.Count; row++)
            {
                //while (!ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "תת", "חלקה"))
                while (!ClassUtils.isArrayIncludeAllStringsParamFromBeggining(NesachTaboo[row], "תת", "חלקה"))
                {
                    row++;
                    if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[row], "סוף", "נתונים"))
                    {
                        endOfdata = row;
                        return;
                    }
                    continue;
                }
                TatHelka tathelka = new TatHelka();
                tathelka.line = row;
                tatHelkot.Add(tathelka);
                continue;
            }

        }
        public void MapProperty()
        {
            List<List<string>> NesachTaboo = slExcelData.DataRows;
            List<string> temp ;
            for (int i = batimproperty.line; i < tatHelkot[0].line; i++)
            {
                temp = new List<string>(NesachTaboo[i]);
                temp = ClassUtils.reverseOrder(temp);
                if (temp[0] == "כתובת:")
                {
                    batimproperty.lineAddress = i;
                    continue;
                }
                if (temp[0] == "הערות:" && batimproperty.lineRemark1 == -1)
                {
                    batimproperty.lineRemark = i;
                    continue;
                }
                if (temp[0] == "הערות")
                {
                    batimproperty.lineRemark1 = i;
                    continue;
                }
                if (temp[0] == "זיקות" && temp[1] == "הנאה")
                {
                    batimproperty.linezikot = i;
                    continue;
                }
            }
        }
        public void MapSubSections()
        {
            List<List<string>> NesachTaboo = slExcelData.DataRows;
            for (int i = 0; i < tatHelkot.Count; i++)
            {
                TatHelka tat = tatHelkot[i];
                int lastline;
                if (i == tatHelkot.Count - 1)
                {
                    lastline = endOfdata;
                }
                else
                {
                    lastline = tatHelkot[i + 1].line;
                }
                for (int j = tatHelkot[i].line; j < lastline; j++)
                {

                    if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "הצמדות"))
                    {
                        tat.lineAttachments = j;
                        //Attachment att = new Attachment();
                        //tat.attachments.Add(att);
                        continue;
                    }
                    else if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "בעלויות"))
                    {
                        tat.lineOwners = j;
                        //Owner own = new Owner();
                        //tat.owners.Add(own);
                        continue;
                    }
                    else if (ClassUtils.isArrayIsUniqueInLine(NesachTaboo[j], "משכנתאות"))
                    {
                        tat.lineMortgage = j;
                        //MortgageTatHelka mort = new MortgageTatHelka();
                        //tat.mortgageTatHelkas.Add(mort);
                        continue;
                    }
                    else if (ClassUtils.isArrayIsUniqueInLine(NesachTaboo[j], "הערות"))
                    {
                        tat.lineRemarks = j;
                        //Remark rem = new Remark();
                        //tat.remarks.Add(rem);
                        continue;
                    }
                    else if (ClassUtils.isArrayIncludeAllStringsParam(NesachTaboo[j], "זיקות", "הנאה"))
                    {
                        tat.lineZikotTatHelka = j;
                        tat.zikotTatHelkas = new List<string>();
                        continue;
                    }
                    else if (ClassUtils.isArrayIsUniqueInLine(NesachTaboo[j], "חכירות"))
                    {
                        tat.lineLeasing = j;
                        //Leasing leas = new Leasing();
                        //tat.leasings.Add(leas);
                        continue;
                    }
                }
            }

        }

    }
}
