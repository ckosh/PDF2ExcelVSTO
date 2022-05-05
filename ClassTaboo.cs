using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2ExcelVsto
{
    class ClassTaboo : ClassBase
    {
        public SLExcelData slExcelData;
        public ClassHeader header { get; set; }
        public TabooType tabooType { get; set; }
        public Nozar nozar { get; set; }
        public List<TatHelka> tatHelkot;
        public List<ZhuiotOwner> zhuiotOwners;
        public List<Leasing> leasings;
        public List<Mortgage> mortgages;
        public List<Remarks> remarks;
        public string PDFFileName;

        public CommonProperty commonProperty { get; set; }
        public DescriptionProperty description { get; set; }

        public ClassTaboo(SLExcelData data, string pdfFileName)
        {
            header = new ClassHeader();
            slExcelData = data;
            nozar = null;
            tatHelkot = new List<TatHelka>();
            PDFFileName = pdfFileName;
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
        public enum TabooType
        {
            Zehuiot,
            MeshutafAll,
            MeshutafTat
        }
        public class Nozar
        {
            public List<string> nozar { get; set; }
        }
        public class TatHelka
        {
            public int Number { get; set; }
            public string area { get; set; }
            public string floor { get; set; }
            public string entrance { get; set; }
            public string house { get; set; }
            public string fraction { get; set; }
            public List<TatOwner> owners { get; set; }
        }
        public class ZhuiotOwner
        {
            public string shtarNum { get; set; }
            public string date { get; set; }
            public string transactionType { get; set; }
            public string ownerName { get; set; }
            public string idType { get; set; }
            public string idNumber { get; set; }
            public string ownerPart { get; set; }
        }
        public class Leasing
        {
            public string LeaserLevel { get; set; }
            public string OriginlShtar { get; set; }
            public string PropertyPart { get; set; }
            public string EndDate { get; set; }
            public List<string> remarks = new List<string>();
            public List<LeasingOwner> leasingOwners = new List<LeasingOwner>();
        }
        public class CommonProperty
        {
            public string municipality { get; set; }
            public string area { get; set; }
            public string subHelkot { get; set; }
            public string takanon { get; set; }
            public string shtarYozer { get; set; }
            public string tikBaitMeshutaf { get; set; }
            public string tikYozer { get; set; }
            public string address { get; set; }
        }
        public class DescriptionProperty
        {
            public string rashuiot { get; set; }
            public string rashuiot1 { get; set; }
            public string rashuiot2 { get; set; }
            public string area { get; set; }
            public string landType { get; set; }
            public string oldNumbers { get; set; }
            public string remarks { get; set; }
            public string connections { get; set; }
            public string designation { get; set; }

        }
        public class TatOwner
        {
            public string transaction { get; set; }
            public string name { get; set; }
            public string idType { get; set; }
            public string idNumber { get; set; }
            public string ownerFraction { get; set; }
            public string shtar { get; set; }
        }
        public class LeasingOwner
        {
            public string shtarNum { get; set; }
            public string date { get; set; }
            public string transactionType { get; set; }
            public string LeaserName { get; set; }
            public string idType { get; set; }
            public string idNumber { get; set; }
            public string LeaserPart { get; set; }
            public string Remarks { get; set; }
        }
        public class Mortgage
        {
            public string shtarNum { get; set; }            
            public string date { get; set; }
            public string MortgageType { get; set; }
           
            public MortgageOwner mortgageOwner = new MortgageOwner();
            public MortgageBorower mortgageBorower = new MortgageBorower();
            public string grade { get; set; }
            public string amount { get; set; }
            public string OriginalShtarNum { get; set; }
            public string propPart { get; set; }
            public List<string> remarks = new List<string>();
        }
        public class MortgageOwner
        {
            public List<string> ownerName = new List<string>();
            public List<string> ownerIDType = new List<string>();
            public List<string> ownerIDNumber = new List<string>();
        }
        public class MortgageBorower
        {
            public List<string> borowerName = new List<string>();
            public List<string> borowerIDType = new List<string>();
            public List<string> borowerIDNumber = new List<string>();
        }
        public class Remarks
        {
            public string shtarNum { get; set; }
            public string date { get; set; }
            public string actionType { get; set; }
            public List<string> onOwner = new List<string>();
            public List<string> idType = new List<string>();
            public List<string> idNumber = new List<string>();
            public List<string> remarks = new List<string>();
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
        public void buildNozarCSV(List<int> sec)
        {
            for (int i = sec[0]; i < sec[1]; i++)
            {
                List<string> lst = slExcelData.DataRows[i];
                if (ClassUtils.isArrayIncludString(lst, "נוצר") > -1)
                {
                    List<string> subList = new List<string> { "הנכס", "נוצר" };
                    if (ClassUtils.isArrayIncludeAllStrings(lst, subList))
                    {
                        nozar = new Nozar();
                        nozar.nozar = ClassUtils.reverseOrder(slExcelData.DataRows[i]);
                    }
                }
            }
        }
        public void buildDesciptionPropertyCSV(List<int> sec, List<int> rows)
        {
            description = new DescriptionProperty();
            List<string> l1 = new List<string>();
            List<string> l2 = new List<string>();
            int k = 0;
            int j;
            l1 = ClasszhuiotUtils.rawlineToKeyPropDescription(slExcelData.DataRows[sec[0]]);
            l2 = ClassUtils.reverseOrder(slExcelData.DataRows[sec[0] + 1]);
            for (j = 0; j < l1.Count; j++)
            {
                if (l1[j] == "רשויות")
                {
                    while (!ClassUtils.isFloatingNumber(l2[j + k]))
                    {
                        description.rashuiot = description.rashuiot + " " + l2[j + k];
                        k++;
                    }
                    continue;
                }
                else if (l1[j] == "שטח במ\"ר")
                {
                    if (ClassUtils.isFloatingNumber(l2[k]))
                    {
                        description.area = l2[k];
                        k++;
                    }
                }
                else if (l1[j] == "סוג המקרקעין")
                {
                    description.landType = l2[k];
                    k++;
                    if (k < l2.Count)
                    {
                        if (l2[k] == "יעוד")
                        {
                            description.landType = description.landType + " " + l2[k];
                            k++;
                        }
                    }
                }
                else if (l1[j] == "יעוד המקרקעין")
                {
                    description.designation = l2[k];
                    k++;
                }
            }
            if ( rows.Count > 1)
            {
                if (rows[0] + 2 < rows[1])
                {
                    l1 = slExcelData.DataRows[rows[0] + 2];
                    description.rashuiot1 = ClassUtils.buildCombinedline(l1);
                    j++;
                }
                if (rows[0] + 3 < rows[1])
                {
                    l1 = slExcelData.DataRows[rows[0] + 3];
                    description.rashuiot1 = ClassUtils.buildCombinedline(l1);
                    j++;
                }
                do
                {
                    string sss = ClassUtils.buildCombinedline(slExcelData.DataRows[j]);
                    j++;
                    if (sss == "המספרים הישנים של החלקה")
                    {
                        description.oldNumbers = ClassUtils.buildCombinedline(slExcelData.DataRows[j]);
                        j++;
                    }
                    else if (sss == "מספר המחוברים")
                    {
                        description.connections = ClassUtils.buildCombinedline(slExcelData.DataRows[j]);
                        j++;
                    }
                    else if (sss == "הערות רשם המקרקעין")
                    {
                        description.remarks = ClassUtils.buildCombinedline(slExcelData.DataRows[j]);
                        j++;
                    }

                } while (j < sec[1]);
            }
        }
        public void buildOwnersZhuiotCSV(List<int> sec, List<int> OwnersRows)
        {
            zhuiotOwners = new List<ZhuiotOwner>();
            List<string> l1 = new List<string>();
            List<string> l2 = new List<string>();
          
            for (int i = 0; i < OwnersRows.Count; i++)
            {
                string cont0 = "";
                if (i == 63)
                {
                    i = 63;
                }
                ZhuiotOwner zhuiotOwner = new ZhuiotOwner();
                l1 = ClasszhuiotUtils.rawlineToKeyPropOwners(slExcelData.DataRows[OwnersRows[i]]);
                l2 = ClasszhuiotUtils.rawlineToValuesPropOwners0(l1, slExcelData.DataRows[OwnersRows[i] + 1], ref cont0);
//                l2 = ClassUtils.rawlineToValuesLeasing0(l1, slExcelData.DataRows[OwnersRows[i] + 1]);
//                l2.Reverse();
                //                l2 = ClassUtils.rawlineToValuesPropOwners(slExcelData.DataRows[OwnersRows[i] + 1]);
                if (l1.Count == l2.Count)
                {
                    for (int k = 0; k < l1.Count; k++)
                    {
                        if (l1[k] == "מס' שטר")
                        {
                            zhuiotOwner.shtarNum = l2[k];
                        }
                        else if (l1[k] == "תאריך")
                        {
                            zhuiotOwner.date = l2[k];
                        }
                        else if (l1[k] == "מהות פעולה")
                        {
                            zhuiotOwner.transactionType = l2[k] + " " + cont0;
                        }
                        else if (l1[k] == "הבעלים")
                        {
                            zhuiotOwner.ownerName = l2[k];
                        }
                        else if (l1[k] == "סוג זיהוי")
                        {
                            zhuiotOwner.idType = l2[k];
                        }
                        else if (l1[k] == "מס' זיהוי")
                        {
                            bool foreign = ClassUtils.isForeignID(l2[k]);
                            if (foreign)
                            {
                                l2[k] = ClassUtils.Reverse(l2[k]);
                            }
                            zhuiotOwner.idNumber = l2[k];
                        }
                    }
                    if (!ClassUtils.isArrayIncludeAllStringsParam(slExcelData.DataRows[OwnersRows[i] + 2], "החלק", "בנכס") && cont0 == "")
                    {
                        List<string> sss = new List<string>(slExcelData.DataRows[OwnersRows[i] + 2]);
                        if (sss[sss.Count - 1] == "לשכת" && sss[sss.Count - 2] == "פרצלציה") // remove from line
                        {
                            sss.RemoveAt(sss.Count - 1);
                            sss.RemoveAt(sss.Count - 1);
                        }
                        zhuiotOwner.ownerName =   zhuiotOwner.ownerName +" " + ClassUtils.buildReverseCombinedLine(sss) ;
                        zhuiotOwner.ownerPart = ClassUtils.buildReverseCombinedLine(slExcelData.DataRows[OwnersRows[i] + 4]);
                    }
                    else if ( cont0 != "")
                    {
                        zhuiotOwner.ownerPart = ClassUtils.buildReverseCombinedLine(slExcelData.DataRows[OwnersRows[i] + 4]);
                    }
                    else
                    {
                        zhuiotOwner.ownerPart = ClassUtils.buildReverseCombinedLine(slExcelData.DataRows[OwnersRows[i] + 3]);
                        string fff = ClassUtils.buildReverseCombinedLine(slExcelData.DataRows[OwnersRows[i] + 4]);
                        if (fff.All(Char.IsDigit))
                        {
                            zhuiotOwner.ownerPart = zhuiotOwner.ownerPart + " " + fff;
                        }
                     }
                }
                zhuiotOwners.Add(zhuiotOwner);
            }
        }
        public void buildRemarks(List<int> sec, List<int> RemarksRows)
        {
            remarks = new List<Remarks>();
            List<string> l1 = new List<string>();
            List<string> l2 = new List<string>();
            int rowNumberj;
            for (int i = 0; i < RemarksRows.Count-1; i++)
            {
                if ( i == 48)
                {
                  
                }
                string cont0 = "";
                string cont1 = "";
                bool checkList = false;
                rowNumberj = RemarksRows[i];
                Remarks remark = new Remarks();
                l1 = ClasszhuiotUtils.rawlineToKeyPropOwners(slExcelData.DataRows[RemarksRows[i]]);
                l2 = ClasszhuiotUtils.rawlineToValuesRemarks0(l1, slExcelData.DataRows[rowNumberj + 1], ref cont0, ref cont1, ref checkList );
//                l2.Reverse();
                if (l1.Count == l2.Count)
                {
                    for (int k = 0; k < l1.Count; k++)
                    {
                        if (l1[k] == "מס' שטר")
                        {
                            remark.shtarNum = l2[k];
                        }
                        else if (l1[k] == "תאריך")
                        {
                            remark.date = l2[k];
                        }
                        else if (l1[k] == "מהות פעולה")
                        {
                            remark.actionType = l2[k] + " " + cont0 + " " + cont1;
                        }
                        else if (l1[k] == "שם המוטב")
                        {
                            remark.onOwner.Add(l2[k]);
                        }
                        else if (l1[k] == "סוג זיהוי")
                        {
                            remark.idType.Add(l2[k]);
                        }
                        else if (l1[k] == "מס' זיהוי")
                        {
                            bool foreign = ClassUtils.isForeignID(l2[k]);
                            if (foreign)
                            {
                                l2[k] = ClassUtils.Reverse(l2[k]);
                            }
                            remark.idNumber.Add(l2[k]);
                        }
                    }
                }
                rowNumberj = rowNumberj + 2;
                if (cont0 == "126") rowNumberj++;
                if (checkList)
                {
                    string textToIgnore = "";
                    while (ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "הערות:") == -1 &&
                           ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "זיהוי") == -1 &&
                           ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "הבעלים") == -1 &&
                           ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "סכום") == -1 &&
                           ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "הבעלות") == -1 &&
                           ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "החלק") == -1 &&
                           ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "פעולה") == -1)
                    {
                        string idNum = " ";
                        string idType = " ";
                        string Name = " ";
                        int nn = ClasszhuiotUtils.parse126Worning(l1,slExcelData.DataRows[rowNumberj], textToIgnore, ref idNum, ref idType, ref Name);
                        if ( nn < l1.Count - 3)
                        {
                            if ( Name != " ")
                            {
                                remark.onOwner[remark.onOwner.Count - 1] = remark.onOwner[remark.onOwner.Count - 1] + " " + Name;
                            }
                        }
                        else
                        {
                            if (idNum != "")
                            {
                                remark.idNumber.Add(idNum);
                            }
                            if (idType != "")
                            {
                                remark.idType.Add(idType);
                            }
                            if (Name != "")
                            {
                                remark.onOwner.Add(Name);
                            }
                        }
                        rowNumberj++;
                    }
                }
                else if (cont0 != "")
                {
                    string[] subs = cont0.Split(' ');
                    List<string> ssss = new List<string>(slExcelData.DataRows[rowNumberj]);
                    if (ssss[ssss.Count - 1] == "לשכת" && ssss[ssss.Count - 2] == "פרצלציה")
                    {
                        ssss.RemoveAt(ssss.Count - 1);
                        ssss.RemoveAt(ssss.Count - 1);
                    }
                    string sss  = ClassUtils.buildCombinedLineSelected(ssss, 0, ssss.Count-subs.Length, false);
                    if ( sss != "")
                    {
                        sss = ClassUtils.ReverseWordsInString(sss);
                        remark.onOwner[0] = remark.onOwner[0] + sss;
                    }
                }
                int offset = rowNumberj ;
                if (cont0 != "") offset++;
                if (cont1 != "") offset++;
                for (int j = offset; j < RemarksRows[i+1]; j++)
                {
                    string sss = ClassUtils.buildCombinedline(slExcelData.DataRows[j]);
                    remark.remarks.Add(sss);
                }
                remarks.Add(remark);
            }

        }
        public void buildMortGage(List<int>sec , List<int> MortgageRows)
        {
            mortgages = new List<Mortgage>();
            List<string> l1 = new List<string>();
            List<string> l2 = new List<string>();
            int rowNumberj;
           
            for (int i = 0; i < MortgageRows.Count-1 ; i++)
            {
                rowNumberj = MortgageRows[i];
                Mortgage mortgage = new Mortgage();

                string cont0 = "";
                string cont1 = "";
                bool checkList = false;
                l1 = ClasszhuiotUtils.rawlineToKeyPropOwners(slExcelData.DataRows[MortgageRows[i]]);
                l2 = ClasszhuiotUtils.rawlineToValuesMortgage(l1, slExcelData.DataRows[rowNumberj + 1], ref cont0, ref cont1, ref checkList);
                if (l1.Count == l2.Count)
                {
                    for (int k = 0; k < l1.Count; k++)
                    {
                        if (l1[k] == "מס' שטר")
                        {
                            mortgage.shtarNum = l2[k];
                        }
                        else if (l1[k] == "תאריך")
                        {
                            mortgage.date = l2[k];
                        }
                        else if (l1[k] == "מהות פעולה")
                        {
                            mortgage.MortgageType = l2[k] + " " + cont0 + " " + cont1;
                        }
                        else if (l1[k] == "בעלי המשכנתה")
                        {
                            mortgage.mortgageOwner.ownerName.Add(l2[k]);
                        }
                        else if (l1[k] == "סוג זיהוי")
                        {
                            mortgage.mortgageOwner.ownerIDType.Add(l2[k]);
                        }
                        else if (l1[k] == "מס' זיהוי")
                        {
                            bool foreign = ClassUtils.isForeignID(l2[k]);
                            if (foreign)
                            {
                                l2[k] = ClassUtils.Reverse(l2[k]);
                            }
                            mortgage.mortgageOwner.ownerIDNumber.Add(l2[k]);
                        }
                    }
                }
                rowNumberj = rowNumberj + 2;
                if (checkList)
                {
                    string textToIgnore = "";
                    while (ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "הלווה") == -1 &&
                           ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "דרגה") == -1)
                    {
                        string idNum = " ";
                        string idType = " ";
                        string Name = " ";
                        List<string> temp = new List<string>(slExcelData.DataRows[rowNumberj]);
                        if (ClassUtils.isArrayIncludString(temp, "פרצלציה") > -1)
                        {
                            temp.RemoveAt(temp.Count - 1);
                            temp.RemoveAt(temp.Count - 1);
                        }
                        int nn = ClasszhuiotUtils.parseMortgageOwnerCont(l1, temp, textToIgnore, ref idNum, ref idType, ref Name);
                        if (nn < l1.Count - 2)
                        {
                            if (Name != " " && idNum == " " & idType == " " )
                            {
                                int size = mortgage.mortgageOwner.ownerName.Count -1;
                                mortgage.mortgageOwner.ownerName[size] = mortgage.mortgageOwner.ownerName[size] + " " + Name;
                            }
                            else
                            {
                                if (idNum != "")
                                {
                                    mortgage.mortgageOwner.ownerIDNumber.Add(idNum);
                                }
                                if (idType != "")
                                {
                                    mortgage.mortgageOwner.ownerIDType.Add(idType);
                                }
                                if (Name != "")
                                {
                                    mortgage.mortgageOwner.ownerName.Add(Name);
                                }
                            }
                        }
                        else
                        {
                            if (idNum != "")
                            {
                                mortgage.mortgageOwner.ownerIDNumber.Add(idNum);
                            }
                            if (idType != "")
                            {
                                mortgage.mortgageOwner.ownerIDType.Add(idType);
                            }
                            if (Name != "")
                            {
                                mortgage.mortgageOwner.ownerName.Add(Name);
                            }
                        }
                        rowNumberj++;
                    }
                }
               
                if (ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "הלווה") > -1)
                {
                    l1 = ClasszhuiotUtils.rawlineToKeyPropborrow(slExcelData.DataRows[rowNumberj]);
                    rowNumberj++;
                    do
                    {
                        l2 = ClasszhuiotUtils.rawlineToValuesMortgageboroer(l1, slExcelData.DataRows[rowNumberj], ref cont0, ref cont1, ref checkList);
                        List<string> l4 = new List<string>(l1);
                        l4.Reverse();
                        if (l4[0] == "מס' זיהוי")
                        {
                            mortgage.mortgageBorower.borowerIDNumber.Add(l2[0]);
                        }
                        if (l4[1] == "סוג זיהוי")
                        {
                            mortgage.mortgageBorower.borowerIDType.Add(l2[1]);
                        }
                        if (l4[2] == "שם הלווה")
                        {
                            mortgage.mortgageBorower.borowerName.Add(l2[2]);
                        }
                        rowNumberj++;
                        if (!ClassUtils.isArrayIncludeOneOfStringParam(slExcelData.DataRows[rowNumberj], "חברה", "ת.ז", "דרכון") && ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "דרגה") == -1)
                        {
                            int jjj = mortgage.mortgageBorower.borowerName.Count;
                            mortgage.mortgageBorower.borowerName[jjj - 1] = mortgage.mortgageBorower.borowerName[jjj - 1] + ClassUtils.buildCombinedline(slExcelData.DataRows[rowNumberj]);
                            rowNumberj++;
                        }


                    } while (ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "דרגה") == -1);

                }
                if (ClassUtils.isArrayIncludString(slExcelData.DataRows[rowNumberj], "דרגה") > -1)
                {
                    l1 = ClasszhuiotUtils.rawlineToKeystage(slExcelData.DataRows[rowNumberj]);
                    l2 = slExcelData.DataRows[rowNumberj + 1];
                    l2.Reverse();
                    int top = 0;
                    for ( int j = 0; j < l1.Count; j++)
                    {
                        if (l1[j] == "דרגה")
                        {
                            mortgage.grade = l2[top];
                            top++;
                            continue;
                        }
                        if (l1[j] == "סכום")
                        {
                            if (ClassUtils.isFloating(l2[top]))
                            {
                                mortgage.amount = l2[top + 1] + " " + l2[top];
                                top++;
                                top++;
                            }
                            else
                            {
                                mortgage.amount = l2[top] + " " + l2[top + 1] + " " + l2[top + 2];
                                top++; top++; top++;
                            }
                            continue;
                        }
                        if (l1[j] == "בתנאי שטר מקורי")
                        {
                            if (ClassUtils.isShtarNumber(l2[top]))
                            {
                                mortgage.OriginalShtarNum = l2[top];
                                top++;
                            }
                        }
                        if (l1[j] == "החלק בנכס")
                        {
                            if (ClassUtils.containHebrew(l2[top]))
                            {
                                for ( int k = top; k < l2.Count; k++)
                                {
                                    mortgage.propPart = mortgage.propPart + " "+ l2[k]; 
                                }
                            }
                            else
                            {
                                mortgage.propPart = l2[top + 2] + " " + l2[top + 1] + " " + l2[top];
                            }
                        }
                    }
                    rowNumberj++;
                    rowNumberj++;
                }

                for ( int j = rowNumberj; j < MortgageRows[i+1]; j++)
                {
                    string sss = ClassUtils.buildCombinedline(slExcelData.DataRows[j]);
                    mortgage.remarks.Add(sss);
                }
                mortgages.Add(mortgage);
            }
            
        }
        public void buildLeasingCSV(List<int> sec, List<int> LeasingRows , ClassExcelOperations excelOperations, string fileName)
        {
            leasings = new List<Leasing>();  // one for all Leasings

            List<string> l1 = new List<string>();
            List<string> l2 = new List<string>();
            int rowNumberj;
            rowNumberj = LeasingRows[0];
            for (int i = 0; i < LeasingRows.Count - 1; i++)
            {
                Leasing leasing = new Leasing(); // one for each part of leasing
                try
                {
                    bool withinfLeasingSection = true;
                    l1 = ClasszhuiotUtils.rawlineToKeyPropOwners(slExcelData.DataRows[rowNumberj]);
                    rowNumberj++;
                    do
                    {
                        l2 = ClasszhuiotUtils.rawlineToValuesLeasing0(l1, slExcelData.DataRows[rowNumberj]);
                        if ( l2.Count != l1.Count)
                        {
                            throw new Exception("שגיאת נתונים");
                        }
                        l2.Reverse();

                        //                   l2 = utilities.rawlineToValuesLeasing(slExcelData.DataRows[rowNumberj]);
                        rowNumberj++;
                        if (l1.Count == l2.Count)
                        {
                            LeasingOwner LeasingOwner = new LeasingOwner();
                            for (int k = 0; k < l1.Count; k++)
                            {
                                if (l1[k] == "מס' שטר")
                                {
                                    LeasingOwner.shtarNum = l2[k];
                                }
                                else if (l1[k] == "תאריך")
                                {
                                    LeasingOwner.date = l2[k];
                                }
                                else if (l1[k] == "מהות פעולה")
                                {
                                    LeasingOwner.transactionType = l2[k];
                                }
                                else if (l1[k] == "שם החוכר")
                                {
                                    LeasingOwner.LeaserName = l2[k];
                                }
                                else if (l1[k] == "סוג זיהוי")
                                {
                                    LeasingOwner.idType = l2[k];
                                }
                                else if (l1[k] == "מס' זיהוי")
                                {
                                    bool foreign = ClassUtils.isForeignID(l2[k]);
                                    if (foreign)
                                    {
                                        l2[k] = ClassUtils.Reverse(l2[k]);
                                    }
                                    LeasingOwner.idNumber = l2[k];
                                }
                            }
                            while (!ClassUtils.isArrayIncludeAllStringsParam(slExcelData.DataRows[rowNumberj], "החלק", "בזכות"))
                            {
                                string s1 = ClassUtils.buildCombinedline(slExcelData.DataRows[rowNumberj]);
                                LeasingOwner.LeaserName = LeasingOwner.LeaserName + " " + s1;
                                rowNumberj++;
                            }
                            if (ClassUtils.isArrayIncludeAllStringsParam(slExcelData.DataRows[rowNumberj], "החלק", "בזכות"))
                            {
                                rowNumberj++;
                                LeasingOwner.LeaserPart = ClassUtils.buildReverseCombinedLine(slExcelData.DataRows[rowNumberj]);
                            }
                            if (ClassUtils.isArrayIncludeAllStringsParam(slExcelData.DataRows[rowNumberj + 1], "הערות"))
                            {
                                rowNumberj++;
                                LeasingOwner.Remarks = ClassUtils.buildCombinedline(slExcelData.DataRows[rowNumberj]);
                            }
                            leasing.leasingOwners.Add(LeasingOwner);
                            rowNumberj++;
                            if (ClassUtils.isArrayIncludeAllStringsParam(slExcelData.DataRows[rowNumberj], "רמת", "חכירה"))
                            {
                                withinfLeasingSection = false;
                            }
                        }
                    } while (withinfLeasingSection);
                    l1 = ClasszhuiotUtils.rawlineToKeyPropLeasing(slExcelData.DataRows[rowNumberj]);
                    rowNumberj++;
                    l2 = ClasszhuiotUtils.rawlineToValuesLeasing1(slExcelData.DataRows[rowNumberj]);
                    if (l1.Count == l2.Count)
                    {
                        for (int k = 0; k < l1.Count; k++)
                        {
                            if (l1[k] == "רמת חכירה")
                            {
                                leasing.LeaserLevel = l2[k];
                            }
                            else if (l1[k] == "בתנאי שטר מקורי")
                            {
                                leasing.OriginlShtar = l2[k];
                            }
                            else if (l1[k] == "תאריך סיום")
                            {
                                leasing.EndDate = l2[k];
                            }
                            else if (l1[k] == "תקופה בשנים")
                            {
                                leasing.EndDate = l2[k];
                            }
                        }
                    }
                    rowNumberj++;
                    l1 = ClasszhuiotUtils.rawlineToKeyPropLeasing(slExcelData.DataRows[rowNumberj]);
                    rowNumberj++;
                    l2 = ClasszhuiotUtils.rawlineToValuesLeasing3(slExcelData.DataRows[rowNumberj]);
                    rowNumberj++;
                    if (l1[0] == "החלק בנכס")
                    {
                        leasing.PropertyPart = l2[0];
                    }
                    if (l2.Count > 1)
                    {
                        leasing.remarks.Add(l2[1]);
                    }
                    string ss = ClassUtils.buildCombinedline(slExcelData.DataRows[rowNumberj]);
                    leasing.remarks.Add(ss);
                    rowNumberj++;
                    while (rowNumberj < LeasingRows[i + 1])
                    {
                        string sss = ClassUtils.buildCombinedline(slExcelData.DataRows[rowNumberj]);
                        leasing.remarks.Add(sss);
                        rowNumberj++;
                    }
                }
                catch (Exception e)
                {
                    excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimError);
                    int row = excelOperations.getBatimErrorLine();
                    int newrow = row + 1;
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 1, fileName);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 2, "חכירות");
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 3, ClassUtils.buildReverseCombinedLine(slExcelData.DataRows[rowNumberj]));
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 4, rowNumberj.ToString()) ;
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 5, e.ToString());
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, 1, 6, newrow.ToString());
                    excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimError, false, 6, row, 1);
                }
                leasings.Add(leasing);
            }
        }

    }
}
