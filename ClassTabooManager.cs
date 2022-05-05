using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PDF2ExcelVsto.ClassTaboo;

namespace PDF2ExcelVsto
{
    class ClassTabooManager
    {
        public ClassFilesHandle filesHandle;
        public List<ClassTaboo> allTaboo;
        List<string> zhuiotCSVfile;
        List<string> batimCSVfile;
        ClassExcelOperations excelOperations;

        public ClassTabooManager(ClassFilesHandle fhd, ClassExcelOperations excel)
        {
            filesHandle = fhd;
            excelOperations = excel;
        }

        public List<ClassTaboo> getAllTaboo()
        {
            return allTaboo;
        }

        public void convertCSVtoExcel()
        {
            allTaboo = new List<ClassTaboo>();
            zhuiotCSVfile = filesHandle.getCSVFiles();
            foreach(string csvfile in zhuiotCSVfile)
            {
                int currentRow;
                SLExcelData slExcelData = new SLExcelData();
                slExcelData.DataRows = File2Data(csvfile);
                string fn = Path.GetFileName(csvfile);
                fn = fn.Replace("csv", "pdf");
                ClassTaboo taboo = new ClassTaboo(slExcelData, fn);
                currentRow = taboo.buildHeader();
                ClassMapTaboo TaboMap = new ClassMapTaboo(taboo.slExcelData.DataRows);
                if (TaboMap.isSectionExists("PropCreation"))
                {
                    List<int> section = TaboMap.getRowsofSection("PropCreation");
                    taboo.buildNozarCSV(section);
                }
                if (TaboMap.isSectionExists("PropDescription"))
                {
                    List<int> section = TaboMap.getRowsofSection("PropDescription");
                    List<int> DescriptionRows = TaboMap.getDescriptionRows();
                    taboo.buildDesciptionPropertyCSV(section, DescriptionRows);
                }
                if (TaboMap.isSectionExists("PropOwners"))
                {
                    List<int> section = TaboMap.getRowsofSection("PropOwners");
                    List<int> OwnersRows = TaboMap.GetOwnersRows();
                    taboo.buildOwnersZhuiotCSV(section, OwnersRows);
                }
                if (TaboMap.isSectionExists("Leasing"))
                {
                    List<int> section = TaboMap.getRowsofSection("Leasing");
                    List<int> LeasingRows = TaboMap.GetLeasingRows();
                    taboo.buildLeasingCSV(section, LeasingRows);
                }
                if (TaboMap.isSectionExists("Mortgage"))
                {
                    List<int> section = TaboMap.getRowsofSection("Mortgage");
                    List<int> MortgageRows = TaboMap.GetMortGageRows();
                    taboo.buildMortGage(section, MortgageRows);
                }
                if (TaboMap.isSectionExists("Remarks"))
                {
                    List<int> section = TaboMap.getRowsofSection("Remarks");
                    List<int> RemarksRows = TaboMap.GetRemarksRows();
                    taboo.buildRemarks(section, RemarksRows);
                }
                allTaboo.Add(taboo);
            }
        }

        public List<List<string>> File2Data(string csvfile)
        {
            List<List<string>> datas = new List<List<string>> ();
            List<string> sss = File.ReadAllLines(csvfile).ToList();
            for (int i = 0; i < sss.Count; i++)
            {
                string[] subs = sss[i].Split(' ');
                List<string> lll = subs.ToList();
                datas.Add(lll);
            }
            return datas;
        }
        public void CreateMortGageTables()
        {
            //            if ((allTaboo is null)) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.Mortgage, "משכנתאות", Color.Blue);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.Mortgage);
            excelOperations.refreshAll();
            int currentrow;
            int startrow;
            startrow = excelOperations.BuildMortgageHeader();
            currentrow = startrow;
            excelOperations.refreshAll();
            foreach (ClassTaboo taboo in allTaboo)
            {
                if ((taboo.mortgages is null)) continue;
                int section = currentrow;
                int owners = currentrow;
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 1, taboo.header.gush);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 2, taboo.header.helka);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 13, taboo.PDFFileName);
                foreach (Mortgage mrt in taboo.mortgages)
                {
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 3, mrt.shtarNum);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 4, mrt.date);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 5, mrt.MortgageType);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 6, mrt.mortgageOwner);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 7, mrt.grade);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 8, mrt.amount);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 9, mrt.OriginalShtarNum);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 10, mrt.propPart);
                    foreach ( string sss in mrt.remarks)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, section, 11, sss);
                        section++;
                    }
                    foreach ( string sss in mrt.onOwner)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, owners, 12, sss);
                        owners++;
                    }
                    currentrow = Math.Max(section, currentrow);
                    currentrow = Math.Max(owners, currentrow);
                }
                int endrow;
                endrow = Math.Max(section, currentrow);
                endrow = Math.Max(owners, currentrow);
                excelOperations.setBoarder(ClassExcelOperations.Sheets.Mortgage, startrow, endrow-1, 1, 13, 1);
                currentrow = endrow;
                startrow = currentrow;
            }
        }
        public void CreateRemarksTables()
        {
            excelOperations.createSheet(ClassExcelOperations.Sheets.Remark, "הערות", Color.Blue);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.Remark);
            excelOperations.refreshAll();
            int currentrow;
            int topBox;
            currentrow = excelOperations.BuildRemarkHeader();
            excelOperations.refreshAll();
            foreach (ClassTaboo taboo in allTaboo)
            {
                if ((taboo.remarks is null)) continue;
                topBox = currentrow;
                int beginFrame;
                int sectionrow = currentrow;
                int remarkrows = currentrow;
                beginFrame = currentrow;
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, currentrow, 1, taboo.header.gush);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, currentrow, 2, taboo.header.helka);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, currentrow, 10, taboo.PDFFileName);
                foreach ( Remarks remark in taboo.remarks)
                {
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, currentrow, 3, remark.shtarNum);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, currentrow, 4, remark.date);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, currentrow, 5, remark.actionType);
                    sectionrow = currentrow;
                    remarkrows = currentrow;

                    for (int i = 0 ; i < remark.onOwner.Count; sectionrow++ , i++)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, sectionrow, 6, remark.onOwner[i]);

                        if (remark.idType.Count >= remark.onOwner.Count) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, sectionrow, 7, remark.idType[i]);
                        if (remark.idNumber.Count >= remark.onOwner.Count) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, sectionrow, 8, remark.idNumber[i]);
                    }
                    for ( int i = 0; i < remark.remarks.Count; remarkrows++ , i++)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, remarkrows, 9, remark.remarks[i]);                       
                    }
                    excelOperations.refreshAll();
                    currentrow = Math.Max(remarkrows, sectionrow);
                    excelOperations.setBoarder(ClassExcelOperations.Sheets.Remark, beginFrame, currentrow-1, 1, 10, 2);
//                    excelOperations.drawLine(ClassExcelOperations.Sheets.Remark, currentrow, 1, 9);
                    excelOperations.refreshAll();
                }
                int endrow = Math.Max(remarkrows, sectionrow) -1;
                excelOperations.setBoarder(ClassExcelOperations.Sheets.Remark, topBox, endrow, 1, 10, 1);
                currentrow = endrow+1;
            }
        }

        public void CreateLeasingTables()
        {
            if ((allTaboo is null)) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.Leasing, "חכירות", Color.Blue);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.Leasing);
            excelOperations.refreshAll();
            int currentrow;            
            currentrow = excelOperations.BuildLeasingHeader();
            excelOperations.refreshAll();
            foreach (ClassTaboo taboo in allTaboo)
            {
                if ( (taboo.leasings is null) ) continue;
                int leasinRow = currentrow;
                int sectionrow = currentrow;
                int remarkrows = currentrow;
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, currentrow, 1, taboo.header.gush);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, currentrow, 2, taboo.header.helka);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, currentrow, 15, taboo.PDFFileName);
                foreach (Leasing leas in taboo.leasings)
                {
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, leasinRow, 10, leas.LeaserLevel);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, leasinRow, 11, leas.OriginlShtar);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, leasinRow, 12, leas.EndDate);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, leasinRow, 13, leas.PropertyPart);

                    foreach (LeasingOwner own in leas.leasingOwners)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, sectionrow, 3, own.shtarNum);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, sectionrow, 4, own.date);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, sectionrow, 5, own.transactionType);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, sectionrow, 6, own.LeaserName);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, sectionrow, 7, own.idType);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, sectionrow, 8, own.idNumber);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, sectionrow, 9, own.LeaserPart);
                        excelOperations.refreshAll();
                        sectionrow++;
                    }
                    
                    foreach ( string ss in leas.remarks)
                    {                        
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, remarkrows, 14, ss);
                        remarkrows++;
                    }
                    excelOperations.refreshAll();
//                   
                }
                int endrow = Math.Max(remarkrows, sectionrow);
                excelOperations.setBoarder(ClassExcelOperations.Sheets.Leasing, currentrow, endrow, 1, 15, 1);
                currentrow = endrow;
            }
        }
        public void CreatePropertyTables()
        {
            if ((allTaboo is null)) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.Property, "תאור הנכס", Color.Blue);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.Property);
            int currentrow;
            int serialNum = 1;
            currentrow = excelOperations.BuildPropertyHeader();
            foreach (ClassTaboo taboo in allTaboo)
            {
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 1, serialNum.ToString());
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 2, taboo.header.gush);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 3, taboo.header.helka);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 4, taboo.description.rashuiot +" " + taboo.description.rashuiot1 + taboo.description.rashuiot2);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 5, taboo.description.area);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 6, taboo.description.landType);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 7, taboo.description.remarks);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 8, taboo.description.oldNumbers);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 9, taboo.PDFFileName);
                currentrow++;
                serialNum++;
            }
        }
        public void CreateOwnersTable()
        {
            if ((allTaboo is null)) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.Owner, "בעלים", Color.Blue);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.Owner);
            excelOperations.refreshAll();
            int currentrow;
            int topBox= 0;
            int serialNum = 1;
            currentrow = excelOperations.buildOwnerHeadr0();
            excelOperations.refreshAll();
            foreach (ClassTaboo taboo in allTaboo)
            {
                topBox = currentrow;
                int beginFrame;
                beginFrame = currentrow;
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 2, taboo.header.gush);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 3, taboo.header.helka);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 12, taboo.PDFFileName);
                foreach (ClassTaboo.ZhuiotOwner own in taboo.zhuiotOwners)
                {
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 1, serialNum.ToString());
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 4, own.shtarNum);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 5, own.date);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 6, own.transactionType);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 7, own.ownerName);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 8, own.idType);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 9, own.idNumber);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 10, own.ownerPart);
                    string vvv = ClassUtils.convertPartToPercent(own.ownerPart);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 11, vvv);
                    excelOperations.refreshAll();
                    currentrow++;
                    serialNum++;
                }
                excelOperations.setBoarder(ClassExcelOperations.Sheets.Owner, beginFrame, currentrow - 1, 1, 12, 1);
                excelOperations.refreshAll();
            }
        }
    }
}
