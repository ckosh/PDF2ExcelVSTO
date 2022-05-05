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
    class ClasszhuiotManager 
    {
        public ClassFilesHandle filesHandle;
        public List<ClassTaboo> allTaboo;
        List<string> zhuiotCSVfile;
        ClassExcelOperations excelOperations;

        public ClasszhuiotManager(ClassFilesHandle fhd, ClassExcelOperations excel)
        {
            filesHandle = fhd;
            excelOperations = excel;
        }
        public void convertZhuiottoExcel()
        {
            allTaboo = new List<ClassTaboo>();
            
            zhuiotCSVfile = filesHandle.getCSVFiles("zhuiot");
            foreach (string csvfile in zhuiotCSVfile)
            {
                int currentRow;
                SLExcelData slExcelData = new SLExcelData();
                slExcelData.DataRows = ClassUtils.File2Data(csvfile);
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
                    taboo.buildLeasingCSV(section, LeasingRows , excelOperations, csvfile );
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
                taboo.Name = "tabu";
                taboo.gush = taboo.header.gush;
                taboo.helka = taboo.header.helka;
                allTaboo.Add(taboo);
            }

            Dictionary<ClassTaboo, int> myDict = new Dictionary<ClassTaboo, int>();
            for ( int i = 0; i < allTaboo.Count; i++)
            {
                myDict.Add(allTaboo[i], Int32.Parse(allTaboo[i].header.helka));
            }
            var sortedDict = from entry in myDict orderby entry.Value ascending select entry;
            allTaboo.Clear();
            for ( int i = 0; i < sortedDict.Count(); i++)
            {
                allTaboo.Add(sortedDict.ElementAt(i).Key);
            }
        }
        public void CreateMortGageTables()
        {
            if ((allTaboo is null) || allTaboo.Count == 0) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.Mortgage, "ז-משכנתאות", Color.Blue);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.Mortgage);
            excelOperations.refreshAll();
            int currentrow;
            
            currentrow = excelOperations.BuildMortgageHeader();
            int borrow =0;
            int owners =0;
            int remarks =0;
            int startrow;

            excelOperations.refreshAll();
            foreach (ClassTaboo taboo in allTaboo)
            {
                if ((taboo.mortgages is null)) continue;
                startrow = currentrow;
                borrow = currentrow;
                owners = currentrow;
                remarks = currentrow;
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 1, taboo.header.gush);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 2, taboo.header.helka);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 18, taboo.PDFFileName);
                foreach (Mortgage mrt in taboo.mortgages)
                {
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 3, mrt.shtarNum);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 4, mrt.date);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 5, mrt.MortgageType);
 
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 12, mrt.grade);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 13, mrt.amount);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 14, mrt.OriginalShtarNum);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 15, mrt.propPart);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, currentrow, 16, ClassUtils.convertPartToPercent(mrt.propPart));
                    for ( int i = 0; i < mrt.mortgageOwner.ownerName.Count; i++)
                    {
                        fillList(mrt.mortgageOwner.ownerIDType, mrt.mortgageOwner.ownerName.Count);
                        fillList(mrt.mortgageOwner.ownerIDNumber, mrt.mortgageOwner.ownerName.Count);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, owners, 6, mrt.mortgageOwner.ownerName[i]);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, owners, 7, mrt.mortgageOwner.ownerIDType[i]);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, owners, 8, mrt.mortgageOwner.ownerIDNumber[i]);
                        owners++;
                    }
                    for ( int i = 0; i < mrt.mortgageBorower.borowerName.Count; i++)
                    {
                        fillList(mrt.mortgageBorower.borowerIDType, mrt.mortgageBorower.borowerName.Count);
                        fillList(mrt.mortgageBorower.borowerIDNumber, mrt.mortgageBorower.borowerName.Count);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, borrow, 9, mrt.mortgageBorower.borowerName[i]);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, borrow, 10, mrt.mortgageBorower.borowerIDType[i]);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, borrow, 11, mrt.mortgageBorower.borowerIDNumber[i]);
                        borrow++;
                    }
                    for ( int i = 0; i < mrt.remarks.Count; i++)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Mortgage, remarks, 17, mrt.remarks[i]);
                        remarks++;
                    }
                    currentrow = Math.Max(borrow, owners);
                    currentrow = Math.Max(currentrow, remarks);
                    borrow = currentrow;
                    owners = currentrow;
                    remarks = currentrow;

                    try
                    {
                        excelOperations.CorrectFormatForSum(ClassExcelOperations.Sheets.Mortgage, 16, startrow, currentrow - 1, "0.00000%");
                    }
                    catch (Exception e)
                    {

                    }
                }
                excelOperations.setBoarder(ClassExcelOperations.Sheets.Mortgage, startrow, currentrow - 1, 1, 17, 2);
                excelOperations.addNameRange(ClassExcelOperations.Sheets.Mortgage, startrow, currentrow-1, 1, 17, Convert.ToInt32(taboo.header.gush), Convert.ToInt32(taboo.header.helka), 0, "ZM");

                //    int endrow;
                //    endrow = Math.Max(section, currentrow);
                //    endrow = Math.Max(owners, currentrow);
                //    excelOperations.setBoarder(ClassExcelOperations.Sheets.Mortgage, startrow, endrow - 1, 1, 13, 1);
                //    currentrow = endrow;
                //    startrow = currentrow;

            }
            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.Mortgage, false , 17, currentrow,2);
        }
        public void CreateRemarksTables()
        {
            if ((allTaboo is null) || allTaboo.Count == 0) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.Remark, "ז-הערות", Color.Blue);
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
                foreach (Remarks remark in taboo.remarks)
                {
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, currentrow, 3, remark.shtarNum);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, currentrow, 4, remark.date);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, currentrow, 5, remark.actionType);
                    sectionrow = currentrow;
                    remarkrows = currentrow;

                    for (int i = 0; i < remark.onOwner.Count; sectionrow++, i++)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, sectionrow, 6, remark.onOwner[i]);

                        if (remark.idType.Count >= remark.onOwner.Count) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, sectionrow, 7, remark.idType[i]);
                        if (remark.idNumber.Count >= remark.onOwner.Count) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, sectionrow, 8, remark.idNumber[i]);
                    }
                    for (int i = 0; i < remark.remarks.Count; remarkrows++, i++)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Remark, remarkrows, 9, remark.remarks[i]);
                    }
                    excelOperations.refreshAll();
                    currentrow = Math.Max(remarkrows, sectionrow);
                    excelOperations.setBoarder(ClassExcelOperations.Sheets.Remark, beginFrame, currentrow - 1, 1, 10, 2);
                    excelOperations.addNameRange(ClassExcelOperations.Sheets.Remark, beginFrame, currentrow - 1, 1, 10, Convert.ToInt32(taboo.header.gush), Convert.ToInt32(taboo.header.helka), 0, "ZR");
                    //                    excelOperations.drawLine(ClassExcelOperations.Sheets.Remark, currentrow, 1, 9);
                    excelOperations.refreshAll();
                }
                int endrow = Math.Max(remarkrows, sectionrow) - 1;
                excelOperations.setBoarder(ClassExcelOperations.Sheets.Remark, topBox, endrow, 1, 10, 1);
                

                currentrow = endrow + 1;
            }
            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.Remark, false, 10, currentrow,2);
        }

        public void CreateLeasingTables()
        {
            if ((allTaboo is null) || allTaboo.Count == 0) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.Leasing, "ז-חכירות", Color.Blue);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.Leasing);
            excelOperations.refreshAll();
            int currentrow;
            currentrow = excelOperations.BuildLeasingHeader();
            excelOperations.refreshAll();
            foreach (ClassTaboo taboo in allTaboo)
            {
                if ((taboo.leasings is null)) continue;
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

                    foreach (string ss in leas.remarks)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Leasing, remarkrows, 14, ss);
                        remarkrows++;
                    }
                    excelOperations.refreshAll();
                    //                   
                }
                int endrow = Math.Max(remarkrows, sectionrow);
                excelOperations.setBoarder(ClassExcelOperations.Sheets.Leasing, currentrow, endrow-1, 1, 15, 1);
                excelOperations.addNameRange(ClassExcelOperations.Sheets.Leasing, currentrow, endrow-1 , 1, 15, Convert.ToInt32(taboo.header.gush), Convert.ToInt32(taboo.header.helka), 0, "ZL");

                currentrow = endrow;
            }
            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.Leasing, false, 15, currentrow,2);
        }
        public void CreatePropertyTables()
        {
            if ((allTaboo is null)  || allTaboo.Count == 0) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.Property, "ז-תאור הנכס", Color.Blue);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.Property);
            int currentrow;
            int serialNum = 1;
            currentrow = excelOperations.BuildPropertyHeader();
            foreach (ClassTaboo taboo in allTaboo)
            {
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 1, serialNum.ToString());
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 2, taboo.header.gush);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 3, taboo.header.helka);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 4, taboo.description.rashuiot + " " + taboo.description.rashuiot1 + taboo.description.rashuiot2);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 5, taboo.description.area);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 6, taboo.description.landType);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 7, taboo.description.remarks);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 8, taboo.description.oldNumbers);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Property, currentrow, 9, taboo.PDFFileName);
                currentrow++;
                serialNum++;
            }
            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.Property, false, 9, currentrow,2);
        }
        public void CreateOwnersTable()
        {
            if ((allTaboo is null) || allTaboo.Count == 0) return;
            
            excelOperations.createSheet(ClassExcelOperations.Sheets.Owner, "ז-בעלים", Color.Blue);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.Owner);
            excelOperations.refreshAll();
            int currentrow;
            int topBox = 0;
            int serialNum = 1;
            int beginFrame=0;
            currentrow = excelOperations.buildOwnerHeadr0();
            excelOperations.refreshAll();
            foreach (ClassTaboo taboo in allTaboo)
            {
                topBox = currentrow;
                bool MortfirstX = true;
                bool LeasfirstX = true;
                bool RemfirstX = true;
                beginFrame = currentrow;
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 2, taboo.header.gush);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 3, taboo.header.helka);
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 15, taboo.PDFFileName);
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
                    if ( taboo.mortgages != null && MortfirstX)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 12, "X");
                        excelOperations.createHyperLink(ClassExcelOperations.Sheets.Owner, currentrow, 12, taboo.header.gush, taboo.header.helka, "0", "ZM");
                        MortfirstX = false;
                    }
                    if ( taboo.leasings != null && LeasfirstX)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 13, "X");
                        excelOperations.createHyperLink(ClassExcelOperations.Sheets.Owner, currentrow, 13, taboo.header.gush, taboo.header.helka, "0", "ZL");
                        LeasfirstX = false;
                    }
                    if (taboo.remarks != null && RemfirstX)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.Owner, currentrow, 14, "X");
                        excelOperations.createHyperLink(ClassExcelOperations.Sheets.Owner, currentrow, 14, taboo.header.gush, taboo.header.helka, "0", "ZR");
                        RemfirstX = false;
                    }

                    excelOperations.refreshAll();
                    currentrow++;
                    serialNum++;
                }
                excelOperations.setBoarder(ClassExcelOperations.Sheets.Owner, beginFrame, currentrow - 1, 1, 14, 1);
                excelOperations.refreshAll();
                excelOperations.CorrectFormatForSum(ClassExcelOperations.Sheets.Owner, 11, beginFrame, currentrow - 1, "0.00000%");
            }
            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.Owner, false, 14, currentrow,2);
            
        }
        public void fillList(List<string> lst , int jj)
        {
            for ( int i = lst.Count ;  i < jj ; i++)
            {
                lst.Add(" ");
            }
        }
    }
}
