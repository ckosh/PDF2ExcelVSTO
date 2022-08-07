using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PDF2ExcelVsto.ClassExcelOperations;

namespace PDF2ExcelVsto 
{
    class ClassbatimManager0 
    {
        public ClassFilesHandle filesHandle;
        public ClassExcelOperations excelOperations;
        public List<Classbatim> allBatim;
        public List<string> batimCSVfile;

        public ClassbatimManager0(ClassFilesHandle fhd, ClassExcelOperations excel)
        {
            filesHandle = fhd;
            excelOperations = excel;
        }
        public void convertBatimtoExcel()
        {
            allBatim = new List<Classbatim>();
            
            batimCSVfile = filesHandle.getCSVFiles("batim");
            foreach (string csvfile in batimCSVfile)
            {
                int currentRow;
                SLExcelData slExcelData = new SLExcelData();
                slExcelData.DataRows = ClassUtils.File2Data(csvfile);
                string fn = Path.GetFileName(csvfile);
                fn = fn.Replace("csv", "pdf");
                Classbatim batim = new Classbatim(slExcelData, fn);
                currentRow = batim.buildHeader();
                batim.MapMainSections();
                batim.MapProperty();
                batim.MapSubSections();
                if (batim.nozar.line > -1)
                {
                    CreateNozarData(batim, slExcelData, batim.nozar.line);
                }
                if (batim.batimproperty.line > -1)
                {
                    CreateBatimProperty(batim, slExcelData);
                }
                createTatHelkot(batim, slExcelData);
                batim.Name = "batim";
                batim.gush = batim.header.gush;
                batim.helka = batim.header.helka;
                allBatim.Add(batim);
            }

            Dictionary<Classbatim, int> myDict = new Dictionary<Classbatim, int>();
            for (int i = 0; i < allBatim.Count; i++)
            {
                myDict.Add(allBatim[i], Int32.Parse(allBatim[i].header.helka));
            }
            var sortedDict = from entry in myDict orderby entry.Value ascending select entry;
            allBatim.Clear();
            for (int i = 0; i < sortedDict.Count(); i++)
            {
                allBatim.Add(sortedDict.ElementAt(i).Key);
            }

        }
        public void CreateBatimRemarksTables()
        {
            if ((allBatim is null)) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.BatimRemarks, "ב-הערות", Color.Green);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimRemarks);
            excelOperations.refreshAll();
            int currentrow = 0;
            currentrow = excelOperations.BuildBatimRemarksHeader();
            excelOperations.refreshAll();
            int startrow = 0;
            int startrowFile = 0;
            int subremarks = 0;
            int subremarks0 = 0;
            bool firstRow = true;
            startrowFile = currentrow;
            foreach (Classbatim batim in allBatim)
            {
                firstRow = true;
//                startrowFile = currentrow;
                if (batim.tatHelkot.Count == 0) continue;
                for (int i = 0; i < batim.tatHelkot.Count; i++)
                {
                    if (batim.tatHelkot[i].remarks.Count == 0) continue;
                    startrow = currentrow;
                    if (firstRow) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, currentrow, 1, batim.header.gush);
                    if (firstRow) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, currentrow, 2, batim.header.helka);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, currentrow, 3, batim.tatHelkot[i].number.ToString());
                    if (firstRow) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, currentrow, 11, batim.PDFFileName);
                    firstRow = false;
                    for (int j = 0; j < batim.tatHelkot[i].remarks.Count; j++)
                    {
                        Classbatim.Remark rem = batim.tatHelkot[i].remarks[j];
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, currentrow, 4, rem.remarkType);
                        subremarks = currentrow-1;
                        for (int k = 0; k < rem.name.Count; k++)
                        {
                            subremarks++;
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, subremarks, 5, rem.name[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, subremarks, 6, rem.idType[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, subremarks, 7, rem.idNumber[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, subremarks, 8, rem.part[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, subremarks, 9, rem.shtar[k]);                           
                        }
                        subremarks0 = currentrow - 1;
                        for ( int l = 0; l < rem.remarklines.Count; l++)
                        {
                            subremarks0++;
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimRemarks, subremarks0, 10, rem.remarklines[l]);
                        }
                        currentrow = Math.Max(currentrow, subremarks);
                        currentrow = Math.Max(currentrow, subremarks0);
                        excelOperations.setBoarder(ClassExcelOperations.Sheets.BatimRemarks, startrow, currentrow , 1, 11, 2);
                        excelOperations.addNameRange(ClassExcelOperations.Sheets.BatimRemarks, startrow, currentrow , 1, 11, Convert.ToInt32(batim.header.gush), Convert.ToInt32(batim.header.helka), Convert.ToInt32(batim.tatHelkot[i].number), "BR");
                        currentrow++;
                    }
                }
                excelOperations.setBoarder(ClassExcelOperations.Sheets.BatimRemarks, startrowFile, currentrow - 1, 1, 11, 1);
 
            }
            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimRemarks, false, 11, currentrow, 2);
        }
        public void CreateBatimMortgage()
        {
            if ((allBatim is null)) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.BatimMortgage, "ב-משכנתאות", Color.Green);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimMortgage);
            excelOperations.refreshAll();
            int currentrow = 0;
            currentrow = excelOperations.BuildBatimMortgageHeader();
            excelOperations.refreshAll();
            int startrow = 0;
            int startrowFile = 0;
            bool firstRow = true;
            int startmainSection;
            int startRemarkSection;
            foreach (Classbatim batim in allBatim)
            {
                firstRow = true;
                startrowFile = currentrow;
                if (batim.tatHelkot.Count == 0) continue;
                for (int i = 0; i < batim.tatHelkot.Count; i++)
                {
                    if (batim.tatHelkot[i].mortgageTatHelkas.Count == 0) continue;
                    startrow = currentrow;
                    startmainSection = currentrow;
                    startRemarkSection = currentrow;

                    if (firstRow) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 1, batim.header.gush);
                    if (firstRow) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 2, batim.header.helka);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 3, batim.tatHelkot[i].number.ToString());
                    if (firstRow) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 12, batim.PDFFileName);
                    firstRow = false;
                    for (int j = 0; j < batim.tatHelkot[i].mortgageTatHelkas.Count; j++)
                    {
                        Classbatim.MortgageTatHelka mort = batim.tatHelkot[i].mortgageTatHelkas[j];
                        for ( int j1 = 0; j1 < mort.Name.Count; j1++)
                        {
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 4, mort.mtype[j1]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 5, mort.Name[j1]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 6, mort.idType[j1]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 7, mort.idNumber[j1]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 8, mort.part[j1]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 9, mort.shtar[j1]);
                            if ( j1 == 0 ) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startmainSection, 10, mort.grade[0]);
                            startmainSection++;
                        }
        
                        if ( mort.mortRemarks.Count > 0)
                        {
                            for ( int k = 0; k < mort.mortRemarks.Count; k++)
                            {
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimMortgage, startRemarkSection, 11, mort.mortRemarks[k]);
                                startRemarkSection++;
                            }
                            startRemarkSection--;
                        }
                        startRemarkSection++;
                        currentrow = Math.Max(startmainSection, startRemarkSection);                        
                    }
                    excelOperations.setBoarder(ClassExcelOperations.Sheets.BatimMortgage, startrow, currentrow - 1, 1, 12, 2);
                    excelOperations.addNameRange(ClassExcelOperations.Sheets.BatimMortgage, startrow, currentrow - 1, 1, 12, Convert.ToInt32(batim.header.gush),Convert.ToInt32(batim.header.helka), Convert.ToInt32(batim.tatHelkot[i].number),"BM");
                }
                excelOperations.setBoarder(ClassExcelOperations.Sheets.BatimMortgage, startrowFile, currentrow - 1, 1, 12, 1);
            }
            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimMortgage, false, 11, currentrow, 2);
        }
        public void createBatimAttachments()
        {
            if ((allBatim is null)) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.BatimAttachments, "ב-הצמדות", Color.Green);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimAttachments);
            excelOperations.refreshAll();
            int currentrow = 0;
            currentrow = excelOperations.BuildBatimAttachmentsHeader();
            excelOperations.refreshAll();
            int startrow = 0;
            int startrowFile = currentrow;
            foreach (Classbatim batim in allBatim)
            {
                if (batim.tatHelkot.Count == 0) continue;
                
                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimAttachments, currentrow, 9, batim.PDFFileName);
                for (int i = 0; i < batim.tatHelkot.Count; i++)
                {
                    if (batim.tatHelkot[i].attachments.Count == 0) continue;
                    startrow = currentrow;
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimAttachments, currentrow, 1, batim.header.gush);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimAttachments, currentrow, 2, batim.header.helka);

                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimAttachments, currentrow, 3, batim.tatHelkot[i].number.ToString());
                    Classbatim.Attachment att = batim.tatHelkot[i].attachments[0];

                    for (int j = 0; j < att.mark.Count; j++)
                    {
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimAttachments, currentrow, 4, att.mark[j]);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimAttachments, currentrow, 5, att.color[j]);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimAttachments, currentrow, 6, att.description[j]);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimAttachments, currentrow, 7, att.commonWith[j]);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimAttachments, currentrow, 8, att.area[j]);
                        currentrow++;
                    }
                    excelOperations.setBoarder(ClassExcelOperations.Sheets.BatimAttachments, startrow, currentrow - 1, 1, 9, 2);
                    excelOperations.addNameRange(ClassExcelOperations.Sheets.BatimAttachments, startrow, currentrow - 1, 1, 9, Convert.ToInt32(batim.header.gush), Convert.ToInt32(batim.header.helka), Convert.ToInt32(batim.tatHelkot[i].number), "BA");
                }
                excelOperations.setBoarder(ClassExcelOperations.Sheets.BatimAttachments, startrowFile, currentrow - 1, 1, 9, 1);
            }
            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimAttachments, false, 9, currentrow, 2);
        }            
        public void CreateBatimLeasing()
        {
            if ((allBatim is null)) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.BatimLeasing, "ב-חכירות", Color.Green);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimLeasing);
            excelOperations.refreshAll();
            int currentrow = 0;
            currentrow = excelOperations.BuildBatimLeasingHeader();
            excelOperations.refreshAll();
            int startrow = 0;
            foreach (Classbatim batim in allBatim)
            {
                if (batim.tatHelkot.Count == 0) continue;
                for (int i = 0; i < batim.tatHelkot.Count; i++)
                {
                    if (batim.tatHelkot[i].leasings.Count == 0) continue;
                    startrow = currentrow;

                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 1, batim.header.gush);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 2, batim.header.helka);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 3, batim.tatHelkot[i].number.ToString());
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 14, batim.PDFFileName);
                    for (int j = 0; j < batim.tatHelkot[i].leasings.Count; j++)
                    {
                        Classbatim.Leasing leas = batim.tatHelkot[i].leasings[j];
                        for (int k = 0; k < leas.leasingType.Count; k++)
                        {
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 4, leas.leasingType[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 5, leas.Name[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 6, leas.idtype[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 7, leas.id[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 8, leas.part[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 9, leas.shtar[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 10, leas.rama[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 11, leas.endDate[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 12, leas.remarks[k]);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimLeasing, currentrow, 13, leas.partpropery[k]);
                            currentrow++;
                        }
                    }
                    excelOperations.setBoarder(ClassExcelOperations.Sheets.BatimLeasing, startrow, currentrow - 1, 1, 14, 2);
                    excelOperations.addNameRange(ClassExcelOperations.Sheets.BatimLeasing, startrow, currentrow - 1, 1, 14, Convert.ToInt32(batim.header.gush), Convert.ToInt32(batim.header.helka), Convert.ToInt32(batim.tatHelkot[i].number), "BL");

                }
            }
            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimLeasing, false, 14, currentrow, 2);
        }
        public void CreateBatimOwnTable()
        {
            if ((allBatim is null)) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.BatimOwners, "ב-בעלים", Color.Green);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimOwners);
            excelOperations.refreshAll();
            int currentrow;
            int startrow = 0;
            int sectionStart = 0;
            currentrow = excelOperations.BuildBatimOwnerHeader();
            excelOperations.refreshAll();
            string fName = "";
            int presentTatHelka = 0;
                foreach (Classbatim batim in allBatim)
                {
                    sectionStart = currentrow;
                    try
                    {
                        fName = batim.PDFFileName;

                        if (batim.tatHelkot.Count == 0) continue;
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 1, batim.header.gush);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 2, batim.header.helka);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 3, batim.batimproperty.areasqmr);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 24, batim.PDFFileName);
                        for (int i = 0; i < batim.tatHelkot.Count; i++)
                        {
                            startrow = currentrow;
                            presentTatHelka = batim.tatHelkot[i].number;
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 4, batim.tatHelkot[i].number.ToString());
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 5, batim.tatHelkot[i].shetah);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 6, batim.tatHelkot[i].floor);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 7, batim.tatHelkot[i].entrance);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 8, batim.tatHelkot[i].agaf);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 9, batim.tatHelkot[i].building);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 10, batim.tatHelkot[i].partincommon);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 11, ClassUtils.convertPartToPercent(batim.tatHelkot[i].partincommon));

                            if (batim.tatHelkot[i].mortgageTatHelkas.Count > 0)
                            {
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 19, "X");
                                excelOperations.createHyperLink(ClassExcelOperations.Sheets.BatimOwners, currentrow, 19, batim.header.gush, batim.header.helka, batim.tatHelkot[i].number.ToString(), "BM");
                            }
                            if (batim.tatHelkot[i].remarks.Count > 0)
                            {
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 20, "X");
                                excelOperations.createHyperLink(ClassExcelOperations.Sheets.BatimOwners, currentrow, 20, batim.header.gush, batim.header.helka, batim.tatHelkot[i].number.ToString(), "BR");
                            }
                            if (batim.tatHelkot[i].leasings.Count > 0)
                            {
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 21, "X");
                                excelOperations.createHyperLink(ClassExcelOperations.Sheets.BatimOwners, currentrow, 21, batim.header.gush, batim.header.helka, batim.tatHelkot[i].number.ToString(), "BL");
                            }
                            if (batim.tatHelkot[i].attachments.Count > 0)
                            {
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 22, "X");
                                excelOperations.createHyperLink(ClassExcelOperations.Sheets.BatimOwners, currentrow, 22, batim.header.gush, batim.header.helka, batim.tatHelkot[i].number.ToString(), "BA");
                            }
                            if (batim.tatHelkot[i].lineZikotTatHelka > 0) excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 23, "X");
                            for (int j = 0; j < batim.tatHelkot[i].owners.Count; j++)
                            {
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 12, batim.tatHelkot[i].owners[j].transaction);
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 13, batim.tatHelkot[i].owners[j].name);
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 14, batim.tatHelkot[i].owners[j].idType);
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 15, batim.tatHelkot[i].owners[j].idNumber);
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 16, batim.tatHelkot[i].owners[j].part);
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 17, ClassUtils.convertPartToPercent(batim.tatHelkot[i].owners[j].part));
                                excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimOwners, currentrow, 18, batim.tatHelkot[i].owners[j].shtar);
                                currentrow++;
                            }
                            excelOperations.setBoarder(ClassExcelOperations.Sheets.BatimOwners, startrow, currentrow - 1, 1, 24, 2);
                        }
                    }
                    catch (Exception e)
                    {
                        excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimError);
                        int row = excelOperations.getBatimErrorLine();
                        int newrow = row + 1;
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 1, fName);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 2, "נתוני בעלים");
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 3, presentTatHelka.ToString()); // batim.presenttatHelka.ToString());
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 4, "...");// ClassUtils.buildReverseCombinedLine(slExcelData.DataRows[currentRow + 1]));
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 5, "??");
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, 1, 6, newrow.ToString());
                        excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimError, false, 6, row, 1);
                        excelOperations.paintRow(ClassExcelOperations.Sheets.BatimOwners, currentrow, 1, 24, Color.DarkGray);

                    }
                excelOperations.CorrectFormatForSum(ClassExcelOperations.Sheets.BatimOwners, 17, sectionStart, currentrow - 1, "0.00000%");
                excelOperations.CorrectFormatForSum(ClassExcelOperations.Sheets.BatimOwners, 11, sectionStart, currentrow - 1, "0.00000%");
            }
            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimOwners, false, 24, currentrow, 2);
        }
        public void CreatePropertyTable()
        {
            if ((allBatim is null)) return;
            excelOperations.createSheet(ClassExcelOperations.Sheets.BatimProperty, "ב-הרכוש המשותף", Color.Green);
            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimProperty);
            excelOperations.refreshAll();
            int currentrow;
            currentrow = excelOperations.BuildBatimPropertyHeader();
            int smallremark = 0;
            int zikotremark = 0;
            int remarkremark = 0;
            int startrow;
            string fName = "";
            
            excelOperations.refreshAll();
            foreach (Classbatim batim in allBatim)
            {
                try
                {
                    if ((batim.batimproperty is null)) continue;
                    smallremark = currentrow;
                    zikotremark = currentrow;
                    remarkremark = currentrow;
                    startrow = currentrow;
                    fName = batim.PDFFileName;

                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 1, batim.header.gush);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 2, batim.header.helka);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 19, batim.PDFFileName);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 18, batim.header.nesachNumber);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 17, batim.header.dateCalendar);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 3, batim.nozar.shtar);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 4, batim.nozar.date);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 5, batim.nozar.shtarType);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 6, batim.batimproperty.rashuiot);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 7, batim.batimproperty.areasqmr);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 8, batim.batimproperty.numOfTatHelkot);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 9, batim.batimproperty.takanon);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 10, batim.batimproperty.shtarYozer);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 11, batim.batimproperty.tikYozer);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 12, batim.batimproperty.tikbaitMeshutaf);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, currentrow, 13, batim.batimproperty.addtress);
                    smallremark = currentrow;
                    for (int j = 0; j < batim.batimproperty.smallremarks.Count; j++)
                    {
                        smallremark = (j > 0 ? smallremark + 1 : smallremark);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, smallremark, 14, batim.batimproperty.smallremarks[j]);
                    }
                    zikotremark = currentrow;
                    for (int j = 0; j < batim.batimproperty.zikotHnah.Count; j++)
                    {
                        zikotremark = (j > 0 ? zikotremark + 1 : zikotremark);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, zikotremark, 15, batim.batimproperty.zikotHnah[j]);
                    }
                    remarkremark = currentrow;
                    for (int j = 0; j < batim.batimproperty.remarks.Count; j++)
                    {
                        remarkremark = (j > 0 ? remarkremark + 1 : remarkremark);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimProperty, remarkremark, 16, batim.batimproperty.remarks[j]);
                    }
                    currentrow = Math.Max(currentrow, smallremark);
                    currentrow = Math.Max(currentrow, zikotremark);
                    currentrow = Math.Max(currentrow, remarkremark);
                    currentrow++;
                    excelOperations.setBoarder(ClassExcelOperations.Sheets.BatimProperty, startrow, currentrow - 1, 1, 19, 2);

                }
                catch (Exception e)
                {
                    excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimError);
                    int row = excelOperations.getBatimErrorLine();
                    int newrow = row + 1;
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 1, fName);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 2, "נתוני תת חלקה");
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 3, "..."); // batim.presenttatHelka.ToString());
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 4, "...");// ClassUtils.buildReverseCombinedLine(slExcelData.DataRows[currentRow + 1]));
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 5, "??");
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, 1, 6, newrow.ToString());

                    excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimError, false, 6, row, 1);
                    excelOperations.paintRow(ClassExcelOperations.Sheets.BatimProperty, currentrow, 1, 19, Color.DarkGray);
                }
            }


            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimProperty, false, 19, currentrow, 2);
        }
        public void CreateNozarData(Classbatim batim, SLExcelData slExcelData, int line)
        {
            List<string> res;
            res = ClassbatimUtils0.parseNozar(slExcelData.DataRows[batim.nozar.line]);
            batim.nozar.shtar = res[0];
            batim.nozar.date = res[1];
            batim.nozar.shtarType = res[2];
        }
        public void CreateBatimProperty(Classbatim batim, SLExcelData slExcelData)
        {
            int currentRow;
            currentRow = batim.batimproperty.line + 1;
            List<string> l1 = ClassbatimUtils0.ParseTopProperty(slExcelData.DataRows[currentRow]);
            List<string> l2 = ClassbatimUtils0.ParseSecondProperty(l1, slExcelData.DataRows[currentRow + 1]);
            int point = 0;
            if (l1.Count == l2.Count)
            {
                if (l1[point] == "רשויות")
                {
                    batim.batimproperty.rashuiot = l2[point];
                    point++;
                }
                if (l1[point] == "שטח במ\"ר")
                {
                    batim.batimproperty.areasqmr = l2[point];
                    point++;
                }
                if (l1[point] == "תת חלקות")
                {
                    batim.batimproperty.numOfTatHelkot = l2[point];
                    point++;
                }
                if (l1[point] == "תקנון")
                {
                    batim.batimproperty.takanon = l2[point];
                    point++;
                }
                if (l1[point] == "שטר יוצר")
                {
                    batim.batimproperty.shtarYozer = l2[point];
                    point++;
                }
                if (point <= l1.Count - 1)
                {
                    if (l1[point] == "תיק יוצר")
                    {
                        batim.batimproperty.tikYozer = l2[point];
                        point++;
                    }
                }
                if (point <= l1.Count - 1)
                {
                    if (l1[point] == "תיק בית משותף")
                    {
                        batim.batimproperty.tikbaitMeshutaf = l2[point];
                        point++;
                    }
                }
            }
            currentRow++;

            Dictionary<string, int> myDict = new Dictionary<string, int>();
            if (batim.batimproperty.lineAddress > 0) myDict.Add("address", batim.batimproperty.lineAddress);
            if (batim.batimproperty.lineRemark > 0) myDict.Add("remark0", batim.batimproperty.lineRemark);
            if (batim.batimproperty.lineRemark1 > 0) myDict.Add("remark1", batim.batimproperty.lineRemark1);
            if (batim.batimproperty.linezikot > 0) myDict.Add("zikot", batim.batimproperty.linezikot);
            myDict.Add("end", batim.tatHelkot[0].line);
            var sortedDict = from entry in myDict orderby entry.Value ascending select entry;
            List<string> temp;

            if (sortedDict.ElementAt(0).Value > currentRow + 1) // if Rashuiot name extens to two lines 
            {
                temp = new List<string>(slExcelData.DataRows[currentRow + 1]);
                batim.batimproperty.rashuiot = batim.batimproperty.rashuiot + " " + ClassUtils.buildCombinedline(temp);
                currentRow++;
            }
            for (int i = 0; i < sortedDict.Count() - 1; i++)
            {
                KeyValuePair<string, int> lastRow = sortedDict.ElementAt(i + 1);
                KeyValuePair<string, int> entry = sortedDict.ElementAt(i);
                if (entry.Key == "address")
                {
                    currentRow++;
                    temp = new List<string>(slExcelData.DataRows[entry.Value]);
                    temp = ClassUtils.reverseOrder(temp);

                    batim.batimproperty.addtress = ClassUtils.buildCombinedLineSelected(temp, 1, temp.Count, false);
                    currentRow++;
                    continue;
                }
                else if (entry.Key == "remark0")
                {
                    for (int j = entry.Value; j < lastRow.Value; j++)
                    {
                        temp = new List<string>(slExcelData.DataRows[j]);
                        temp = ClassUtils.reverseOrder(temp);
                        int from = (j == entry.Value ? 1 : 0);
                        batim.batimproperty.smallremarks.Add(ClassUtils.buildCombinedLineSelected(temp, from, temp.Count, false));
                        continue;
                    }
                }
                else if (entry.Key == "zikot")
                {
                    for (int j = entry.Value + 1; j < lastRow.Value; j++)
                    {
                        temp = new List<string>(slExcelData.DataRows[j]);
                        temp = ClassUtils.reverseOrder(temp);
                        batim.batimproperty.zikotHnah.Add(ClassUtils.buildCombinedLineSelected(temp, 0, temp.Count, false));
                        continue;
                    }
                }
                else if (entry.Key == "remark1")
                {
                    for (int j = entry.Value + 1; j < lastRow.Value; j++)
                    {
                        temp = new List<string>(slExcelData.DataRows[j]);
                        temp = ClassUtils.reverseOrder(temp);
                        batim.batimproperty.remarks.Add(ClassUtils.buildCombinedLineSelected(temp, 0, temp.Count, false));
                        continue;
                    }
                }
                continue;
            }
        }
        public void createTatHelkot(Classbatim batim, SLExcelData slExcelData)
        {
            int currentRow;
            List<string> temp;
            for (int i = 0; i < batim.tatHelkot.Count; i++)
            {
                currentRow = batim.tatHelkot[i].line;
                temp = new List<string>(slExcelData.DataRows[currentRow]);
                batim.tatHelkot[i].number = Int32.Parse(temp[0]);
                batim.presenttatHelka = Int32.Parse(temp[0]);
                currentRow++;
                try
                {
                    List<string> l2;
                    List<string> l1 = ClassbatimUtils0.ParseTopTatHelka(slExcelData.DataRows[currentRow]);
                    if (slExcelData.DataRows[currentRow + 1].Count < l1.Count)
                    {
                        string spatch = slExcelData.DataRows[currentRow + 1][0];
                        currentRow++;
                        slExcelData.DataRows[currentRow + 1].Insert(1, spatch);
                    }
                    l2 = ClassbatimUtils0.ParseSecondTatHelka(l1, slExcelData.DataRows[currentRow + 1]);
                    int point = 0;
                    if (l1.Count == l2.Count)
                    {
                        if (l1[point] == "שטח במ\"ר")
                        {
                            batim.tatHelkot[i].shetah = l2[point];
                            point++;
                        }
                        if (l1[l1.Count - 1] == "החלק ברכוש המשותף")
                        {
                            batim.tatHelkot[i].partincommon = l2[l2.Count - 1];
                        }
                        if (l1[point] == "תיאור קומה")
                        {
                            batim.tatHelkot[i].floor = l2[point];
                            point++;
                        }
                        if (l1[point] == "כניסה")
                        {
                            batim.tatHelkot[i].entrance = l2[point];
                            point++;
                        }
                        if (l1[point] == "אגף")
                        {
                            batim.tatHelkot[i].agaf = l2[point];
                            point++;
                        }
                        if (l1[point] == "מבנה")
                        {
                            batim.tatHelkot[i].building = l2[point];
                            point++;
                        }
                        //if (l1[point] == "החלק ברכוש המשותף")
                        //{
                        //    batim.tatHelkot[i].partincommon = l2[point];
                        //    point++;
                        //}
                    }
                }
                catch (Exception e)
                {
                    excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimError);
                    int row = excelOperations.getBatimErrorLine();
                    int newrow = row + 1;
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 1, batim.PDFFileName);
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 2, "נתוני תת חלקה");
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 3, batim.presenttatHelka.ToString());
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 4, ClassUtils.buildReverseCombinedLine(slExcelData.DataRows[currentRow + 1]));
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 5, "??");
                    excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, 1, 6, newrow.ToString());
                    
                    excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimError, false, 6, row, 1);

                }
                ///
                /// soert order
                /// 
                Dictionary<string, int> myDict = new Dictionary<string, int>();
                if (batim.tatHelkot[i].lineOwners > 0) myDict.Add("owners", batim.tatHelkot[i].lineOwners);
                if (batim.tatHelkot[i].lineAttachments > 0) myDict.Add("attachments", batim.tatHelkot[i].lineAttachments);
                if (batim.tatHelkot[i].lineLeasing > 0) myDict.Add("leasing", batim.tatHelkot[i].lineLeasing);
                if (batim.tatHelkot[i].lineRemarks > 0) myDict.Add("remarks", batim.tatHelkot[i].lineRemarks);
                if (batim.tatHelkot[i].lineMortgage > 0) myDict.Add("mortgage", batim.tatHelkot[i].lineMortgage);
                if (batim.tatHelkot[i].lineZikotTatHelka > 0) myDict.Add("zikotTatHelka", batim.tatHelkot[i].lineZikotTatHelka);
                if (i == batim.tatHelkot.Count - 1)
                {
                    myDict.Add("end", batim.endOfdata);
                }
                else
                {
                    myDict.Add("end", batim.tatHelkot[i + 1].line);
                }
                var sortedDict = from entry in myDict orderby entry.Value ascending select entry;

                ///// create owners
                ///
                for (int j = 0; j < sortedDict.Count() - 1; j++)
                {
                    KeyValuePair<string, int> lastRow = sortedDict.ElementAt(j + 1);
                    KeyValuePair<string, int> entry = sortedDict.ElementAt(j);
                    currentRow = entry.Value + 1;
                    if (entry.Key == "owners")
                    {
                        do
                        {
                            temp = new List<string>(slExcelData.DataRows[currentRow]);
                            temp = ClassUtils.reverseOrder(temp);
                            List<string> l3 = new List<string>();
                            string sss;

                            if (ClassUtils.isShtarNumber(temp[temp.Count - 1]))
                            {
                                Classbatim.Owner own = new Classbatim.Owner();
                                l3 = ClassbatimUtils0.parseOwners(temp);
                                own.transaction = l3[0];
                                own.name = l3[1];
                                own.idType = l3[2];
                                own.idNumber = l3[3];
                                own.part = l3[4];
                                own.shtar = l3[5];
                                batim.tatHelkot[i].owners.Add(own);
                            }
                            else
                            {
                                sss = ClassUtils.buildReverseCombinedLine(temp);
                                batim.tatHelkot[i].owners[batim.tatHelkot[i].owners.Count - 1].name = batim.tatHelkot[i].owners[batim.tatHelkot[i].owners.Count - 1].name + sss;
                            }
                            currentRow++;
                        } while (currentRow < lastRow.Value);
                    }
                }
                //
                // create leasing
                //
                for (int j = 0; j < sortedDict.Count() - 1; j++)
                {
                    KeyValuePair<string, int> lastRow = sortedDict.ElementAt(j + 1);
                    KeyValuePair<string, int> entry = sortedDict.ElementAt(j);
                    currentRow = entry.Value + 1;
                    if (entry.Key == "leasing")
                    {
                        try
                        {
                            int counter = 0;
                            bool first = true;
                            do
                            {
                                temp = new List<string>(slExcelData.DataRows[currentRow]);
                                temp = ClassUtils.reverseOrder(temp);
                                List<string> l3 = new List<string>();
                                if (ClassUtils.isShtarNumber(temp[temp.Count - 1]))
                                {
                                    bool skipNextLine = false ;
                                    Classbatim.Leasing leas = new Classbatim.Leasing();
                                    l3 = ClassbatimUtils0.parseLeaser(temp , ref skipNextLine);
                                    leas.leasingType.Add(l3[0]);
                                    leas.Name.Add(l3[1]);
                                    leas.idtype.Add(l3[2]);
                                    leas.id.Add(l3[3]);
                                    leas.part.Add(l3[4]);
                                    leas.shtar.Add(l3[5]);
                                    // dummy add
                                    leas.rama.Add("");
                                    leas.endDate.Add("");
                                    leas.remarks.Add("");
                                    leas.partpropery.Add("");
                                    batim.tatHelkot[i].leasings.Add(leas);
                                    if (first)
                                    {
                                        counter = batim.tatHelkot[i].leasings.Count;
                                        first = false;
                                    }
                                    currentRow++;
                                    if (skipNextLine) currentRow++;
                                    // check if name flows to other line
                                    List<string> temp1 = new List<string>(slExcelData.DataRows[currentRow]);
                                    temp1 = ClassUtils.reverseOrder(temp1);
                                    if(temp1.Count < 2)
                                    {
                                        string lastval = leas.Name[leas.Name.Count - 1];
                                        lastval = lastval + " " + temp1[0];
                                        leas.Name[leas.Name.Count - 1] = lastval;
                                        currentRow++;
                                    }
                                }
                                else if (ClassUtils.isArrayIncludeOneOfStringParam(temp, "רמה:", "סיום:")) // end of section 
                                {
                                    Classbatim.Leasing leas = batim.tatHelkot[i].leasings[counter - 1];
                                    l3 = ClassbatimUtils0.parseLastLeasRow(temp);
                                    leas.rama[0] = l3[0];
                                    leas.endDate[0] = l3[1];
                                    leas.remarks[0] = l3[2];
                                    leas.partpropery[0] = l3[3];
                                    currentRow++;
                                    first = true;
                                }
                                else
                                {
                                    throw new Exception("double line");
                                }
                            } while (currentRow < lastRow.Value);
                        }
                        catch (Exception e)
                        {
                            excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimError);
                            int row = excelOperations.getBatimErrorLine();
                            int newrow = row + 1;
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 1, batim.PDFFileName);
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 2, "חכירות");
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 3, batim.presenttatHelka.ToString());
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 4, ClassUtils.buildReverseCombinedLine(temp));
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 5, "??");
                            excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, 1, 6, newrow.ToString());
                            excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimError, false, 6, row, 1);
                        }

                    }
                }
                //
                // mortgage 
                //
                for (int j = 0; j < sortedDict.Count() - 1; j++)
                {
                    KeyValuePair<string, int> lastRow = sortedDict.ElementAt(j + 1);
                    KeyValuePair<string, int> entry = sortedDict.ElementAt(j);
                    currentRow = entry.Value + 1;
                    try
                    {
                        if (entry.Key == "mortgage")
                        {                       
                            do
                            {
                                Classbatim.MortgageTatHelka mort = new Classbatim.MortgageTatHelka();
                                string ac = "";

                                temp = new List<string>(slExcelData.DataRows[currentRow]);
                                temp = ClassUtils.reverseOrder(temp);
                                while (ClassUtils.isArrayIncludString(temp, "דרגה:") == -1)
                                {
                                    int numpar = 0;
                                    string MogType= "";
                                    string Name= "";
                                    string IDtype= "";
                                    string IDnumber= "";
                                    string part = "";
                                    string shtar = "";

                                    int[] array = new int[temp.Count];
                                    for (int j2 = 0; j2 < temp.Count; j2++) array[j2] = 0;

                                    int pos = ClassUtils.findShtarNumberWithin(temp);
                                    if (pos > -1)
                                    {
                                        shtar = temp[pos];
                                        numpar++;
                                        //mort.shtar.Add(temp[pos]);
                                        array[pos] = 1;
                                    }
                                    pos = ClassUtils.findIDNumberWithin(temp);
                                    if ( pos > -1)
                                    {
                                        IDnumber = temp[pos];
                                        numpar++;
                                        //mort.idNumber.Add(temp[pos]);
                                        array[pos] = 1;
                                    }
                                    pos = ClassUtils.findIDtypeWithin(temp);
                                    if ( pos > -1)
                                    {
                                        IDtype = temp[pos];
                                        numpar++;
                                        //mort.idType.Add(temp[pos]);
                                        array[pos] = 1;
                                    }
                                    List<int> poss = ClassUtils.findMortgageTypeWithin(temp);
                                    if ( poss.Count > 0)
                                    {
                                        ac = "";
                                        for ( int j1 = 0; j1 < poss.Count; j1++)
                                        {
                                            ac = ac + temp[poss[j1]] + " ";
                                            array[poss[j1]] = 1;
                                        }
                                        MogType = ac;
                                        numpar++;
                                        //mort.mtype.Add(ac);
                                    }
                                    poss = ClassUtils.findPartOfMortgage(temp);
                                    if (poss.Count > 0)
                                    {
                                        ac = "";
                                        for (int j1 = 0; j1 < poss.Count; j1++)
                                        {
                                            ac = ac + temp[poss[j1]] + " ";
                                            array[poss[j1]] = 1;
                                        }
                                        part = ac;
                                        numpar++;
                                        //mort.part.Add(ac);
                                    }
                                    // build remaining - name of mortgage holder
                                    ac = "";
                                    for ( int j1 = 0;j1 < temp.Count; j1++)
                                    {
                                        if (array[j1] == 0) ac = ac + temp[j1] + " ";
                                    }
                                    Name = ac;
                                    numpar++;
                                    //mort.Name.Add(ac);
                                    if ( numpar > 1)
                                    {
                                        mort.shtar.Add(shtar);
                                        mort.idNumber.Add(IDnumber);
                                        mort.idType.Add(IDtype);
                                        mort.mtype.Add(MogType);
                                        mort.part.Add(part);
                                        mort.Name.Add(Name);
                                    }
                                    else
                                    {
                                        int size = mort.Name.Count;
                                        string ss1 = mort.Name[size - 1] + Name; 
                                        mort.Name[size - 1] = ss1;
                                    }

                                    currentRow++;
                                    temp = new List<string>(slExcelData.DataRows[currentRow]);
                                    temp = ClassUtils.reverseOrder(temp);
                                }
                                if (ClassUtils.isArrayIncludString(temp, "דרגה:") > -1)
                                {
                                    mort.grade.Add(temp[1]);
                                }

                                int pos1 = ClassUtils.isArrayIncludString(temp, "הערות:");
                                if (pos1 > -1)
                                {
                                    String ttt = "";
                                    for (int k = 3; k < temp.Count; k++)
                                    {
                                        ttt = ttt + temp[k] + " ";
                                    }
                                    mort.mortRemarks.Add(ttt);
                                    currentRow++;
                                    while (currentRow < lastRow.Value)
                                    {
                                        ttt = "";
                                        List<string> temp1 = new List<string>(slExcelData.DataRows[currentRow]);
                                        temp1 = ClassUtils.reverseOrder(temp1);
                                        for (int k = 0; k < temp1.Count; k++)
                                        {
                                            ttt = ttt + temp1[k] + " ";
                                        }
                                        mort.mortRemarks.Add(ttt);
                                        currentRow++;
                                    }
                                }
                                batim.tatHelkot[i].mortgageTatHelkas.Add(mort);
                                currentRow++;
                            } while (currentRow < lastRow.Value);
                        }
                    }
                    catch (Exception e)
                    {
                        excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimError);
                        int row = excelOperations.getBatimErrorLine();
                        int newrow = row + 1;
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 1, batim.PDFFileName);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 2, "משכנתאות");
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 3, batim.presenttatHelka.ToString());
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 4, ClassUtils.buildReverseCombinedLine(temp));
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 5, "??");
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, 1, 6, newrow.ToString());
                        excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimError, false, 6, row, 1);
                    }
                }
                //
                //  remarks
                //
                for ( int j = 0; j < sortedDict.Count() - 1; j++)
                {
                    KeyValuePair<string, int> lastRow = sortedDict.ElementAt(j + 1);
                    KeyValuePair<string, int> entry = sortedDict.ElementAt(j);
                    currentRow = entry.Value + 1;
                    try
                    {
                        if (entry.Key == "remarks")
                        {
                            do
                            {
                                string cont = "";
                                bool firstLine = true;
                                Classbatim.Remark rem = new Classbatim.Remark();
                                do
                                {
                                    temp = new List<string>(slExcelData.DataRows[currentRow]);
                                    temp = ClassUtils.reverseOrder(temp);
                                    List<string> l3 = new List<string>();
                                    if ( firstLine)
                                    {
                                        l3 = ClassbatimUtils0.ParseRemark(temp, ref cont);
                                        rem.remarkType = l3[0];
                                    }
                                    else
                                    {
                                        l3 = ClassbatimUtils0.ParseRemarkSecondLine(temp);
                                    }                                    
                                    rem.name.Add(l3[1]);
                                    rem.idType.Add(l3[2]);
                                    rem.idNumber.Add(l3[3]);
                                    rem.part.Add(l3[4]);
                                    rem.shtar.Add(l3[5]);

                                    firstLine = false;
                                    currentRow++;
                                    if (cont != "") currentRow++;   /// skip a line because it took two lines 
                                    temp = new List<string>(slExcelData.DataRows[currentRow]);
                                    temp = ClassUtils.reverseOrder(temp);
                                } while (ClassUtils.isArrayIncludString(temp, "הערות:") == -1 && !ClassUtils.isShtarNumber(temp[temp.Count - 1]) && currentRow < lastRow.Value);
                                if (ClassUtils.isArrayIncludString(temp, "הערות:") > -1)
                                {
                                    rem.remarklines.Add(ClassUtils.buildReverseCombinedLine(temp));
                                    currentRow++;
                                }
                                while (!ClassUtils.isShtarNumber(temp[temp.Count - 1]) && currentRow < lastRow.Value  )
                                {
                                    rem.remarklines.Add(ClassUtils.buildReverseCombinedLine(temp));
                                    currentRow++;
                                    temp = new List<string>(slExcelData.DataRows[currentRow]);
                                    temp = ClassUtils.reverseOrder(temp);

                                } //while (!ClassUtils.isShtarNumber(temp[temp.Count - 1]) && currentRow < lastRow.Value);
                                
                                batim.tatHelkot[i].remarks.Add(rem);
                  
                            } while (currentRow < lastRow.Value);
                        }
                    }
                    catch (Exception e)
                    {
                        excelOperations.setActiveSheet(ClassExcelOperations.Sheets.BatimError);
                        int row = excelOperations.getBatimErrorLine();
                        int newrow = row + 1;
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 1, batim.PDFFileName);
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 2, "הערות");
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 3, batim.presenttatHelka.ToString());
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 4, ClassUtils.buildReverseCombinedLine(temp));
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 5, "??");
                        excelOperations.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, 1, 6, newrow.ToString());
                        excelOperations.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimError, false, 6, row, 1);

                    }
                }
                //
                //  attachment
                //
                for (int j = 0; j < sortedDict.Count() - 1; j++)
                {
                    KeyValuePair<string, int> lastRow = sortedDict.ElementAt(j + 1);
                    KeyValuePair<string, int> entry = sortedDict.ElementAt(j);
                    currentRow = entry.Value + 1;
                    try
                    {
                        if (entry.Key == "attachments")
                        {
                            bool firstLine = true;
                            List<string> l3 = new List<string>();
                            Classbatim.Attachment att = new Classbatim.Attachment();
                            do
                            {                                
                                temp = new List<string>(slExcelData.DataRows[currentRow]);
                                temp = ClassUtils.reverseOrder(temp);
                                List<string> l4;
                                if (firstLine)
                                {
                                    l3 = ClassbatimUtils0.ParseAttachments(temp);
                                    l4 = new List<string>(l3.Count);
                                    firstLine = false;
                                    currentRow++;
                                }
                                else
                                {
                                    if (temp[0] == "סימון")
                                    {
                                        currentRow++;
                                    }
                                    else
                                    {
                                        l4 = ClassbatimUtils0.ParseAttachmentsValues(temp, l3);
                                        att.mark.Add(l4[0]);
                                        att.color.Add(l4[1]);
                                        att.description.Add(l4[2]);
                                        att.commonWith.Add(l4[3]);
                                        att.area.Add(l4[4]);
                                        currentRow++;
                                    }
                                }                                
                            } while (currentRow < lastRow.Value);
                            batim.tatHelkot[i].attachments.Add(att);
                        }

                    }
                    catch (Exception e)
                    { 
                    }
                }
            }
        }
     }
}
