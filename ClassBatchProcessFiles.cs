using PDF2ExcelVsto.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF2ExcelVsto
{
    class ClassBatchProcessFiles
    {
        ClassExcelOperations excelOperations  ;
        ClassFilesHandle fileHandler;
        ClasszhuiotManager zuiotManager;
        ClassbatimManager0 batimManager;
        bool DebugMode;
        int customerType;

        public ClassBatchProcessFiles(string[] sarray, bool debugMode, string tempFolder, bool batchMode, int custType)
        {
            DebugMode = debugMode;
            customerType = custType;
            if (excelOperations != null)
            {
                excelOperations.deleteAllPages();
            }
            excelOperations = new ClassExcelOperations(batchMode);
            fileHandler = new ClassFilesHandle(excelOperations, DebugMode, tempFolder);
            zuiotManager = new ClasszhuiotManager(fileHandler, excelOperations);
            batimManager = new ClassbatimManager0(fileHandler, excelOperations);
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.PDFfiles);
            excelOperations.createSheet(ClassExcelOperations.Sheets.PDFfiles, "נסחים", Color.Black);
            fileHandler.clearCSVFiles("batim");
            fileHandler.clearCSVFiles("zhuiot");
            excelOperations.ListPdfFiles(sarray);
        }

        public string  convert()
        {
            string resultfile;
            fileHandler.convertPDF2CSV();
            batimManager.convertBatimtoExcel();
            zuiotManager.convertZhuiottoExcel();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimProperty);
            batimManager.CreatePropertyTable();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimLeasing);
            batimManager.CreateBatimLeasing();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimMortgage);
            batimManager.CreateBatimMortgage();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimRemarks);
            batimManager.CreateBatimRemarksTables();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimAttachments);
            batimManager.createBatimAttachments();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimOwners);
            batimManager.CreateBatimOwnTable();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Property);
            zuiotManager.CreatePropertyTables();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Leasing);
            zuiotManager.CreateLeasingTables();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Mortgage);
            zuiotManager.CreateMortGageTables();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Remark);
            zuiotManager.CreateRemarksTables();

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Owner);
            zuiotManager.CreateOwnersTable();
            //
            //  
            //
            ClassJoinSplitManager joinSplitManager;
            joinSplitManager = new ClassJoinSplitManager(fileHandler, excelOperations, batimManager, zuiotManager);
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.JoinSplit);
            joinSplitManager.CreateJoinSplitTable();

            //if (customerType > 80)
            //{
            //    ClassJoinSplitManager joinSplitManager;
            //    joinSplitManager = new ClassJoinSplitManager(fileHandler, excelOperations, batimManager, zuiotManager);
            //    excelOperations.deleteSheet(ClassExcelOperations.Sheets.JoinSplit);
            //    joinSplitManager.CreateJoinSplitTable();
            //}

            resultfile = fileHandler.PDFfolder + "\\Tabu_results_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            excelOperations.SaveResultExcel(resultfile);
            resultfile = resultfile + ".xlsx";
            return resultfile;
        }

        public List<int> getTotalNumberOfOwners()
        {
            List<int> numberOfOwners = new List<int>();

 
            if ( batimManager.allBatim.Count > 0)
            {
                foreach (Classbatim batim in batimManager.allBatim)
                {
                    int ret = 0;
                    for ( int i = 0; i < batim.tatHelkot.Count; i++)
                    {
                        ret = ret + batim.tatHelkot[i].owners.Count;
                    }
                    numberOfOwners.Add(ret);
                }
            }
            if ( zuiotManager.allTaboo.Count > 0)
            {
                if (zuiotManager.allTaboo.Count > 0)
                {
                    foreach (ClassTaboo taboo in zuiotManager.allTaboo)
                    {
                        numberOfOwners.Add(taboo.zhuiotOwners.Count);
                    }
                }
            }
            return numberOfOwners;
        }
    }
}
