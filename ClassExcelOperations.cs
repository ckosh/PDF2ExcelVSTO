using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
//using Microsoft.Office.Tools.Excel;

namespace PDF2ExcelVsto
{
    class ClassExcelOperations
    {
        public Excel.Application xlApp;
        public Workbook xlWorkBook;
        public Worksheet xlFilesWorkSheet;
        public Worksheet xlOwnersSheet;
        public Worksheet xlRemarksSheet;
        public Worksheet xlLeasingSheet;
        public Worksheet xlZikotSheet;
        public Worksheet xlMortgageSheet;
        public Worksheet xlPropertySheet;
        public Worksheet xlBatimPropertySheet;
        public Worksheet xlBatimOwnersSheet;
        public Worksheet xlBatimErrorSheet;
        public Worksheet xlBatimLeasingSheet;
        public Worksheet xlBatimMortgageSheet;
        public Worksheet xlBatimRemarksSheet;
        public Worksheet xlBatimAttachmentsSheet;
        public Worksheet xlJoinSplitSheet;

        public dynamic excelSheet;
        public string PDFFolder;
        private bool BatchMode;

        public enum Sheets
        {
            BatimError,
            Owner,
            Remark,
            Mortgage,
            Leasing,
            Zikot,
            PDFfiles,
            Property,
            BatimProperty,
            BatimOwners,
            BatimLeasing,
            BatimMortgage,
            BatimRemarks,
            BatimAttachments,
            JoinSplit
        }

        public ClassExcelOperations(bool bmode)
        {
            if ( bmode )
            {
                xlApp = new Excel.Application();
                xlApp.Visible = true;
                xlApp.Workbooks.Add();
            }
            else
            {
                do
                {
                    try
                    {
                        xlApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (Exception e)
                    {
                        Excel.Application xlApp0 = new Excel.Application();
                        xlApp0.Visible = true;
                        xlApp0.Workbooks.Add();
                        xlApp0.Workbooks.Close();
                        xlApp0.Quit();
                    }

                } while (xlApp == null);
            }

            BatchMode = bmode;
            xlApp.ErrorCheckingOptions.BackgroundChecking = false;
            xlApp.DisplayAlerts = false;
            xlWorkBook = (Excel.Workbook) xlApp.ActiveWorkbook;
            if (xlWorkBook == null)
            {
                xlWorkBook = xlApp.Workbooks.Add();
            }
            else
            {
//                deleteAllPages();
//                deleteAllSheets();

            }
//            xlFilesWorkSheet = xlWorkBook.Worksheets.Item[1];
            excelSheet = xlApp.ActiveSheet as Excel.Worksheet  ;

            //            xlApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            //            xlWorkBook = xlApp.ActiveWorkbook;
            //            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimLeasing);
            try
            {
                DeleteSheetByName("שגיאות");
            }
            catch (Exception e)
            {

            }
            createSheet(Sheets.BatimError, "שגיאות", Color.Red);
            BuildBatimErrorHeader();
        }
        public void DeleteSheetByName(string name)
        {
            xlApp.DisplayAlerts = false;
            for ( int i = xlApp.ActiveWorkbook.Worksheets.Count; i > 0; i--)
            {
                Worksheet wkSheet = (Worksheet)xlApp.ActiveWorkbook.Worksheets[i];
                if (wkSheet.Name == name)
                {
                    wkSheet.Delete();
                }
            }
            xlApp.DisplayAlerts = true;
        }
        public void CreateTitle(int row, int column0, int column2,  string title, double width)
        {
            string lRow = ClassUtils.ColumnLabel(row);
            string lColumn = ClassUtils.ColumnLabel(column0);
            string totalCellName = lColumn + ":" + lColumn;
            xlFilesWorkSheet.Columns[totalCellName].ColumnWidth = width;
            xlFilesWorkSheet.Cells[row, column0].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xlFilesWorkSheet.Cells[row, column0].Value = title;
            
        }
        public void ListPdfFiles(string[]  files)
        {
            CreateTitle(1, 1, 1, "שם נסח", 40.0);
            CreateTitle(1, 2, 2, "גוש", 10.0);
            CreateTitle(1, 3, 3, "חלקה", 10.0);
            CreateTitle(1, 4, 4, "סוג נסח", 15.0);
            int startRow = 2;
            PDFFolder = Path.GetDirectoryName(files[0]);
            for ( int i = 0; i < files.Length; i++)
            {
                string result = Path.GetFileName(files[i]);
                xlFilesWorkSheet.Cells[i+startRow, 1].Value = result;
            }
        }

        public List<string> getPdfFileNames()
        {
            List<string> PDFfiles = new List<string>();
            excelSheet = xlWorkBook.Worksheets["נסחים"];
            int last = xlFilesWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            for (int i = 2; i < last + 1; i++)
            {
                PDFfiles.Add(excelSheet.Cells[i, 1].Value);
            }
            return PDFfiles;
        }

        public void putParamsToTable(int row, string nesachType, string gush, string helka)
        {
            xlFilesWorkSheet.Cells[row, 2].Value = gush;
            xlFilesWorkSheet.Cells[row, 3].Value = helka;
            xlFilesWorkSheet.Cells[row, 4].Value = nesachType;
        }

        public void deleteSheet(Sheets sn)
        {
            switch (sn)
            {
                case Sheets.Owner:
                    if (xlOwnersSheet != null)
                    {
                        xlOwnersSheet.Delete();
                    }
                    break;
                case Sheets.Leasing:
                    if (xlLeasingSheet != null)
                    {
                        xlLeasingSheet.Delete();
                    }
                    break;
                case Sheets.Mortgage:
                    if (xlMortgageSheet != null)
                    {
                        xlMortgageSheet.Delete();
                    }
                    break;
                case Sheets.Property:
                    if (xlPropertySheet != null)
                    {
                        xlPropertySheet.Delete();
                    }
                    break;
                case Sheets.Remark:
                    if (xlRemarksSheet != null)
                    {
                        xlRemarksSheet.Delete();
                    }
                    break;
                case Sheets.Zikot:
                    if (xlZikotSheet != null)
                    {
                        xlZikotSheet.Delete();
                    }
                    break;
                case Sheets.BatimProperty:
                    if (xlBatimPropertySheet != null)
                    {
                        xlBatimPropertySheet.Delete();
                    }
                    break;
                case Sheets.BatimOwners:
                    if ( xlBatimOwnersSheet != null)
                    {
                        xlBatimOwnersSheet.Delete();
                    }
                    break;
                case Sheets.BatimError:
                    if (xlBatimErrorSheet != null)
                    {
                        xlBatimErrorSheet.Delete();
                    }
                    break;
                case Sheets.BatimLeasing:
                    if (xlBatimLeasingSheet != null)
                    {
                        xlBatimLeasingSheet.Delete();
                    }
                    break;
                case Sheets.BatimMortgage:
                    if (xlBatimMortgageSheet != null)
                    {
                        xlBatimMortgageSheet.Delete();
                    }
                    break;
                case Sheets.BatimRemarks:
                    if (xlBatimRemarksSheet != null)
                    {
                        xlBatimRemarksSheet.Delete();
                    }
                    break;
                case Sheets.PDFfiles:
                    if (xlFilesWorkSheet != null)
                    {
                        xlFilesWorkSheet.Delete();
                    }
                    break;
                case Sheets.BatimAttachments:
                    if (xlBatimAttachmentsSheet != null)
                    {
                        xlBatimAttachmentsSheet.Delete();
                    }
                    break;
                case Sheets.JoinSplit:
                    if ( xlJoinSplitSheet != null)
                    {
                        xlJoinSplitSheet.Delete();
                    }
                    break;
            }
        }
        public void createSheet(Sheets sn, string name, Color col)
        {
            xlApp.DisplayAlerts = false;
            DeleteSheetByName(name);
            switch (sn) {
                case Sheets.Owner:
                    xlOwnersSheet = xlWorkBook.Worksheets.Add();
                    xlOwnersSheet.Name = name;
                    xlOwnersSheet.Tab.Color = col;
                    xlOwnersSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.Leasing:
                    xlLeasingSheet = xlWorkBook.Worksheets.Add();
                    xlLeasingSheet.Name = name;
                    xlLeasingSheet.Tab.Color = col;
                    xlLeasingSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.Mortgage:
                    xlMortgageSheet = xlWorkBook.Worksheets.Add();
                    xlMortgageSheet.Name = name;
                    xlMortgageSheet.Tab.Color = col;
                    xlMortgageSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.PDFfiles:
                    xlFilesWorkSheet = xlWorkBook.Worksheets.Add();
                    xlFilesWorkSheet.Name = name;
                    xlFilesWorkSheet.Tab.Color = col;
                    xlFilesWorkSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.Remark:
                    xlRemarksSheet = xlWorkBook.Worksheets.Add();
                    xlRemarksSheet.Name = name;
                    xlRemarksSheet.Tab.Color = col;
                    xlRemarksSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.Zikot:
                    xlZikotSheet = xlWorkBook.Worksheets.Add();
                    xlZikotSheet.Name = name;
                    xlZikotSheet.Tab.Color = col;
                    xlZikotSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.Property:
                    xlPropertySheet = xlWorkBook.Worksheets.Add();
                    xlPropertySheet.Name = name;
                    xlPropertySheet.Tab.Color = col;
                    xlPropertySheet.DisplayRightToLeft = true;
                    break;
                case Sheets.BatimProperty:
                    xlBatimPropertySheet = xlWorkBook.Worksheets.Add();
                    xlBatimPropertySheet.Name = name;
                    xlBatimPropertySheet.Tab.Color = col;
                    xlBatimPropertySheet.DisplayRightToLeft = true;
                    break;
                case Sheets.BatimOwners:
                    xlBatimOwnersSheet = xlWorkBook.Worksheets.Add();
                    xlBatimOwnersSheet.Name = name;
                    xlBatimOwnersSheet.Tab.Color = col;
                    xlBatimOwnersSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.BatimLeasing:
                    xlBatimLeasingSheet = xlWorkBook.Worksheets.Add();
                    xlBatimLeasingSheet.Name = name;
                    xlBatimLeasingSheet.Tab.Color = col;
                    xlBatimLeasingSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.BatimError:
                    xlBatimErrorSheet = xlWorkBook.Worksheets.Add();
                    xlBatimErrorSheet.Name = name;
                    xlBatimErrorSheet.Tab.Color = col;
                    xlBatimErrorSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.BatimMortgage:
                    xlBatimMortgageSheet = xlWorkBook.Worksheets.Add();
                    xlBatimMortgageSheet.Name = name;
                    xlBatimMortgageSheet.Tab.Color = col;
                    xlBatimMortgageSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.BatimRemarks:
                    xlBatimRemarksSheet = xlWorkBook.Worksheets.Add();
                    xlBatimRemarksSheet.Name = name;
                    xlBatimRemarksSheet.Tab.Color = col;
                    xlBatimRemarksSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.BatimAttachments:
                    xlBatimAttachmentsSheet = xlWorkBook.Worksheets.Add();
                    xlBatimAttachmentsSheet.Name = name;
                    xlBatimAttachmentsSheet.Tab.Color = col;
                    xlBatimAttachmentsSheet.DisplayRightToLeft = true;
                    break;
                case Sheets.JoinSplit:
                    xlJoinSplitSheet = xlWorkBook.Worksheets.Add();
                    xlJoinSplitSheet.Name = name;
                    xlJoinSplitSheet.Tab.Color = col;
                    xlJoinSplitSheet.DisplayRightToLeft = true;
                    break;
            }
            xlApp.DisplayAlerts = true;
        }

        public void setActiveSheet(Sheets sn)
        {
            switch (sn)
            {
                case Sheets.Owner:
                    xlOwnersSheet.Select();
                    break;
                case Sheets.Leasing:
                    xlLeasingSheet.Select();
                    break;
                case Sheets.Mortgage:
                    xlMortgageSheet.Select();
                    break;
                case Sheets.PDFfiles:
                    xlFilesWorkSheet.Select();
                    break;
                case Sheets.Remark:
                    xlRemarksSheet.Select();
                    break;
                case Sheets.Zikot:
                    xlZikotSheet.Select();
                    break;
                case Sheets.BatimProperty:
                     xlBatimPropertySheet.Select();
                    break;
                case Sheets.BatimOwners:
                    xlBatimOwnersSheet.Select();
                    break;
                case Sheets.BatimLeasing:
                    xlBatimLeasingSheet.Select();
                    break;
                case Sheets.BatimError:
                    xlBatimErrorSheet.Select();
                    break;
                case Sheets.BatimMortgage:
                    xlBatimMortgageSheet.Select();
                    break;
                case Sheets.BatimRemarks:
                    xlBatimRemarksSheet.Select();
                    break;
                case Sheets.BatimAttachments:
                    xlBatimAttachmentsSheet.Select();
                    break;
                case Sheets.JoinSplit:
                    xlJoinSplitSheet.Select();
                    break;
            }
        }

        public void refreshAll()
        {
            xlWorkBook.RefreshAll();
        }
        public Excel.Worksheet getSheet(Sheets sn)
        {
            Excel.Worksheet retSheet = null;
            switch (sn)
            {
                case Sheets.Owner:
                    retSheet = xlOwnersSheet ;
                   break;
                case Sheets.Leasing:
                    retSheet = xlLeasingSheet;
                    break;
                case Sheets.Mortgage:
                    retSheet = xlMortgageSheet;
                    break;
                case Sheets.PDFfiles:
                    retSheet = xlFilesWorkSheet;
                    break;
                case Sheets.Remark:
                    retSheet = xlRemarksSheet;
                    break;
                case Sheets.Zikot:
                    retSheet = xlZikotSheet;
                    break;
                case Sheets.Property:
                    retSheet = xlPropertySheet;
                    break;
                case Sheets.BatimProperty:
                    retSheet = xlBatimPropertySheet;
                    break;
                case Sheets.BatimOwners:
                    retSheet = xlBatimOwnersSheet;
                    break;
                case Sheets.BatimLeasing:
                    retSheet = xlBatimLeasingSheet;
                    break;
                case Sheets.BatimError:
                    retSheet = xlBatimErrorSheet;
                    break;
                case Sheets.BatimMortgage:
                    retSheet = xlBatimMortgageSheet;
                    break;
                case Sheets.BatimRemarks:
                    retSheet = xlBatimRemarksSheet;
                    break;
                case Sheets.BatimAttachments:
                    retSheet = xlBatimAttachmentsSheet;
                    break;
                case Sheets.JoinSplit:
                    retSheet = xlJoinSplitSheet;
                    break;
            }
            return retSheet;
        }
        public int BuildBatimErrorHeader()
        {
            int rowNumber = 1;
            HeadTitle(xlBatimErrorSheet, "קובץ", 1, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimErrorSheet, "פסקה", 1, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimErrorSheet, "תת חלקה", 1, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimErrorSheet, "פסקה", 1, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimErrorSheet, "הערות", 1, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimErrorSheet, "2", 1, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);

            return rowNumber;
        }
        public int getBatimErrorLine()
        {
            int ret = 0;
            try
            {
                var cellValue = (xlBatimErrorSheet.Cells[1, 6] as Excel.Range).Value;
                ret = Convert.ToInt32(cellValue);
                return ret;
            }
            catch(Exception e)
            { 
                MessageBox.Show(e.ToString());
            }
            return ret;
        }
        public int BuildJoinSplitHeader()
        {
            int rowNumber = 3;
            System.Drawing.Color PattensBlue = GetFromRGB(0xDA, 0xEE, 0xF3);
            string NewShekel = "\u20AA";
            string test1 = "שווי הזכויות במצב הנכנס (" + NewShekel +")";

            HeadTitle(xlJoinSplitSheet, "פינוי בינוי", 1, 1, 17, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 13, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "נתוני המקרקעין", 2, 1, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 113, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "מצב נכנס", 2, 7, 17, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 13, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "ספירה", 3, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "גוש", 3, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "חלקה", 3, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "שטח החלקה הרשום (במ\"ר)", 3, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "ייעוד החלקה", 3, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "שטח החלקה הכלול באיחוד וחלוקה", 3, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "מס' תת חלקה", 3, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "שם הבעלים / חוכר הרשום", 3, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "ת.ז / ח.פ", 3, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "חלק הבעלים בנכס", 3, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "שטח תת חלקה רשום (במ\"ר)", 3, 11, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "חלק ברכוש המשותף (בשבר)", 3, 12, 12, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "חלק ברכוש המשותף (ב%)", 3, 13, 13, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "שווי הזכויות במצב הנכנס (" + NewShekel + ")", 3, 14, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "שווי תרומת המבנים (" + NewShekel + ")", 3, 15, 15, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "שווי זכויות + מחוברים (" + NewShekel + ")", 3, 16, 16, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "שווי יחסי (באחוזים)", 3, 17, 17, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 61, true, XlBorderWeight.xlThin);


            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(1), Type.Missing]).ColumnWidth = 6.82;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(2), Type.Missing]).ColumnWidth = 4.55;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(3), Type.Missing]).ColumnWidth = 7.0;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(4), Type.Missing]).ColumnWidth = 9.55;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(5), Type.Missing]).ColumnWidth = 8.0;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(6), Type.Missing]).ColumnWidth = 10.0;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(7), Type.Missing]).ColumnWidth = 8.82;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(8), Type.Missing]).ColumnWidth = 18.0;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(9), Type.Missing]).ColumnWidth = 13.36;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(10), Type.Missing]).ColumnWidth = 8.73;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(11), Type.Missing]).ColumnWidth = 7.36;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(12), Type.Missing]).ColumnWidth = 7.27;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(13), Type.Missing]).ColumnWidth = 7.36;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(14), Type.Missing]).ColumnWidth = 12.73;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(15), Type.Missing]).ColumnWidth = 11.27;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(16), Type.Missing]).ColumnWidth = 12.55;
            ((Microsoft.Office.Interop.Excel.Range)xlJoinSplitSheet.Columns[ClassUtils.ColumnLabel(17), Type.Missing]).ColumnWidth = 7.36;

            HeadTitle(xlJoinSplitSheet, "", 4, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "a", 4, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "b", 4, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "c", 4, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "", 4, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "", 4, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "d", 4, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "e", 4, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "f", 4, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "g", 4, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "h", 4, 11, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "i", 4, 12, 12, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "j", 4, 13, 13, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "", 4, 14, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "", 4, 15, 15, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "", 4, 16, 16, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            HeadTitle(xlJoinSplitSheet, "", 4, 17, 17, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 10, true, PattensBlue, 15, true, XlBorderWeight.xlThin);
            rowNumber = 5;
            return rowNumber;
        }
        public int BuildBatimRemarksHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlBatimRemarksSheet, "הערות - בתים משותפים", 1, 1, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);

            HeadTitle(xlBatimRemarksSheet, "גוש", 2, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimRemarksSheet, "חלקה", 2, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimRemarksSheet, "תת חלקה", 2, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimRemarksSheet, "סוג הערה", 2, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimRemarksSheet, "שם", 2, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimRemarksSheet, "סוג זיהוי", 2, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimRemarksSheet, "מס. זיהוי", 2, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimRemarksSheet, "חלק", 2, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimRemarksSheet, "שטר", 2, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimRemarksSheet, "הערה", 2, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimRemarksSheet, "קובץ", 2, 11, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            return rowNumber;
        }
        public int BuildBatimMortgageHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlBatimMortgageSheet, "משכנתאות - בתים משותפים", 1, 1, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);

            HeadTitle(xlBatimMortgageSheet, "גוש", 2, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimMortgageSheet, "חלקה", 2, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimMortgageSheet, "תת חלקה", 2, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimMortgageSheet, "משכנתה", 2, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimMortgageSheet, "ממשכן", 2, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimMortgageSheet, "סוג זיהוי", 2, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimMortgageSheet, "מס. זיהוי", 2, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimMortgageSheet, "חלק", 2, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimMortgageSheet, "שטר", 2, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimMortgageSheet, "דרגה", 2, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimMortgageSheet, "קובץ", 2, 11, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            return rowNumber;
        }
        public int BuildBatimAttachmentsHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlBatimAttachmentsSheet, "הצמדות - בתים משותפים", 1, 1, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimAttachmentsSheet, "גוש", 2, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimAttachmentsSheet, "חלקה", 2, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimAttachmentsSheet, "תת חלקה", 2, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimAttachmentsSheet, "סימון בתשריט", 2, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimAttachmentsSheet, "צבע בתשריט", 2, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimAttachmentsSheet, "תיאור הצמדה", 2, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimAttachmentsSheet, "משותפת ל", 2, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimAttachmentsSheet, "שטח במ\"ר", 2, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimAttachmentsSheet, "קובץ", 2, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            return rowNumber;
        }
        public int BuildBatimLeasingHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlBatimLeasingSheet, "חכירות - בתים משותפים", 1, 1, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "גוש", 2, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "חלקה", 2, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "תת חלקה", 2, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "חכירה", 2, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "שם", 2, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "סוג זיהוי", 2, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "מס. זיהוי", 2, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "חלק בנכס", 2, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "שטר", 2, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "רמה", 2, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "תאריך סיום", 2, 11, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "הערות", 2, 12, 12, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "חלק בנכס", 2, 13, 13, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimLeasingSheet, "קובץ", 2, 14, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);

            return rowNumber;
        }
        public int BuildBatimOwnerHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlBatimOwnersSheet, "בעלים - בתים משותפים", 1, 1, 18, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "לחץ על סימון להגיע לדף הנתונים", 1, 19, 23, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "", 1, 24, 24, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);

            HeadTitle(xlBatimOwnersSheet, "גוש", 2, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "חלקה", 2, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
             HeadTitle(xlBatimOwnersSheet, "שטח חלקה במ\"ר", 2, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "תת חלקה", 2, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "שטח במ\"ר", 2, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "תיאור קומה", 2, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "כניסה", 2, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "אגף", 2, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "מבנה", 2, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "החלק ברכוש המשותף", 2, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "החלק באחוזים", 2, 11, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);


            HeadTitle(xlBatimOwnersSheet, "קניין", 2, 12, 12, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "שם", 2, 13, 13, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "סוג זיהוי", 2, 14, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "מס. זיהוי", 2, 15, 15, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "החלק בתת חלקה", 2, 16, 16, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "החלק באחוזים", 2, 17, 17, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);


            HeadTitle(xlBatimOwnersSheet, "שטר", 2, 18, 18, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "משכנתאות", 2, 19, 19, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "הערות", 2, 20, 20, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "חכירות", 2, 21, 21, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "הצמדות", 2, 22, 22, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "זיקות הנאה", 2, 23, 23, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimOwnersSheet, "קובץ", 2, 24, 24, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            return rowNumber;
        }
        public int BuildBatimPropertyHeader()
        {
            int rowNumber = 4;
            HeadTitle(xlBatimPropertySheet, "רכוש משותף - בתים משותפים", 1, 1, 19, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);

            HeadTitle(xlBatimPropertySheet, "גו\"ח", 2, 1, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet,"הנכס נוצר", 2, 3, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "הרכוש המשותף", 2, 6, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "הערות - זיקות הנאה", 2, 15, 16, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "גירסת נסח", 2, 17, 19, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "גוש", 3, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "חלקה", 3, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "שטר", 3, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "מיום", 3, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "סוג השטר", 3, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "רשויות", 3, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "שטח במ\"ר", 3, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "תת חלקות", 3, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "תקנון", 3, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "שטר יוצר", 3, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "תיק יוצר", 3, 11, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "תיק בית משותף", 3, 12, 12, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "כתובת", 3, 13, 13, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "הערות", 3, 14, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "זיקות הנאה", 3, 15, 15, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "הערות", 3, 16, 16, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "תאריך", 3, 17, 17, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "מס. נסח", 3, 18, 18, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);
            HeadTitle(xlBatimPropertySheet, "קובץ", 3, 19, 19, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 20, true, XlBorderWeight.xlMedium);

            return rowNumber;
        }
        public int BuildMortgageHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlMortgageSheet, "משכנתאות - פנקס הזכויות", 1, 1, 18, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);
            HeadTitle(xlMortgageSheet, "גוש", 2, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "חלקה", 2, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "מס. שטר", 2, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "תאריך", 2, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "מהות פעולה", 2, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "בעלי משכנתה", 2, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "סוג זיהוי", 2, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "מס' זיהוי", 2, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "שם הלווה", 2, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "סוג זיהוי", 2, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "מס' זיהוי", 2, 11, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "דרגה", 2, 12, 12, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "סכום", 2, 13, 13, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "בתנאי שטר מקורי", 2, 14, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "החלק בנכס", 2, 15, 15, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "החלק בשבר", 2, 16, 16, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "הערות", 2, 17, 17, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlMortgageSheet, "שם קובץ", 2, 18, 18, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);

            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(1), Type.Missing]).ColumnWidth = 5.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(2), Type.Missing]).ColumnWidth = 4.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(3), Type.Missing]).ColumnWidth = 10.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(4), Type.Missing]).ColumnWidth = 9.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(5), Type.Missing]).ColumnWidth = 6.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(6), Type.Missing]).ColumnWidth = 20.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(7), Type.Missing]).ColumnWidth = 5.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(8), Type.Missing]).ColumnWidth = 9.0;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(9), Type.Missing]).ColumnWidth = 10.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(10), Type.Missing]).ColumnWidth = 5.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(11), Type.Missing]).ColumnWidth = 6.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(12), Type.Missing]).ColumnWidth = 6.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(13), Type.Missing]).ColumnWidth = 13.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(14), Type.Missing]).ColumnWidth = 10.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(15), Type.Missing]).ColumnWidth = 10.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(16), Type.Missing]).ColumnWidth = 30.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(17), Type.Missing]).ColumnWidth = 15.00;
            ((Microsoft.Office.Interop.Excel.Range)xlMortgageSheet.Columns[ClassUtils.ColumnLabel(18), Type.Missing]).ColumnWidth = 15.00;

            return rowNumber;
        }
        public int BuildRemarkHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlRemarksSheet, "הערות - פנקס הזכויות", 1, 1,10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);
            HeadTitle(xlRemarksSheet, "גוש", 2, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlRemarksSheet, "חלקה", 2, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlRemarksSheet, "מס. שטר", 2, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlRemarksSheet, "תאריך", 2, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlRemarksSheet, "מהות פעולה", 2, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlRemarksSheet, "שם המוטב", 2, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlRemarksSheet, "סוג זיהוי", 2, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlRemarksSheet, "מס. זיהוי", 2, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlRemarksSheet, "הערות", 2, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlRemarksSheet, "שם קובץ", 2, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);

            ((Microsoft.Office.Interop.Excel.Range)xlRemarksSheet.Columns[ClassUtils.ColumnLabel(1), Type.Missing]).ColumnWidth = 6.00;
            ((Microsoft.Office.Interop.Excel.Range)xlRemarksSheet.Columns[ClassUtils.ColumnLabel(2), Type.Missing]).ColumnWidth = 5.00;
            ((Microsoft.Office.Interop.Excel.Range)xlRemarksSheet.Columns[ClassUtils.ColumnLabel(3), Type.Missing]).ColumnWidth = 15.00;
            ((Microsoft.Office.Interop.Excel.Range)xlRemarksSheet.Columns[ClassUtils.ColumnLabel(4), Type.Missing]).ColumnWidth = 12.00;
            ((Microsoft.Office.Interop.Excel.Range)xlRemarksSheet.Columns[ClassUtils.ColumnLabel(5), Type.Missing]).ColumnWidth = 30.00;
            ((Microsoft.Office.Interop.Excel.Range)xlRemarksSheet.Columns[ClassUtils.ColumnLabel(6), Type.Missing]).ColumnWidth = 20.00;
            ((Microsoft.Office.Interop.Excel.Range)xlRemarksSheet.Columns[ClassUtils.ColumnLabel(7), Type.Missing]).ColumnWidth = 6.00;
            ((Microsoft.Office.Interop.Excel.Range)xlRemarksSheet.Columns[ClassUtils.ColumnLabel(8), Type.Missing]).ColumnWidth = 15.0;
            ((Microsoft.Office.Interop.Excel.Range)xlRemarksSheet.Columns[ClassUtils.ColumnLabel(9), Type.Missing]).ColumnWidth = 40.00;
            ((Microsoft.Office.Interop.Excel.Range)xlRemarksSheet.Columns[ClassUtils.ColumnLabel(10), Type.Missing]).ColumnWidth = 10.00;
            return rowNumber;
        }
        public int BuildLeasingHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlLeasingSheet, "חכירות - פנקס הזכויות", 1, 1, 15, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);
            HeadTitle(xlLeasingSheet, "גוש", 2, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "חלקה", 2, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "מס. שטר", 2, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "תאריך", 2, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "מהות הפעולה", 2, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "שם החוכר", 2, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "סוג זיהוי", 2, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "מס. זיהוי", 2, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "החלק בנכס", 2, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "רמת חכירה", 2, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "בתנאי שטר מקורי", 2, 11, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "תאריך סיום", 2, 12, 12, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "החלק בנכס", 2, 13, 13, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "הערות", 2, 14, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlLeasingSheet, "שם קובץ", 2, 15, 15, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);

            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(1), Type.Missing]).ColumnWidth = 6.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(2), Type.Missing]).ColumnWidth = 5.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(3), Type.Missing]).ColumnWidth = 12.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(4), Type.Missing]).ColumnWidth = 12.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(5), Type.Missing]).ColumnWidth = 25.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(6), Type.Missing]).ColumnWidth = 15.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(7), Type.Missing]).ColumnWidth = 4.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(8), Type.Missing]).ColumnWidth = 10.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(9), Type.Missing]).ColumnWidth = 15.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(10), Type.Missing]).ColumnWidth = 18.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(11), Type.Missing]).ColumnWidth = 12.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(12), Type.Missing]).ColumnWidth = 10.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(13), Type.Missing]).ColumnWidth = 8.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(14), Type.Missing]).ColumnWidth = 30.00;
            ((Microsoft.Office.Interop.Excel.Range)xlLeasingSheet.Columns[ClassUtils.ColumnLabel(15), Type.Missing]).ColumnWidth = 10.00;
            return rowNumber;

        }
        public int BuildPropertyHeader()
        {
            int rowNumber = 3;
            HeadTitle(xlPropertySheet, "תאור הנכס - פנקס הזכויות", 1, 1, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);
            HeadTitle(xlPropertySheet, "מס\"ד", 2, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlPropertySheet, "גוש", 2, 2,2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlPropertySheet, "חלקה", 2, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlPropertySheet, "רשויות", 2, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlPropertySheet, "שטח במ\"ר", 2, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlPropertySheet, "סוג המקרקעין", 2, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlPropertySheet, "הערות רשם המקרקעין", 2, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlPropertySheet, "המספרים הישנים של החלקה", 2, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlPropertySheet, "שם הקובץ", 2, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);

            ((Microsoft.Office.Interop.Excel.Range)xlPropertySheet.Columns[ClassUtils.ColumnLabel(1), Type.Missing]).ColumnWidth = 4.00;
            ((Microsoft.Office.Interop.Excel.Range)xlPropertySheet.Columns[ClassUtils.ColumnLabel(2), Type.Missing]).ColumnWidth = 10.00;
            ((Microsoft.Office.Interop.Excel.Range)xlPropertySheet.Columns[ClassUtils.ColumnLabel(3), Type.Missing]).ColumnWidth = 10.00;
            ((Microsoft.Office.Interop.Excel.Range)xlPropertySheet.Columns[ClassUtils.ColumnLabel(4), Type.Missing]).ColumnWidth = 30.00;
            ((Microsoft.Office.Interop.Excel.Range)xlPropertySheet.Columns[ClassUtils.ColumnLabel(5), Type.Missing]).ColumnWidth = 15.00;
            ((Microsoft.Office.Interop.Excel.Range)xlPropertySheet.Columns[ClassUtils.ColumnLabel(6), Type.Missing]).ColumnWidth = 15.00;
            ((Microsoft.Office.Interop.Excel.Range)xlPropertySheet.Columns[ClassUtils.ColumnLabel(7), Type.Missing]).ColumnWidth = 40.00;
            ((Microsoft.Office.Interop.Excel.Range)xlPropertySheet.Columns[ClassUtils.ColumnLabel(8), Type.Missing]).ColumnWidth = 20.00;
            ((Microsoft.Office.Interop.Excel.Range)xlPropertySheet.Columns[ClassUtils.ColumnLabel(9), Type.Missing]).ColumnWidth = 20.00;

            return rowNumber;
        }
        public int buildOwnerHeadr0()
        {
            int rowNumber = 1;
            HeadTitle(xlOwnersSheet, "בעלויות - פנחס הזכויות", 1, 1, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);
            HeadTitle(xlOwnersSheet, "לחץ על הסמן לצפיה", 1, 12, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);
            HeadTitle(xlOwnersSheet, "", 1, 15, 15, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 12, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlMedium);

            HeadTitle(xlOwnersSheet, "מס\"ד", 2, 1, 1, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "גוש", 2, 2, 2, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "חלקה", 2, 3, 3, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "מס' שטר", 2, 4, 4, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "תאריך", 2, 5, 5, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "מהות פעולה", 2, 6, 6, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "הבעלים", 2, 7, 7, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "סוג זיהוי", 2, 8, 8, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "מס' זיהוי", 2, 9, 9, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "החלק בנכס", 2, 10, 10, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "אחוז בנכס", 2, 11, 11, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "משכנתאות", 2, 12, 12, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "חכירות", 2, 13, 13, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "הערות", 2, 14, 14, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);
            HeadTitle(xlOwnersSheet, "שם קובץ", 2, 15, 15, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignCenter, 11, true, System.Drawing.Color.Aqua, 40, true, XlBorderWeight.xlThin);

            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(1), Type.Missing]).ColumnWidth = 4.00;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(2), Type.Missing]).ColumnWidth = 8.00;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(3), Type.Missing]).ColumnWidth = 4.00;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(4), Type.Missing]).ColumnWidth = 14.00;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(5), Type.Missing]).ColumnWidth = 10.00;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(6), Type.Missing]).ColumnWidth = 20.00;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(7), Type.Missing]).ColumnWidth = 15.0;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(8), Type.Missing]).ColumnWidth = 6.00;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(9), Type.Missing]).ColumnWidth = 12.00;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(10), Type.Missing]).ColumnWidth = 20.00;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(11), Type.Missing]).ColumnWidth = 10.00;
            ((Microsoft.Office.Interop.Excel.Range)xlOwnersSheet.Columns[ClassUtils.ColumnLabel(12), Type.Missing]).ColumnWidth = 10.00;
            rowNumber = 3;
            return rowNumber;
        }
        public void addNameRange(Sheets sn, int irow1, int irow2, int icol1, int icol2,int gush, int helka, int tat, string prefix)
        {
            string sss = prefix + "_" + gush.ToString() + "_" + helka.ToString() + "_" + tat.ToString();
            Excel.Worksheet asheet = getSheet(sn);
            Range frame = asheet.Range[asheet.Cells[irow1, icol1], asheet.Cells[irow2, icol2]].Cells;
            frame.Name = sss;
        }
        public void setBoarder(Sheets sn, int irow1, int irow2, int icol1, int icol2, int thickness)
        {
            XlBorderWeight thick; 
            if ( thickness == 0)
            {
                thick = XlBorderWeight.xlThick;
            }
            else if (thickness == 1)
            {
                thick = XlBorderWeight.xlMedium;
            }
            else if (thickness == 2)
            {
                thick = XlBorderWeight.xlThin;
            }
            else if (thickness == 3)
            {
                thick = XlBorderWeight.xlHairline;
            }
            else
            {
                thick = XlBorderWeight.xlMedium;
            }

            Worksheet asheet = getSheet(sn);            
            Range frame = asheet.Range[asheet.Cells[irow1, icol1], asheet.Cells[irow2, icol2]].Cells;
            frame.BorderAround2(Type.Missing, thick, XlColorIndex.xlColorIndexAutomatic, Type.Missing);

//            frame.EntireColumn.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
        }

        public void mergeCells(Sheets asheet, int row1, int col1, int row2, int col2)
        {
            Excel.Worksheet xlsheet = getSheet(asheet);
            Range titleRang = xlsheet.Range[xlsheet.Cells[row1, col1], xlsheet.Cells[row2, col2]].Cells;
            titleRang.Merge(Type.Missing);
            titleRang.Style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }

        public void HeadTitle(Worksheet theSheet, string title, int irow, int icol, int endcol, XlHAlign hAlign, XlVAlign vAlign, int fonSize, bool bBolt, System.Drawing.Color titleColor, int rowHeight , bool boarder, XlBorderWeight weight )

        {
            theSheet.Cells[irow, icol] = title;
            theSheet.Cells[irow, icol].Style.WrapText = true;
            Range titleRang = theSheet.Range[theSheet.Cells[irow, icol], theSheet.Cells[irow, endcol]].Cells;
            titleRang.Merge(Type.Missing);
            titleRang.HorizontalAlignment = hAlign;
            titleRang.VerticalAlignment = vAlign;
            titleRang.Font.Size = fonSize;
            titleRang.Font.Bold = bBolt;
            titleRang.Font.Name = "Tahoma";
            titleRang.Interior.Color = titleColor;
            titleRang.RowHeight = rowHeight;
            titleRang.EntireColumn.NumberFormat = "@";

            if (boarder)
            {
                titleRang.BorderAround(XlLineStyle.xlContinuous,weight,XlColorIndex.xlColorIndexAutomatic,XlColorIndex.xlColorIndexAutomatic);
            }
        }

        public void paintRow(Sheets sn, int irow, int icol, int endcol, System.Drawing.Color Color)
        {
            Worksheet asheet = getSheet(sn);
            Range Rang = asheet.Range[asheet.Cells[irow, icol], asheet.Cells[irow, endcol]].Cells;
            Rang.Interior.Color = Color;
        }

        public void PutValueInSheetRowColumn(Sheets sn, int row, int column, string  val)
        {
            Worksheet asheet = getSheet(sn);
            if (!(val is null)) 
            {
                asheet.Cells[row, column] = val;
            }   
        }

        public void createHyperLink(Sheets sn, int row, int column, string gush, string helka, string tat, string prefix)
        {
            string subAddress = prefix + "_" + gush + "_" + helka + "_" + tat;
            Worksheet asheet = getSheet(sn);
            Object Anchor = asheet.Cells[row,column];
            Object TextToDisplay = "X";
            asheet.Hyperlinks.Add(Anchor, "", subAddress, "", "X");
        }
        public void setSheetCellWrapText(Sheets sn, bool onoff, int columns , int rows , int rowtofreez)
        {
            Worksheet asheet = getSheet(sn);
            Range rrr = asheet.Cells[rows, columns];
            rrr.Select();
            rrr.Style.WrapText = onoff;
            asheet.Columns.AutoFit();

            Excel.Range activeCell = asheet.Cells[1, 1];
            activeCell.Select();

            asheet.Application.ActiveWindow.SplitRow = rowtofreez;
            asheet.Application.ActiveWindow.FreezePanes = true;
            refreshAll();
        }
        public void SaveResultExcel(string location)
        {
            xlApp.ActiveWorkbook.SaveAs(location);
            xlApp.ActiveWorkbook.Close();
            if ( BatchMode)
            {
                Marshal.ReleaseComObject(xlApp);
                xlApp.Quit();
            }
            else
            {               
                xlApp.Workbooks.Add();
            }
        }
        
        public void CorrectFormatForSum(Sheets sn, int columns, int rows1, int rows2, string numFormat)
        {
            Worksheet asheet = getSheet(sn);
            var startCell = (Range)asheet.Cells[rows1, columns];
            var endCell = (Range)asheet.Cells[rows2, columns];
            Range rrr = asheet.Range[startCell, endCell];
            rrr.Select();
            Range oTargetRange = asheet.Range[startCell, endCell];
            try
            {
                rrr.TextToColumns(oTargetRange, XlTextParsingType.xlDelimited,
                    XlTextQualifier.xlTextQualifierDoubleQuote, false, true, false, false, false, true, "-");
            }
            catch (Exception ex)
            {

            }
            rrr.NumberFormat = numFormat;

        }
        public Color GetFromRGB(byte r, byte g, byte b)
        {
            Color RGBColor = new Color();
            RGBColor = Color.FromArgb(r, g, b);
            return RGBColor;
        }

        public  class PutCellParameters
        {
            public bool ifmerge { get; set; }
            public int Rowextension { get; set; }
            public int Columnextension { get; set; }
            public Color colorbackground { get; set; }
            public Excel.XlVAlign xlVAlign { get; set; }
            public Excel.XlHAlign xlHAlign { get; set; }
            public bool ifFrame { get; set; }
            public XlBorderWeight Weight { get; set; }
            public int fontSize { get; set; }

        }

        public void putValueWithParameter(Sheets sheet, string value, int row, int col, PutCellParameters param)
        {
            Worksheet asheet = getSheet(sheet);
            Range sellection;
            if (param.ifmerge)
            {
                sellection = asheet.Range[asheet.Cells[row, col], asheet.Cells[row + param.Rowextension-1 , col + param.Columnextension-1]].Cells;
                sellection.Merge(Type.Missing);
            }
            else
            {
                sellection = asheet.Range[asheet.Cells[row, col], asheet.Cells[row, col]].Cells;
            }
            sellection.VerticalAlignment = param.xlVAlign;
            sellection.HorizontalAlignment = param.xlHAlign;
            sellection.Interior.Color = param.colorbackground;
            sellection.BorderAround2(Type.Missing, param.Weight, XlColorIndex.xlColorIndexAutomatic, Type.Missing);
            sellection.Font.Size = param.fontSize;
            asheet.Cells[row, col] = value;
            sellection.Select();
        }

        public void deleteAllPages()
        {
            xlApp.DisplayAlerts = false;
            deleteSheet(ClassExcelOperations.Sheets.BatimAttachments);
            deleteSheet(ClassExcelOperations.Sheets.BatimError);
            deleteSheet(ClassExcelOperations.Sheets.BatimLeasing);
            deleteSheet(ClassExcelOperations.Sheets.BatimMortgage);
            deleteSheet(ClassExcelOperations.Sheets.BatimOwners);
            deleteSheet(ClassExcelOperations.Sheets.BatimProperty);
            deleteSheet(ClassExcelOperations.Sheets.BatimRemarks);
            deleteSheet(ClassExcelOperations.Sheets.JoinSplit);
            deleteSheet(ClassExcelOperations.Sheets.Leasing);
            deleteSheet(ClassExcelOperations.Sheets.Mortgage);
            deleteSheet(ClassExcelOperations.Sheets.Owner);
            deleteSheet(ClassExcelOperations.Sheets.PDFfiles);
            deleteSheet(ClassExcelOperations.Sheets.Property);
            deleteSheet(ClassExcelOperations.Sheets.Remark);
            deleteSheet(ClassExcelOperations.Sheets.Zikot);
        }
    }
}
