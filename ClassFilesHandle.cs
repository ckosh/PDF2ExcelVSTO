using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDF2ExcelVsto
{
    class ClassFilesHandle
    {
        private List<string> zhuiotCSV = new List<string>();
        private List<string> batimCSV = new List<string>();

        ClassExcelOperations excelOperation;
        public string PDFfolder;
        public bool DebugMode;
        public ClassFilesHandle(ClassExcelOperations excelop, bool debugMode, string tempFolder)
        {
            excelOperation = excelop;
            DebugMode = debugMode;
            PDFfolder = tempFolder;
        }
        public void collectPDFFiles()
        {
            System.Windows.Forms.OpenFileDialog openFileDialog1;
            openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Title = "בחר קובצי נסח טאבו";
            openFileDialog1.Filter = "PDF files (*.PDF)|*.pdf";
            openFileDialog1.Multiselect = true;
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                PDFfolder= Path.GetDirectoryName(openFileDialog1.FileNames[0]); 
                excelOperation.ListPdfFiles(openFileDialog1.FileNames);
            }
        }
        public void clearCSVFiles(string pdfType)
        {
            if (pdfType == "zhuiot")
            {
                zhuiotCSV.Clear();

            }
            else if (pdfType == "batim")
            {
                batimCSV.Clear();

            }
        }
        public List<string> getCSVFiles(string pdfType)
        {
            List<string> csvfiles = null;
            if (pdfType == "zhuiot")
            {
                csvfiles = new List<string>(zhuiotCSV);
            }
            else if(pdfType == "batim")
            {
                csvfiles = new List<string>(batimCSV);
            }
            return csvfiles;
         }

        public void convertPDF2CSV()
        {
            int excelRow = 1;
            string tempDir = PDFfolder + "\\CSV\\";
            if (!System.IO.Directory.Exists(tempDir))
            {
                System.IO.Directory.CreateDirectory(tempDir);
            }
            List<string> PDFfiles = excelOperation.getPdfFileNames();
            foreach (string sss in PDFfiles)
            {
                try
                {
                    excelRow++;
                    string NesachType = "";
                    string Gush = "";
                    string Helka = "";
                    List<string> CSVPages = new List<string>();
                    int num;
                    string ssslower = sss;
                    var regex = new Regex(@"[A-Z]", RegexOptions.IgnoreCase);
                    ssslower = regex.Replace(ssslower, m => m.ToString().ToLower());

                    string fullPath = Path.Combine(PDFfolder, ssslower);
                    StringBuilder text = new StringBuilder();
                    PdfReader pdfReader = new PdfReader(fullPath);
                    PdfDocument pdfDoc = new PdfDocument(pdfReader);
                    num = pdfDoc.GetNumberOfPages();
                    for (int page = 1; page <= num; page++)
                    {

                        if (DebugMode)
                        {
                            excelOperation.setActiveSheet(ClassExcelOperations.Sheets.BatimError);
                            int row = excelOperation.getBatimErrorLine();
                            int newrow = row + 1;
                            excelOperation.PutValueInSheetRowColumn(ClassExcelOperations.Sheets.BatimError, row, 1, page.ToString());
                            excelOperation.setSheetCellWrapText(ClassExcelOperations.Sheets.BatimError, false, 6, row, 1);
                        }

                        string pageContent = "";
                        ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();
                        try
                        {
                            pageContent = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                        }
                        catch (System.NullReferenceException e)
                        {
                            //                     e.ToString();
                        }
                        CSVPages.Add(pageContent);
                        //                    CSVPages.Add(System.Environment.NewLine);
                    }
                    pdfDoc.Close();
                    pdfReader.Close();

                    string CSVFile = ssslower.Replace("pdf", "csv");
                    string fulCSVName = tempDir + CSVFile;
                    TextWriter tw = new StreamWriter(fulCSVName);
                    foreach (string s in CSVPages)
                    {
                        //                    string s1 = ClassUtils.ConvertToHebrew(s);
                        //                    tw.WriteLine(s1);
                        string[] s0 = s.Split('\n');
                        for (int i = 0; i < s0.Length; i++)
                        {
                            string[] s1 = s0[i].Split(' ');
                            List<string> list = ClassUtils.removeAllBlancs(s1);
                            if (list.Count == 0) continue;
                            //                        List<string> list = new List<string>(s1);
                            List<string> converted = ClassUtils.ConvertToHebrew0(list);
                            bool realNesach = false;
                            switch (i)
                            {
                                case 0:
                                    realNesach = ClassUtils.isItARealNesach(converted, "תאריך");
                                    break;
                                case 2:
                                    realNesach = ClassUtils.isItARealNesach(converted, "שעה:");
                                    break;
                                case 3:
                                    realNesach = ClassUtils.isItARealNesach(converted, "נסח");
                                    break;
                                case 4:
                                    realNesach = ClassUtils.isItARealNesach(converted, "מקרקעין:");
                                    break;
                                case 5:
                                    realNesach = ClassUtils.isItARealNesach(converted, "מפנקס");
                                    break;
                                default:
                                    realNesach = true;
                                    break;
                            }
                            if ( !realNesach)
                            {
                                throw new Exception("נסח שגוי");
                            }
                            if (NesachType == "")
                            {
                                if (ClassUtils.isArrayIncludString(converted, "העתק") > -1)
                                {
                                    if (ClassUtils.isArrayIncludString(converted, "הזכויות") > -1)
                                    {
                                        NesachType = "זכויות";
                                        zhuiotCSV.Add(fulCSVName);
                                    }
                                    else if (ClassUtils.isArrayIncludString(converted, "משותפים") > -1)
                                    {
                                        NesachType = "בתים משותפים";
                                        batimCSV.Add(fulCSVName);
                                    }
                                }
                            }
                            if (Gush == "" || Helka == "")
                            {
                                if (ClassUtils.isArrayIncludString(converted, "גוש") > -1)
                                {
                                   Gush = converted[converted.Count - 2];
                                   Helka = converted[converted.Count - 4];
                                   if (ClassUtils.isArrayIncludString(converted, "תת") > -1 && ClassUtils.isArrayIncludString(converted, "חלקה:") > -1)
                                    {
                                        batimCSV.RemoveAt(batimCSV.Count - 1);
                                        throw new Exception("נסח תת חלקה- לא נתמך");
                                    }
                                }
                            }
                            tw.WriteLine(string.Join(" ", converted));
                        }
                        tw.WriteLine('\n');
                    }
                    tw.Close();
                    excelOperation.putParamsToTable(excelRow, NesachType, Gush, Helka);

                }
                catch (Exception e)
                {
                    string ssss = e.Message.ToString();
                    excelOperation.putParamsToTable(excelRow, ssss, "", "");
                }
            }


        }

    }
}
