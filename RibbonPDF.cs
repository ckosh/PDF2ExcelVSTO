using GrabNadlanLicense;
using PDF2ExcelVsto.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Windows.Input;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new RibbonPDF();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PDF2ExcelVsto
{
    [ComVisible(true)]
    public class RibbonPDF : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private bool expired = false;
        ClassFilesHandle fileHandler;
        ClassExcelOperations excelOperations = null;
        //        ClassTabooManager tabooManager;
        ClasszhuiotManager zuiotManager;
        ClassbatimManager0 batimManager;
        ClassJoinSplitManager joinSplitManager;

        bool pressed = false;

        public RibbonPDF()
         {
            //  check expiting data 

            Vba2VSTO vba2VSTO = new Vba2VSTO();

            //excelOperations = new ClassExcelOperations(false);
            //fileHandler = new ClassFilesHandle(excelOperations, false,"");
            //zuiotManager = new ClasszhuiotManager(fileHandler, excelOperations);
            //batimManager = new ClassbatimManager0(fileHandler, excelOperations);
            //joinSplitManager = new ClassJoinSplitManager(fileHandler, excelOperations, batimManager, zuiotManager);
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PDF2ExcelVsto.RibbonPDF.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void selectPDF_Click(Office.IRibbonControl control)
        {
            if (expired)
            {
                printMessage();
                return;
            }
            if (excelOperations != null)
            {
                excelOperations.deleteAllPages();
            }
            excelOperations = new ClassExcelOperations(false);
            fileHandler = new ClassFilesHandle(excelOperations, false, "");
            zuiotManager = new ClasszhuiotManager(fileHandler, excelOperations);
            batimManager = new ClassbatimManager0(fileHandler, excelOperations);
            joinSplitManager = new ClassJoinSplitManager(fileHandler, excelOperations, batimManager, zuiotManager);

            excelOperations.deleteSheet(ClassExcelOperations.Sheets.PDFfiles);
            excelOperations.createSheet(ClassExcelOperations.Sheets.PDFfiles, "נסחים", Color.Black);
            fileHandler.clearCSVFiles("batim");
            fileHandler.clearCSVFiles("zhuiot");
            fileHandler.collectPDFFiles();
            fileHandler.convertPDF2CSV();
        }
        public void buildTabu_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            if (pressed) return;
            pressed = true;
            //            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            batimManager.convertBatimtoExcel();
            zuiotManager.convertZhuiottoExcel();
//            Mouse.OverrideCursor = System.Windows.Input.Cursors.Arrow;
            pressed = false;

        }
        public void createOwners_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Owner);
            zuiotManager.CreateOwnersTable();
        }
        public void createProperty_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Property);
            zuiotManager.CreatePropertyTables();
        }
        public void createLeasing_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Leasing);
            zuiotManager.CreateLeasingTables();
        }
        public void createBatimLeasing_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimLeasing);
            batimManager.CreateBatimLeasing();
        }
        public void createBatimMortgage_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimMortgage);
            batimManager.CreateBatimMortgage();
        }

        public void createBatimAttachments_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimAttachments);
            batimManager.createBatimAttachments();
        }

        public void createBatimRemarks_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimRemarks);
            batimManager.CreateBatimRemarksTables();
        }

        public void createMortGage_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Mortgage);
            zuiotManager.CreateMortGageTables();
        }
        public void createRemark_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.Remark);
            zuiotManager.CreateRemarksTables();
        }
        public void batimProperty_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimProperty);
            batimManager.CreatePropertyTable();
        }
        public void batimOwners_Click(Office.IRibbonControl control)
        {
            if (expired) return;
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.BatimOwners);
            batimManager.CreateBatimOwnTable();
        }
        public void JoinSplit_Click(Office.IRibbonControl control)
        {
            excelOperations.deleteSheet(ClassExcelOperations.Sheets.JoinSplit);
            joinSplitManager.CreateJoinSplitTable();
         }
        public void setup_Click(Office.IRibbonControl control)
        {
            FormOptions frm = new FormOptions();
            frm.ShowDialog();
        }
        public System.Drawing.Image get_batimproperty_icon(Office.IRibbonControl control)
        {
            return Resources.apartments;
        }
        public System.Drawing.Image get_Attachments_icon(Office.IRibbonControl control)
        {
            return Resources.Attachments;
        }

        public System.Drawing.Image get_leasing_icon(Office.IRibbonControl control)
        {
            return Resources.rent;
        }
        public System.Drawing.Image get_option_icon(Office.IRibbonControl control)
        {
            return Resources.setup;
        }
        public System.Drawing.Image get_Mortgage_icon(Office.IRibbonControl control)
        {
            return Resources.property_mortgage;
        }
        public System.Drawing.Image get_PDF_icon(Office.IRibbonControl control)
        {
            return Resources.icons8_pdf_80;
        }
        public System.Drawing.Image get_mincer_icon(Office.IRibbonControl control)
        {
            return Resources.mincer;
        }
        public System.Drawing.Image get_table_icon(Office.IRibbonControl control)
        {
            return Resources.landlord;
        }

        public System.Drawing.Image get_property_icon(Office.IRibbonControl control)
        {
            return Resources.land;
        }
        public System.Drawing.Image get_Remark_icon(Office.IRibbonControl control)
        {
            return Resources.remarks;
        }
        public System.Drawing.Image get_JoinSplit_icon(Office.IRibbonControl control)
        {
            return Resources.scissors;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void printMessage()
        {
            string Massage = "רשיון שימוש התכנה פג !!!!!";
            MessageBox.Show(Massage);
        }
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
