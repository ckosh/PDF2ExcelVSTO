using log4net;
using PDF2ExcelVsto.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Resources;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Resources;

namespace PDF2ExcelVsto
{
    public partial class FormOptions : Form
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public FormOptions()
        {
            InitializeComponent();
        }

        private void buttonBatchMode_Click(object sender, EventArgs e)
        {
            buttonBatchMode.Enabled = false;
            ClassBatch batchClass = new ClassBatch(false);
            //            batchClass.connectToPop();
            //int interval = Convert.ToInt32(Resources.SampleSeconds) * 1000;
            string tempFolder = textBoxTempFolder.Text;
            int interval = Convert.ToInt32(textBoxDelay.Text) * 1000;
            bool debugMode = checkBoxDebugMode.Checked;
            Log.Info("application started");
            while (true)
            {
                try
                {
                    int mailcount;
                    if (debugMode)
                    {
                        mailcount = batchClass.getNumberOfmessagesGMAIL();
                    }
                    else
                    {
                        mailcount = batchClass.getNumberOfmessages();
                    }
                    if (mailcount > 0)
                    {
                        int firstMail = 1;
                        int count = batchClass.ActMailJob(firstMail, tempFolder, debugMode);
                        count = count + Int32.Parse(labelsent.Text);
                        labelsent.Text = count.ToString();
                    }
                }
                catch (Exception ee)
                {
                    Log.Error("exception encountered");
                    Console.WriteLine(ee.ToString());
                }
                Thread.Sleep(interval);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string host = "mail.grabnadlan.co.il";
            string user = "tabu2excel@grabnadlan.co.il";
            var client = new SmtpClient(host);
            client.Credentials = new System.Net.NetworkCredential(user, "L#(Bcw^Wi7{7");
            
            string to = "chaim.koshizky@gmail.com";
            MailMessage message = new MailMessage(user, to);
            message.Subject = "testc image";
            message.IsBodyHtml = true;
            ResourceManager rm = Resources.ResourceManager;
            Bitmap myImage = (Bitmap)rm.GetObject("registrationReply");
            myImage.Save(@"F:\Temp\registrationReply.png");


            var filePath = @"F:\Temp\registrationReply.png";
            //            var filePath = System.Reflection.Assembly.GetExecutingAssembly()
            //                   .Location + @"C:\Users\user\source\repos\PDF2ExcelVsto\Resources\setup.png";
            var inlineLogo = new LinkedResource(filePath, "image/png");
            inlineLogo.ContentId = Guid.NewGuid().ToString();
            string body = string.Format(@"<img src= ""cid:{0}"" />", inlineLogo.ContentId);
            var view = AlternateView.CreateAlternateViewFromString(body, null, "text/html");
            view.LinkedResources.Add(inlineLogo);
            message.AlternateViews.Add(view);
            client.Send(message);
            message.Dispose();
            client.Dispose();

        }

        private void FormOptions_Load(object sender, EventArgs e)
        {

        }

        private void textBoxDelay_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
