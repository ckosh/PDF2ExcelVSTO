

using GrabNadlanLicense;
using OpenPop.Mime;
using OpenPop.Pop3;
using PDF2ExcelVsto.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static GrabNadlanLicense.ClassMongoDBPDF;

namespace PDF2ExcelVsto
{
    public class ClassBatch
    {
        public string customerMail;
        public int NumberOfPDFFiles;
        public string TempFolder;

        public Pop3Client objPop3Client;
        public string host;
        public string user;
        public string password;
        public int port;
        bool useSsl;
        string[] PdfFileNames;
        bool DebugMode;
        bool batchMode = false;
        int customertype;

        static Dictionary<string, DateTime> messageTime = new Dictionary<string, DateTime>();

        public ClassBatch(bool bmode)
        {
            objPop3Client = new Pop3Client();
            host = "mail.grabnadlan.co.il";
            user = "tabu2excel@grabnadlan.co.il";
            password = "L#(Bcw^Wi7{7";
            port = 110;
            useSsl = false;
            //           password = "fDb^PAisuvv0"; // users
            //            password = "7{7iW^wcB(#L";
            //            password = "zmuTA$3P+sla";
            batchMode = bmode;
            customertype = 0;
            messageTime.Clear();
//            TempFolder = Resources.TempFolder;
        }
        public void connectToPop()
        {
            try
            {
                objPop3Client.Connect(host, port, useSsl);
                objPop3Client.Authenticate(user, password);
            }
            catch (Exception e)
            {
            }
        }

        public  int getNumberOfmessagesGMAIL()
        {
            Pop3Client client = new Pop3Client();
            client.Connect("pop.gmail.com", 995, true);
            client.Authenticate("recent:koshizky.chaim@gmail.com", "zdr18b#!64bKlde");
            int messageCount = client.GetMessageCount();
            client.Disconnect();
            return messageCount;
        }
        public int getNumberOfmessages()
        {
            objPop3Client.Connect(host, port, useSsl);
            objPop3Client.Authenticate(user, password);
            int j = objPop3Client.GetMessageCount();
            objPop3Client.Disconnect();
            return j;
        }
        public int ActMailJob(int mailCount, string tempFolder, bool debugMode)
        {
            string resultExcelFile;
            TempFolder = tempFolder;
            DebugMode = debugMode;
            MessageModel message = new MessageModel();
            if ( DebugMode)
            {
                message = GetEmailContentGmail(mailCount);
            }
            else
            {
                message = GetEmailContent(mailCount);
            }

            if ( !removeObstical(message) ) //  over 30 minutes in que - remove
            {
                deleteMail(mailCount);
                string sub = "ארעה שגיעה בניתוח הנסחים - שלח את הקבצים שנית ל - grabnadlan@gmail.com";
                sendMail(message.FromID.ToLower(), sub, null, sub);
                return 0;

            }



            if (message.Subject == "Registration")
            {
                var text = message.Body;
                bool bret;
                string[] stringSeparators = new string[] { "\n" };
                string[] lines = text.Split(stringSeparators, StringSplitOptions.None);
                string mail = getRequestedEmail("E_Mail:", lines).ToLower();
                mail = mail.Trim();
                bret = CheckUserRegistered(mail);
                if ( bret)
                {
                    string sub = "שגיאת רישום - משתמש כבר רשום ";
                    sendMail(mail, sub, null, sub);
                }
                else
                {
                    bool bret0 = ClassUtils.verifyMailAddress(mail);
                    if (bret0)
                    {
                        PDFCustomers cust = new PDFCustomers();
                        cust.Mail = mail;
                        cust.Office_Name = getRequestedEmail("Office:", lines).Trim();
                        cust.Phone = getRequestedEmail("Phone:", lines).Trim();
                        cust.customerStatus = 0;
                        cust.User_Name = getRequestedEmail("Name:", lines).Trim();

                        bool ret = false;
                        Vba2VSTO vba2VSTO = new Vba2VSTO();
                        ret = vba2VSTO.SavePDFCustomerToDB(cust);
                        if (ret)
                        {
                            customerMail = message.FromID;
                            string sub = "אישור רישום לשרות הסבת נסחי טאבו לאקסל";
                            ConfirmationMail(mail, sub);
                            deleteAllFilesFromDirectory(TempFolder);
                            string body0 = cust.Mail + Environment.NewLine;
                            body0 = body0 + cust.Office_Name + Environment.NewLine;
                            body0 = body0 + cust.Phone;

                            ConfirmationMailSimple("grabnadlan@gmail.com", sub, body0);

                        }
                    }
                    else
                    {
                        ConfirmationMailSimple("grabnadlan@gmail.com", "שגיאת כתובת מייל", mail);
                    }
                 }
                deleteMail(mailCount);
                return 0;
            }
            customerMail = message.FromID.ToLower();
            customertype = GetCustomerType(customerMail);
            if (!CheckUserRegistered(customerMail))
            {
                ///  return mail non registered
                ///  delete mail 
                string sub = "משתמש אינו רשום במערכת";
                sendMail(customerMail, sub, null, sub);
                deleteMail(mailCount);
                deleteAllFilesFromDirectory(TempFolder);
                return 0;
            }
            if (!CheckUserPermission(customerMail))
            {
                ///  return mail non registered
                ///  delete mail 
                string sub = "משתמש אינו מאושר לשימוש";
                sendMail(customerMail, sub, null, sub);
                deleteMail(mailCount);
                deleteAllFilesFromDirectory(TempFolder);
                return 0;
            }
            NumberOfPDFFiles = 0;
            if (message.Attachment != null)
            {
                NumberOfPDFFiles = message.Attachment.Count;
            }
            if (NumberOfPDFFiles == 0)
            {
                string sub = "!מייל לא מכיל נסחים";
                sendMail(customerMail, sub, null, sub);
                deleteMail(mailCount);
                return 0;
            }
            //if (!AllPDFFiles(PdfFileNames))
            //{
            //    string sub = "לא כל הקבצים - PDF ";
            //    sendMail(customerMail, sub, null, sub);
            //    deleteMail(mailCount);
            //    return 0;
            //}
            ClassBatchProcessFiles processFiles = new ClassBatchProcessFiles(PdfFileNames,DebugMode, TempFolder, batchMode);
            resultExcelFile = processFiles.convert();
            string body = NumberOfPDFFiles.ToString() + " מספר נסחים ";
            sendMail(customerMail, "תוצאות הסבת נסחי טאבו", resultExcelFile, body);
            deleteAllFilesFromDirectory(TempFolder);
            deleteAllFilesFromDirectory(TempFolder+"\\CSV");
            deleteMail(mailCount);
            if (customertype < 99)
            {
                savecBillingData(customerMail, NumberOfPDFFiles, resultExcelFile);
            }            
            return NumberOfPDFFiles;
        }
        public MessageModel GetEmailContentGmail(int intMessageNumber)
        {
            
            Pop3Client client = new Pop3Client();
            client.Connect("pop.gmail.com", 995, true);
            client.Authenticate("recent:koshizky.chaim@gmail.com", "zdr18b#!64bKlde");
            MessageModel message = new MessageModel();
            OpenPop.Mime.Message msg = client.GetMessage(intMessageNumber);
            MessagePart plainTextPart = null, HTMLTextPart = null;
            message.MessageID = msg.Headers.MessageId == null ? "" : msg.Headers.MessageId.Trim();
            message.FromID = msg.Headers.From.Address.Trim();
            message.FromName = msg.Headers.From.DisplayName.Trim();
            message.Subject = msg.Headers.Subject.Trim();
            plainTextPart = msg.FindFirstPlainTextVersion();
            message.Body = (plainTextPart == null ? "" : plainTextPart.GetBodyAsText().Trim());
            HTMLTextPart = msg.FindFirstHtmlVersion();
            message.Html = (HTMLTextPart == null ? "" : HTMLTextPart.GetBodyAsText().Trim());

            List<MessagePart> attachment = msg.FindAllAttachments();
            if (attachment.Count > 0)
            {
                PdfFileNames = new string[attachment.Count];
                for (int j = 0; j < attachment.Count; j++)
                {
                    byte[] content = attachment[j].Body;
                    string[] stringParts = attachment[j].FileName.Split(new char[] { '.' });
                    string stringType = stringParts[1];
                    BinaryWriter Writer = null;
                    string Name = TempFolder + "\\" + attachment[j].FileName;
                    PdfFileNames[j] = attachment[j].FileName;
                    try
                    {
                        // Create a new stream to write to the file
                        Writer = new BinaryWriter(File.OpenWrite(Name));

                        // Writer raw data                
                        Writer.Write(content);
                        Writer.Flush();
                        Writer.Close();
                    }
                    catch
                    {
                    }

                }
                message.FileName = attachment[0].FileName.Trim();
                message.Attachment = attachment;

            }
            client.Disconnect();
            return message;
        }
        public MessageModel GetEmailContent(int intMessageNumber)
        {
            objPop3Client.Connect(host, port, useSsl);
            objPop3Client.Authenticate(user, password);

            MessageModel message = new MessageModel();
            OpenPop.Mime.Message objMessage;
            MessagePart plainTextPart = null, HTMLTextPart = null;
            objMessage = objPop3Client.GetMessage(intMessageNumber);
            message.MessageID = objMessage.Headers.MessageId == null ? "" : objMessage.Headers.MessageId.Trim();
            message.FromID = objMessage.Headers.From.Address.Trim();
            message.FromName = objMessage.Headers.From.DisplayName.Trim();
            message.Subject = objMessage.Headers.Subject.Trim();

            plainTextPart = objMessage.FindFirstPlainTextVersion();
            message.Body = (plainTextPart == null ? "" : plainTextPart.GetBodyAsText().Trim());

            HTMLTextPart = objMessage.FindFirstHtmlVersion();
            message.Html = (HTMLTextPart == null ? "" : HTMLTextPart.GetBodyAsText().Trim());


            List<MessagePart> attachment = objMessage.FindAllAttachments();
            if (attachment.Count > 0)
            {
                PdfFileNames = new string[attachment.Count];
                for (int j = 0; j < attachment.Count; j++)
                {
                    byte[] content = attachment[j].Body;
                    string[] stringParts = attachment[j].FileName.Split(new char[] { '.' });
                    string stringType = stringParts[1];
                    BinaryWriter Writer = null;
                    string Name = TempFolder + "\\" + attachment[j].FileName;
                    PdfFileNames[j] = attachment[j].FileName;
                    try
                    {
                        // Create a new stream to write to the file
                        Writer = new BinaryWriter(File.OpenWrite(Name));

                        // Writer raw data                
                        Writer.Write(content);
                        Writer.Flush();
                        Writer.Close();
                    }
                    catch
                    {
                    }

                }
                message.FileName = attachment[0].FileName.Trim();
                message.Attachment = attachment;
                
            }
            objPop3Client.Disconnect();
            return message;
        }
        public bool CheckUserRegistered(string userMail)
        {
            bool ret = false ;
            Vba2VSTO vba2VSTO = new Vba2VSTO();
            ret = vba2VSTO.GetRegisteredCustomerFromMongoDBPDF(userMail);

            return ret;
        }

        public int GetCustomerType(string customerMail)
        {
            int iret = 0;
            Vba2VSTO vba2VSTO = new Vba2VSTO();
            iret = vba2VSTO.GetDBPDFCustomerType(customerMail);
            return iret;
        }
        public bool CheckUserPermission(string userMail)
        {
            bool ret = false;
            Vba2VSTO vba2VSTO = new Vba2VSTO();
            ret = vba2VSTO.GetRegisteredPermissionFromMongoDBPDF(userMail);
            return ret;
        }

        public void savecBillingData( string e_mail, int numberOfFiles, string excelFileName)
        {
            Vba2VSTO vba2VSTO = new Vba2VSTO();
            vba2VSTO.SaveBillingData(e_mail, numberOfFiles, excelFileName);
        }
        public void sendMail(string to, string subject, string filePath, string Body)
        {
            MailMessage message = new MailMessage(user, to);
            message.Subject = subject;
            message.Body = Body;

            System.Net.Mail.Attachment data;
            if ( filePath != null)
            {
                data = new System.Net.Mail.Attachment(filePath); //, MediaTypeNames.Application.Octet
                message.Attachments.Add(data);
            }
            if (DebugMode)
            {
                var credentials = new NetworkCredential("koshizky.chaim@gmail.com", "zdr18b#!64bKlde");
                message.From = new MailAddress("koshizky.chaim@gmail.com");
                var client = new SmtpClient()
                {
                    Port = 587,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Host = "smtp.gmail.com",
                    EnableSsl = true,
                    Credentials = credentials
                };
                try
                {
                    client.Send(message);
                }
                catch (Exception e)
                {

                }
                message.Dispose();
                client.Dispose();
            }
            else
            {
                SmtpClient client = new SmtpClient(host);
                client.Credentials = new System.Net.NetworkCredential("tabu2excel@grabnadlan.co.il", "L#(Bcw^Wi7{7");
                try
                {
                    client.Send(message);
                }
                catch (Exception e)
                {
                }               
                message.Dispose();
                client.Dispose();
            }
        }

        //public void ForwardMail( MailMessage.ForwardAsAttachment msg)
        //{
        //    SmtpClient client = new SmtpClient(host);
        //    client.Credentials = new System.Net.NetworkCredential("tabu2excel@grabnadlan.co.il", "L#(Bcw^Wi7{7");
        //    MailMessage message = new MailMessage(user, "tabubackup@grabnadlan.co.il");

        //}
        public void ConfirmationMail(string to, string subject)
        {
            MailMessage message = new MailMessage(user, to);
            message.Subject = subject;
            message.IsBodyHtml = true;
            ResourceManager rm = Resources.ResourceManager;
            Bitmap myImage = (Bitmap)rm.GetObject("registrationReply");
            var filePath = TempFolder + "\\registrationReply.png";
            myImage.Save(filePath);
            var inlineLogo = new LinkedResource(filePath, "image/png");
            inlineLogo.ContentId = Guid.NewGuid().ToString();
            string body = string.Format(@"<img src= ""cid:{0}"" />", inlineLogo.ContentId);
            var view = AlternateView.CreateAlternateViewFromString(body, null, "text/html");
            view.LinkedResources.Add(inlineLogo);
            message.AlternateViews.Add(view);
            if (DebugMode)
            {
                var credentials = new NetworkCredential("koshizky.chaim@gmail.com", "zdr18b#!64bKlde");
                var mail = new MailMessage()
                {
                    From = new MailAddress("koshizky.chaim@gmail.com"),
                    Subject = subject,
                    IsBodyHtml = true,
                    Body = body
                };
                mail.AlternateViews.Add(view);
                mail.To.Add(new MailAddress("koshizky.chaim@gmail.com"));

                var client = new SmtpClient()
                {
                    Port = 587,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Host = "smtp.gmail.com",
                    EnableSsl = true,
                    Credentials = credentials
                };
                try
                {
                    client.Send(mail);
                }
                catch (Exception e)
                {

                }               
                message.Dispose();
                client.Dispose();              
            }
            else
            {
                SmtpClient client = new SmtpClient(host);
                client.Credentials = new System.Net.NetworkCredential("tabu2excel@grabnadlan.co.il", "L#(Bcw^Wi7{7");
                try
                {
                    client.Send(message);
                }
                catch (Exception e)
                {

                }
                message.Dispose();
                client.Dispose();
            }
        }
        public void ConfirmationMailSimple(string to, string subject, string body)
        {
            MailMessage message = new MailMessage(user, to);
            message.Subject = subject;
            message.IsBodyHtml = false;
            message.Body = body;
            SmtpClient client = new SmtpClient(host);
            client.Credentials = new System.Net.NetworkCredential("tabu2excel@grabnadlan.co.il", "L#(Bcw^Wi7{7");
            try
            {
                client.Send(message);
            }
            catch (Exception e)
            {

            }
            message.Dispose();
            client.Dispose();
        }
        public void deleteMail(int j)
        {
            objPop3Client.Connect(host, port, useSsl);
            objPop3Client.Authenticate(user, password);
            objPop3Client.DeleteMessage(j);
            objPop3Client.Disconnect();
        }

        public void deleteAllFilesFromDirectory(string directory)
        {
            DirectoryInfo di = new DirectoryInfo(directory);
            FileInfo[] files = di.GetFiles();
            foreach (FileInfo file in files)
            {
                file.Delete();
            }
        }

       public bool AllPDFFiles(string[] ss)
        {
            bool bret = true;
            for ( int i = 0; i < ss.Length; i++)
            {
                string ex = ( Path.GetExtension(ss[i])).ToUpper();
                if (ex != ".PDF")
                {
                    bret = false;
                    break;
                }
            }
            return bret;
        }

        public string getRequestedEmail(string param, string[] lines)
        {
            string reqString = "";
            int paramLenth = param.Length ;
            foreach (string s in lines)
            {
                int pos = s.IndexOf(param, 0);
                if (pos > -1)
                {
                    reqString = s.Substring(paramLenth);
                    reqString = reqString.Replace("\r", string.Empty);
                    break;
                }
            }
            return reqString;
        }
        public class MessageModel
        {
            public string MessageID;
            public string FromID;
            public string FromName;
            public string Subject;
            public string Body;
            public string Html;
            public string FileName;
            public List<MessagePart> Attachment;
        }

        private bool  removeObstical(MessageModel msg)
        {
            bool bret = true;
            string mesID = msg.MessageID;
            DateTime dateTime = DateTime.Now;

            string key;
            DateTime tim;


            if (messageTime.Count == 0)
            {
                messageTime.Add(msg.MessageID, dateTime);
                return bret; ;
            }
            else
            {
                var first = messageTime.First();
                key = first.Key;
                tim = first.Value;
            }

            if ( key == mesID)
            {
                var diff = DateTime.Now - tim;
                if ( diff.Minutes > 30)
                {
                    messageTime.Clear();
                    bret = false;
                    return bret;
                }
            }
            return bret;

        }
    }
}
